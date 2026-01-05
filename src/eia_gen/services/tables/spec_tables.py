from __future__ import annotations

from collections import Counter
from typing import Any

from eia_gen.models.case import Case
from eia_gen.models.sources import SourceRegistry
from eia_gen.services.docx.types import TableData
from eia_gen.services.tables.path import coerce_number, coerce_text, infer_src_path, resolve_path
from eia_gen.spec.models import TableDefaults, TableSpec


def _normalize_ids(ids: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for s in ids:
        s2 = (s or "").strip()
        if not s2:
            continue
        if s2 in seen:
            continue
        seen.add(s2)
        out.append(s2)
    return out


def _row_src_from_columns(row_obj: Any, columns: list[dict[str, Any]]) -> list[str]:
    ids: list[str] = []
    for col in columns:
        value_path = col.get("path")
        if not isinstance(value_path, str):
            continue
        src_path = infer_src_path(value_path)
        if not src_path:
            continue
        src_val = resolve_path(row_obj, src_path)
        if isinstance(src_val, list):
            ids.extend([str(x) for x in src_val])
        elif isinstance(src_val, str):
            ids.append(src_val)
        elif src_val is None:
            continue
        else:
            # dict-like src (rare)
            try:
                ids.extend([str(x) for x in src_val])
            except Exception:
                pass
    return _normalize_ids(ids) or ["S-TBD"]


def build_table(case: Case, sources: SourceRegistry, spec: TableSpec, defaults: TableDefaults) -> TableData:
    def _col_empty(c: dict[str, Any]) -> str:
        return "" if bool(c.get("allow_empty")) else defaults.empty_cell

    # Per-table override (default: global).
    include_src_column = defaults.include_src_column
    if getattr(spec, "include_src_column", None) is not None:
        include_src_column = bool(getattr(spec, "include_src_column"))

    if spec.mode == "sources_registry" or spec.id == "TBL-SOURCE-REGISTER":
        headers = ["Source ID", "유형", "자료명", "제공기관", "기준일/기간", "파일/링크", "비고"]
        rows: list[list[str]] = []
        for s in sources.sources:
            rows.append(
                [
                    s.id,
                    s.type or "",
                    s.title or "",
                    s.publisher or "",
                    s.period or s.date or "",
                    s.file_path or s.file_or_url or "",
                    s.note or s.notes or "",
                ]
            )
        return TableData(caption=spec.caption, headers=headers, rows=rows)

    if spec.mode == "static":
        headers = [str(h) for h in (getattr(spec, "headers", None) or [])]
        rows = [[str(c) for c in r] for r in (getattr(spec, "rows", None) or []) if isinstance(r, list)]

        # Best-effort: ensure at least a 1x1 table to avoid DOCX errors.
        if not headers:
            headers = ["(empty)"]
        if not rows:
            rows = [[defaults.empty_cell for _ in headers]]

        if include_src_column and (not headers or str(headers[-1]).strip().upper() != "SRC"):
            headers = [*headers, "SRC"]
            # Fill SRC column with table-level source_ids (joined) when present.
            src_joined = ", ".join(_normalize_ids([str(x) for x in (getattr(spec, "source_ids", None) or [])]))
            if not src_joined:
                src_joined = "S-TBD"
            rows = [[*(r[: len(headers) - 1]), src_joined] for r in rows]

        src_ids = _normalize_ids([str(x) for x in (getattr(spec, "source_ids", None) or [])])
        return TableData(caption=spec.caption, headers=headers, rows=rows, source_ids=src_ids)

    if spec.mode == "assembled":
        return _build_assembled(case, spec, defaults)

    if not spec.data_path:
        return TableData(caption=spec.caption, headers=["(no data_path)"], rows=[[defaults.empty_cell]])

    data = resolve_path(case, spec.data_path)
    rows_src: list[str] = []

    headers = [c.title for c in spec.columns]
    col_defs = [c.model_dump() for c in spec.columns]
    if include_src_column:
        headers = [*headers, "SRC"]

    rows: list[list[str]] = []

    if isinstance(data, dict):
        # Not used in v1 except zoning breakdown; handled later if needed.
        items = list(data.items())
        for idx, (k, v) in enumerate(items, start=1):
            row_cells: list[str] = []
            for c in col_defs:
                empty = _col_empty(c)
                t = c.get("type")
                if t == "auto_index":
                    row_cells.append(str(idx))
                elif t == "key_name":
                    row_cells.append(str(k))
                elif t == "number":
                    vv = resolve_path({"dict_value": v}, c.get("value_from") or "dict_value.v")
                    uu = resolve_path({"dict_value": v}, c.get("unit_from") or "dict_value.u")
                    row_cells.append(coerce_number(vv, uu, empty))
                elif t == "computed_ratio":
                    num = resolve_path({"dict_value": v}, c.get("numerator_from") or "dict_value.v")
                    den = resolve_path(case, c.get("denominator_path") or "")
                    try:
                        ratio = (float(num) / float(den) * 100.0) if num is not None and den else None
                    except Exception:
                        ratio = None
                    row_cells.append(coerce_number(ratio, "%", empty))
                else:
                    row_cells.append(empty)

            # try get src from dict_value
            src_val = resolve_path({"dict_value": v}, "dict_value.src")
            src_ids = []
            if isinstance(src_val, list):
                src_ids = [str(x) for x in src_val]
            elif isinstance(src_val, str):
                src_ids = [src_val]
            src_ids = _normalize_ids(src_ids) or ["S-TBD"]

            if include_src_column:
                row_cells.append(", ".join(src_ids))
            rows.append(row_cells)
            rows_src.extend(src_ids)
        return TableData(caption=spec.caption, headers=headers, rows=rows, source_ids=_normalize_ids(rows_src))

    if not isinstance(data, list):
        return TableData(caption=spec.caption, headers=headers, rows=[[defaults.empty_cell]], source_ids=["S-TBD"])

    for idx, item in enumerate(data, start=1):
        row_cells: list[str] = []
        for c in col_defs:
            empty = _col_empty(c)
            t = c.get("type")
            if t == "auto_index":
                row_cells.append(str(idx))
                continue

            value_path = c.get("path")
            if not isinstance(value_path, str):
                row_cells.append(empty)
                continue

            if t in {"text", "enum", "date_ym"}:
                val = resolve_path(item, value_path)
                row_cells.append(coerce_text(val, empty))
                continue

            if t == "number":
                val = resolve_path(item, value_path)
                unit = None
                unit_path = c.get("unit_path")
                if isinstance(unit_path, str):
                    unit = resolve_path(item, unit_path)
                row_cells.append(coerce_number(val, coerce_text(unit, ""), empty))
                continue

            if t == "list":
                val = resolve_path(item, value_path)
                if isinstance(val, list):
                    row_cells.append(", ".join([coerce_text(x, "") for x in val if coerce_text(x, "")]))
                else:
                    row_cells.append(coerce_text(val, empty))
                continue

            row_cells.append(empty)

        src_ids = _row_src_from_columns(item, col_defs) if include_src_column else []
        if include_src_column:
            row_cells.append(", ".join(src_ids))
            rows_src.extend(src_ids)
        rows.append(row_cells)

    return TableData(caption=spec.caption, headers=headers, rows=rows, source_ids=_normalize_ids(rows_src))


def _build_assembled(case: Case, spec: TableSpec, defaults: TableDefaults) -> TableData:
    headers = ["항목", "지표", "값", "SRC"]
    rows: list[list[str]] = []
    src_all: list[str] = []

    for row_def in spec.rows_definition:
        label = str(row_def.get("label") or "").strip() or defaults.empty_cell
        items = row_def.get("items") or []
        if not isinstance(items, list):
            continue
        for item_def in items:
            if not isinstance(item_def, dict):
                continue
            name = str(item_def.get("name") or "").strip() or defaults.empty_cell
            path = str(item_def.get("path") or "").strip()
            if not path:
                rows.append([label, name, defaults.empty_cell, "S-TBD"])
                src_all.append("S-TBD")
                continue

            resolved = resolve_path(case, path)
            # 1) list expansion => multiple rows
            if isinstance(resolved, list):
                for idx, obj in enumerate(resolved, start=1):
                    val = resolve_path(obj, str(item_def.get("value") or "v"))
                    unit = resolve_path(obj, str(item_def.get("unit") or "u"))
                    src = resolve_path(obj, str(item_def.get("src") or "src"))
                    src_ids = src if isinstance(src, list) else [src] if isinstance(src, str) else []
                    src_ids = _normalize_ids([str(x) for x in src_ids]) or ["S-TBD"]

                    value_txt = coerce_number(val, coerce_text(unit, ""), defaults.empty_cell)
                    rows.append([label, f"{name}({idx})", value_txt, ", ".join(src_ids)])
                    src_all.extend(src_ids)
                continue

            # 2) single object or dict
            obj = resolved
            if isinstance(obj, dict) and ("v" in obj or "value" in obj):
                val = obj.get("v", obj.get("value"))
                unit = obj.get("u")
                src = obj.get("src")
            else:
                val = resolve_path(obj, str(item_def.get("value") or "v")) if obj is not None else None
                unit = resolve_path(obj, str(item_def.get("unit") or "u")) if obj is not None else None
                src = resolve_path(obj, str(item_def.get("src") or "src")) if obj is not None else None

            src_ids: list[str] = []
            if isinstance(src, list):
                src_ids = [str(x) for x in src]
            elif isinstance(src, str):
                src_ids = [src]
            src_ids = _normalize_ids(src_ids) or ["S-TBD"]

            value_txt = coerce_number(val, coerce_text(unit, ""), defaults.empty_cell)
            rows.append([label, name, value_txt, ", ".join(src_ids)])
            src_all.extend(src_ids)

    return TableData(caption=spec.caption, headers=headers, rows=rows, source_ids=_normalize_ids(src_all))
