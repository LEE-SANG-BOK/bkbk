from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.table import Table
from docx.text.paragraph import Paragraph

from eia_gen.spec.models import SpecBundle


@dataclass(frozen=True)
class TemplateCheckReport:
    template_path: str
    spec_dir: str
    duplicate_anchors_in_spec: list[str]
    duplicate_anchors_in_template: list[dict[str, int | str]]
    missing_anchors: list[str]
    ignored_missing_anchors: list[str]
    table_anchors_missing_table: list[str]
    table_anchors_column_mismatch: list[dict[str, int | str]]

    def has_errors(self) -> bool:
        return bool(
            self.duplicate_anchors_in_spec
            or self.duplicate_anchors_in_template
            or self.missing_anchors
            or self.table_anchors_missing_table
            or self.table_anchors_column_mismatch
        )

    def to_dict(self) -> dict:
        return {
            "template_path": self.template_path,
            "spec_dir": self.spec_dir,
            "duplicate_anchors_in_spec": self.duplicate_anchors_in_spec,
            "duplicate_anchors_in_template": self.duplicate_anchors_in_template,
            "missing_anchors": self.missing_anchors,
            "ignored_missing_anchors": self.ignored_missing_anchors,
            "table_anchors_missing_table": self.table_anchors_missing_table,
            "table_anchors_column_mismatch": self.table_anchors_column_mismatch,
        }


def _find_table_after_paragraph(paragraph: Paragraph) -> Table | None:
    nxt = paragraph._p.getnext()
    if isinstance(nxt, CT_Tbl):
        return Table(nxt, paragraph._parent)
    return None


def _iter_paragraphs_in_table(table: Table) -> list[Paragraph]:
    out: list[Paragraph] = []
    for row in table.rows:
        for cell in row.cells:
            out.extend(_iter_paragraphs_in_cell(cell))
    return out


def _iter_paragraphs_in_cell(cell) -> list[Paragraph]:
    out: list[Paragraph] = []
    out.extend(list(getattr(cell, "paragraphs", []) or []))
    for t in list(getattr(cell, "tables", []) or []):
        out.extend(_iter_paragraphs_in_table(t))
    return out


def _iter_all_paragraphs(doc: Document) -> list[Paragraph]:
    out: list[Paragraph] = []
    out.extend(list(getattr(doc, "paragraphs", []) or []))
    for t in list(getattr(doc, "tables", []) or []):
        out.extend(_iter_paragraphs_in_table(t))
    return out


def _expected_table_headers(spec: SpecBundle, table_id: str) -> list[str]:
    tables_by_id = {t.id: t for t in spec.tables.tables}
    t_spec = tables_by_id.get(table_id)
    if not t_spec:
        return ["(unknown table spec)"]

    # Mirror eia_gen.services.tables.spec_tables.build_table() header logic
    if t_spec.mode == "sources_registry" or t_spec.id == "TBL-SOURCE-REGISTER":
        return ["Source ID", "유형", "자료명", "제공기관", "기준일/기간", "파일/링크", "비고"]

    if t_spec.mode == "assembled":
        return ["항목", "지표", "값", "SRC"]

    if t_spec.mode == "static":
        headers = [str(h) for h in (getattr(t_spec, "headers", None) or [])]
        if not headers:
            headers = ["(empty)"]
        include_src = bool(getattr(t_spec, "include_src_column", None))
        if getattr(t_spec, "include_src_column", None) is None:
            include_src = bool(spec.tables.defaults.include_src_column)
        if include_src and (not headers or str(headers[-1]).strip().upper() != "SRC"):
            headers = [*headers, "SRC"]
        return headers

    headers = [c.title for c in t_spec.columns]
    include_src = bool(spec.tables.defaults.include_src_column)
    if getattr(t_spec, "include_src_column", None) is not None:
        include_src = bool(getattr(t_spec, "include_src_column"))
    if include_src:
        headers = [*headers, "SRC"]
    return headers


def check_template_docx(
    *,
    template_path: Path,
    spec: SpecBundle,
    spec_dir: Path,
    ignore_missing_anchor_prefixes: list[str] | None = None,
) -> TemplateCheckReport:
    doc = Document(str(template_path))
    anchor_entries: list[tuple[str, object]] = []
    for a in spec.template_map.anchors:
        anchor = str(getattr(a, "anchor", "") or "").strip()
        if not anchor:
            continue
        anchor_entries.append((anchor, a.insert))

    dup_in_spec: list[str] = []
    seen: set[str] = set()
    for anchor, _ in anchor_entries:
        if anchor in seen:
            dup_in_spec.append(anchor)
        else:
            seen.add(anchor)
    dup_in_spec = sorted(set(dup_in_spec))

    anchor_map = {anchor: insert for anchor, insert in anchor_entries}

    para_map: dict[str, list[Paragraph]] = {}
    dup_in_tpl: list[dict[str, int | str]] = []
    for p in _iter_all_paragraphs(doc):
        t = p.text.strip()
        if not t:
            continue
        para_map.setdefault(t, []).append(p)

    for anchor in sorted(anchor_map.keys()):
        hits = len(para_map.get(anchor, []))
        if hits > 1:
            dup_in_tpl.append({"anchor": anchor, "count": hits})

    prefixes = [p.strip() for p in (ignore_missing_anchor_prefixes or []) if p.strip()]

    missing_anchors: list[str] = []
    ignored_missing_anchors: list[str] = []
    table_anchors_missing_table: list[str] = []
    table_anchors_column_mismatch: list[dict[str, int | str]] = []

    for anchor, insert in anchor_map.items():
        a = anchor.strip()
        if a not in para_map:
            if any(a.startswith(pfx) for pfx in prefixes):
                ignored_missing_anchors.append(anchor)
            else:
                missing_anchors.append(anchor)
            continue

        if insert.type != "table":
            continue

        # For table anchors, duplicates are considered errors; we still check the first hit best-effort.
        p = para_map[a][0]
        table = _find_table_after_paragraph(p)
        if table is None:
            table_anchors_missing_table.append(anchor)
            continue

        expected_headers = _expected_table_headers(spec, insert.id)
        expected_cols = len(expected_headers)
        actual_cols = len(table.rows[0].cells) if table.rows else 0
        if actual_cols < expected_cols:
            table_anchors_column_mismatch.append(
                {
                    "anchor": anchor,
                    "table_id": insert.id,
                    "expected_cols": expected_cols,
                    "actual_cols": actual_cols,
                }
            )

    return TemplateCheckReport(
        template_path=str(template_path),
        spec_dir=str(spec_dir),
        duplicate_anchors_in_spec=dup_in_spec,
        duplicate_anchors_in_template=dup_in_tpl,
        missing_anchors=sorted(missing_anchors),
        ignored_missing_anchors=sorted(ignored_missing_anchors),
        table_anchors_missing_table=sorted(table_anchors_missing_table),
        table_anchors_column_mismatch=table_anchors_column_mismatch,
    )


def scaffold_template_docx(
    *,
    template_path: Path,
    out_path: Path,
    spec: SpecBundle,
    spec_dir: Path,
    min_data_rows: int = 1,
    table_style: str = "Table Grid",
) -> TemplateCheckReport:
    """Ensure template has (at least) one placeholder table right after each table anchor.

    This enables the renderer's in-place table filling, which significantly improves
    layout fidelity (borders/widths/shading/merged cells can be set in the template).
    """
    doc = Document(str(template_path))
    anchor_map = {str(a.anchor or "").strip(): a.insert for a in spec.template_map.anchors if str(a.anchor or "").strip()}
    paras = {p.text.strip(): p for p in _iter_all_paragraphs(doc)}

    # Create placeholder tables where missing.
    for anchor, insert in anchor_map.items():
        if insert.type != "table":
            continue
        p = paras.get(anchor.strip())
        if p is None:
            continue
        if _find_table_after_paragraph(p) is not None:
            continue

        headers = _expected_table_headers(spec, insert.id)
        cols = max(1, len(headers))
        rows = 1 + max(1, int(min_data_rows))

        table = doc.add_table(rows=rows, cols=cols)
        try:
            table.style = table_style
        except Exception:
            pass

        # Header row text (makes manual styling in Word easier).
        for j, h in enumerate(headers):
            table.cell(0, j).text = str(h)

        # One template row for cloning (leave blank placeholders).
        for i in range(1, rows):
            for j in range(cols):
                table.cell(i, j).text = ""

        # Move table right after the anchor paragraph.
        p._p.addnext(table._tbl)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))

    # Re-check scaffolded output.
    return check_template_docx(template_path=out_path, spec=spec, spec_dir=spec_dir)
