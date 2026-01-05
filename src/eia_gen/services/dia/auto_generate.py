from __future__ import annotations

from typing import Any

from eia_gen.models.case import Case
from eia_gen.models.fields import TextField, normalize_source_ids


_PLACEHOLDER = "【작성자 기입 필요】"


def _tf(text: str, src_ids: list[str] | None = None) -> TextField:
    src = normalize_source_ids(src_ids)
    if not src:
        src = ["S-TBD"]
    return TextField(t=str(text or "").strip(), src=src)


def _any_text(v: Any) -> str:
    if isinstance(v, TextField):
        return (v.t or "").strip()
    if isinstance(v, dict):
        if "t" in v:
            return str(v.get("t") or "").strip()
        if "v" in v:
            return str(v.get("v") or "").strip()
    return str(v or "").strip()


def _collect_src_ids(*values: Any) -> list[str]:
    ids: list[str] = []
    for v in values:
        if isinstance(v, TextField):
            ids.extend(v.src or [])
        elif isinstance(v, dict):
            src = v.get("src")
            if isinstance(src, list):
                ids.extend([str(x).strip() for x in src if str(x).strip()])
            elif isinstance(src, str) and src.strip():
                ids.append(src.strip())
    # Keep S-TBD if present (signals placeholder); de-dup.
    out: list[str] = []
    seen: set[str] = set()
    for s in ids:
        s2 = str(s).strip()
        if not s2 or s2 in seen:
            continue
        seen.add(s2)
        out.append(s2)
    return out


def ensure_dia_standard_forms(case: Case, *, max_rows: int = 8) -> dict[str, Any]:
    """Best-effort DIA standard-form auto generation (appendix feel).

    Current scope:
    - Ensure `disaster.maintenance_ledger` has at least 1 row so DIA templates can render a
      maintenance ledger form even when the user hasn't filled DRR_MAINTENANCE yet.
    - Generated cells use placeholder text to avoid false completeness.
    """

    disaster = getattr(case, "disaster", None)
    if not isinstance(disaster, dict):
        return {"generated": {}, "errors": {}}

    generated: dict[str, Any] = {}
    errors: dict[str, str] = {}

    try:
        mledger = disaster.get("maintenance_ledger")
        if isinstance(mledger, list) and mledger:
            return {"generated": {}, "errors": {}}

        drainage = disaster.get("drainage_facilities")
        candidates: list[tuple[str, list[str]]] = []
        if isinstance(drainage, list):
            for d in drainage:
                if not isinstance(d, dict):
                    continue
                fid = d.get("facility_id")
                typ = d.get("type")
                asset_name = _any_text(fid) or _any_text(typ)
                src_ids = _collect_src_ids(fid, typ, d.get("capacity"), d.get("discharge_to"))
                if asset_name:
                    candidates.append((asset_name, src_ids))

        rows: list[dict[str, Any]] = []
        truncated = False
        if candidates:
            if len(candidates) > max_rows:
                candidates = candidates[:max_rows]
                truncated = True
            for asset_name, src_ids in candidates:
                rows.append(
                    {
                        "asset_id": _tf(asset_name, src_ids),
                        "inspection_cycle": _tf(_PLACEHOLDER, ["S-TBD"]),
                        "inspection_item": _tf(_PLACEHOLDER, ["S-TBD"]),
                        "responsible_role": _tf(_PLACEHOLDER, ["S-TBD"]),
                        "record_format": _tf(_PLACEHOLDER, ["S-TBD"]),
                        "evidence_id_template": _tf(_PLACEHOLDER, ["S-TBD"]),
                    }
                )
        else:
            rows.append(
                {
                    "asset_id": _tf(_PLACEHOLDER, ["S-TBD"]),
                    "inspection_cycle": _tf(_PLACEHOLDER, ["S-TBD"]),
                    "inspection_item": _tf(_PLACEHOLDER, ["S-TBD"]),
                    "responsible_role": _tf(_PLACEHOLDER, ["S-TBD"]),
                    "record_format": _tf(_PLACEHOLDER, ["S-TBD"]),
                    "evidence_id_template": _tf(_PLACEHOLDER, ["S-TBD"]),
                }
            )

        disaster["maintenance_ledger"] = rows
        generated["maintenance_ledger"] = {
            "rows": len(rows),
            "from_drainage_facilities": bool(candidates),
            "truncated": bool(truncated),
        }

        # Expose a marker for writer/QA next-actions (stored as extra field).
        extra = case.model_extra
        if isinstance(extra, dict):
            extra["dia_auto_generated"] = generated
        else:  # pragma: no cover
            setattr(case, "dia_auto_generated", generated)
    except Exception as e:
        errors["maintenance_ledger"] = f"{type(e).__name__}: {e}"

    return {"generated": generated, "errors": errors}
