from __future__ import annotations

from pathlib import Path
from typing import Any

from eia_gen.models.case import Case


def guess_case_dir_from_derived_dir(derived_dir: Path) -> Path | None:
    """Best-effort guess for case_dir when derived_dir == case_dir/attachments/derived."""
    try:
        if derived_dir.name == "derived" and derived_dir.parent.name == "attachments":
            return derived_dir.parent.parent
    except Exception:
        return None
    return None


def _relative_path_str(path: Path, base_dir: Path | None) -> str:
    if base_dir is None:
        return str(path)
    try:
        return str(path.resolve().relative_to(base_dir.resolve())).replace("\\", "/")
    except Exception:
        return str(path)


def _normalize_tokens(tokens: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for token in tokens:
        s = (token or "").strip()
        if not s:
            continue
        if s in seen:
            continue
        seen.add(s)
        out.append(s)
    return out


def record_derived_evidence(
    case: Case,
    *,
    derived_path: Path,
    related_fig_id: str,
    report_anchor: str,
    src_ids: list[str] | None = None,
    evidence_type: str = "derived_png",
    title: str | None = None,
    note: str | None = None,
    pdf_page: int | None = None,
    pdf_page_source: str | None = None,
    used_in: str | None = None,
    data_origin: str = "MODEL_OUTPUT",
    sensitive: str = "N",
    doc_target: str | None = None,
    case_dir: Path | None = None,
) -> str:
    """Record an engine-produced artifact so it can be exported to source_register.xlsx.

    This intentionally avoids writing back to the input XLSX. The recorded manifest lives in
    `case.model_extra["derived_evidence_manifest"]` during the current run.
    """

    derived_path = Path(derived_path)
    evidence_id = f"EV-DERIVED-{derived_path.stem}"
    src_ids_norm = _normalize_tokens([str(s) for s in (src_ids or [])])

    entry: dict[str, Any] = {
        "evidence_id": evidence_id,
        "evidence_type": evidence_type,
        "title": title or related_fig_id,
        "file_path": _relative_path_str(derived_path, case_dir),
        "related_fig_id": related_fig_id,
        "used_in": used_in or report_anchor or related_fig_id,
        "data_origin": data_origin,
        "sensitive": sensitive,
        "src_ids": ";".join(src_ids_norm),
        "note": note or "",
    }
    if pdf_page is not None:
        entry["pdf_page"] = int(pdf_page)
    if pdf_page_source:
        entry["pdf_page_source"] = str(pdf_page_source)
    if doc_target:
        entry["doc_target"] = doc_target
    if report_anchor:
        entry["report_anchor"] = report_anchor

    manifest = getattr(case, "derived_evidence_manifest", None)
    if not isinstance(manifest, list):
        manifest = []
        setattr(case, "derived_evidence_manifest", manifest)

    # Dedup by evidence_id; update missing fields best-effort.
    for r in manifest:
        if not isinstance(r, dict):
            continue
        if str(r.get("evidence_id") or "").strip() != evidence_id:
            continue
        for k, v in entry.items():
            if k not in r or r.get(k) in {None, "", []}:
                r[k] = v
        return evidence_id

    manifest.append(entry)
    return evidence_id
