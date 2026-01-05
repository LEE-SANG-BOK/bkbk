from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


@dataclass(frozen=True)
class PdfIndexHit:
    kind: str
    page: int  # 1-based
    label: str
    text: str


def _iter_index_files(index_path: Path) -> list[Path]:
    p = index_path
    if p.is_file():
        return [p]

    files: list[Path] = []
    for name in ("combined_index.json", "pass2_hits.json"):
        files.extend(sorted(p.rglob(name)))
    # fallback: any json
    if not files:
        files = sorted(p.rglob("*.json"))
    return files


def load_pdf_index_hits(index_path: str | Path) -> tuple[list[PdfIndexHit], str | None]:
    """Load OCR index hits.

    Supports:
    - `combined_index.json` (extract_pdf_index_twopass.py)
    - `pass2_hits.json` (extract_pdf_index_twopass.py intermediate)
    - A directory containing those files (recursively)

    Returns:
    - hits (deduplicated)
    - pdf_path from the index file, when available
    """
    p = Path(index_path)
    pdf_path: str | None = None

    hits: list[PdfIndexHit] = []

    for f in _iter_index_files(p):
        try:
            obj = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue

        if not pdf_path:
            raw = obj.get("pdf_path")
            if isinstance(raw, str) and raw.strip():
                pdf_path = raw.strip()

        raw_hits = None
        if isinstance(obj.get("pass2"), dict) and isinstance(obj["pass2"].get("hits"), list):
            raw_hits = obj["pass2"]["hits"]
        elif isinstance(obj.get("hits"), list):
            raw_hits = obj.get("hits")

        if not raw_hits:
            continue

        for h in raw_hits:
            if not isinstance(h, dict):
                continue
            kind = str(h.get("kind") or "").strip().lower()
            if kind not in {"chapter", "figure", "table"}:
                continue
            try:
                page = int(h.get("page"))
            except Exception:
                continue
            label = str(h.get("label") or "").strip()
            text = str(h.get("text") or "").strip()
            if not label:
                continue
            hits.append(PdfIndexHit(kind=kind, page=page, label=label, text=text))

    # de-dup, preserve stable order
    seen: set[tuple[str, int, str, str]] = set()
    out: list[PdfIndexHit] = []
    for h in sorted(hits, key=lambda x: (x.page, x.kind, x.label, x.text)):
        key = (h.kind, h.page, h.label, h.text)
        if key in seen:
            continue
        seen.add(key)
        out.append(h)

    return out, pdf_path


def filter_hits(
    hits: Iterable[PdfIndexHit],
    *,
    kinds: set[str] | None = None,
    include_labels: set[str] | None = None,
    include_text_any: list[str] | None = None,
    exclude_text_any: list[str] | None = None,
) -> list[PdfIndexHit]:
    kinds = {k.lower() for k in (kinds or {"figure", "table"})}
    include_labels = {s.strip() for s in (include_labels or set()) if s.strip()} or None
    include_text_any = [s for s in (include_text_any or []) if s.strip()]
    exclude_text_any = [s for s in (exclude_text_any or []) if s.strip()]

    out: list[PdfIndexHit] = []
    for h in hits:
        if h.kind.lower() not in kinds:
            continue
        if include_labels is not None and h.label not in include_labels:
            continue

        text_norm = (h.text or "").lower()
        if include_text_any:
            if not any(k.lower() in text_norm for k in include_text_any):
                continue
        if exclude_text_any:
            if any(k.lower() in text_norm for k in exclude_text_any):
                continue

        out.append(h)

    return out


_ID_SAFE_RE = re.compile(r"[^A-Za-z0-9_-]+")


def _safe_id(s: str) -> str:
    s2 = _ID_SAFE_RE.sub("_", str(s or "")).strip("_")
    return s2 or "X"


def build_pdf_page_data_requests(
    hits: list[PdfIndexHit],
    *,
    pdf_path: str,
    src_id: str,
    req_prefix: str = "REQ-PDF",
    enabled: bool = True,
    run_mode: str = "ONCE",
    dpi: int = 250,
    priority_base: int = 50,
) -> list[dict[str, str | int | bool]]:
    rows: list[dict[str, str | int | bool]] = []

    for i, h in enumerate(hits, start=1):
        kind_tag = "FIG" if h.kind == "figure" else "TBL" if h.kind == "table" else _safe_id(h.kind)
        label = _safe_id(h.label.replace(".", "_"))
        req_id = f"{req_prefix}-{kind_tag}-{label}-p{h.page:03d}"

        title = f"<{kind_tag} {h.label}> {h.text}".strip()
        params = {
            "pdf_path": pdf_path,
            "page": int(h.page),
            "dpi": int(dpi),
            "title": title,
            "data_origin": "LITERATURE",
        }

        rows.append(
            {
                "req_id": req_id,
                "enabled": bool(enabled),
                "priority": int(priority_base) + i,
                "connector": "PDF_PAGE",
                "purpose": "EVIDENCE",
                "src_id": src_id,
                "params_json": json.dumps(params, ensure_ascii=False),
                "output_sheet": "",
                "merge_strategy": "REPLACE_SHEET",
                "upsert_keys": "",
                "run_mode": run_mode,
                "last_run_at": "",
                "last_evidence_ids": "",
                "note": title,
            }
        )

    return rows
