from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from eia_gen.models.case import Case
from eia_gen.services.docx.types import FigureData
from eia_gen.spec.models import FigureSpec


_SCALE_RE = re.compile(r"^1\s*:\s*(\d+)\s*$")
_FILE_PAGE_RE = re.compile(r"(?i)(?:[#?@]page=)(\d+)\b")
_GEN_METHOD_PAGE_RE = re.compile(r"(?i)(?:^|\b)(?:PDF_PAGE|FROM_PDF_PAGE|PAGE)\s*[:=]\s*(\d+)\b")
_GEN_METHOD_AUTH_RE = re.compile(r"(?i)(?:^|\b)AUTHENTICITY\s*[:=]\s*(REFERENCE|OFFICIAL)\b")


def _normalize_authenticity(*, authenticity: str | None, source_origin: str | None) -> str | None:
    a = (authenticity or "").strip().upper()
    if a in {"OFFICIAL", "REFERENCE"}:
        return a
    so = (source_origin or "").strip().upper()
    if so in {"REFERENCE", "REF"}:
        return "REFERENCE"
    if so == "OFFICIAL":
        return "OFFICIAL"
    return None


def _parse_scale(value: Any) -> int | None:
    if value is None:
        return None
    if isinstance(value, int):
        return value
    s = str(value).strip()
    m = _SCALE_RE.match(s)
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


def is_required(case: Case, spec: FigureSpec) -> bool:
    if spec.required:
        return True
    if spec.required_if is None:
        return False
    cond = spec.required_if
    if not cond.scoping_item_id:
        return False
    # Find scoping item
    for s in case.scoping_matrix:
        if s.item_id == cond.scoping_item_id:
            cat = (s.category.t or "").strip()
            return cat in set(cond.category_in or [])
    return False


def _resolve_existing_path(file_path: str | None, search_dirs: list[Path] | None) -> Path | None:
    fp = (file_path or "").strip()
    if not fp:
        return None
    # Allow optional "#page=N" / "?page=N" fragments in file_path (PDF page hint).
    fp = re.sub(r"(?i)(?:[#?@]page=)\\d+\\b", "", fp)
    p = Path(fp).expanduser()
    if p.is_absolute():
        return p if p.exists() else None
    if search_dirs:
        for base in search_dirs:
            cand = (base / p).expanduser()
            if cand.exists():
                return cand
    return p if p.exists() else None


def _apply_pdf_page_hint(*, file_path: str | None, gen_method: str | None) -> str | None:
    """Fold a '#page=N' style hint into gen_method so materialize can pick the correct PDF page."""
    gm = (gen_method or "").strip() or None
    if gm and _GEN_METHOD_PAGE_RE.search(gm):
        return gm
    fp = (file_path or "").strip()
    if not fp:
        return gm
    m = _FILE_PAGE_RE.search(fp)
    if not m:
        return gm
    try:
        page = max(1, int(m.group(1)))
    except Exception:
        return gm
    hint = f"PDF_PAGE:{page}"
    if gm:
        return f"{gm} {hint}".strip()
    return hint


def _apply_authenticity_hint(
    *,
    authenticity: str | None,
    source_origin: str | None,
    gen_method: str | None,
) -> str | None:
    """Fold FIGURES authenticity into gen_method so downstream renderers can enforce guardrails."""
    gm = (gen_method or "").strip() or None
    if gm and _GEN_METHOD_AUTH_RE.search(gm):
        return gm
    auth = _normalize_authenticity(authenticity=authenticity, source_origin=source_origin)
    if auth == "REFERENCE":
        hint = "AUTHENTICITY:REFERENCE"
    elif auth == "OFFICIAL":
        hint = "AUTHENTICITY:OFFICIAL"
    else:
        return gm
    if gm:
        return f"{gm} {hint}".strip()
    return hint


def _apply_reference_caption_prefix(
    caption: str,
    *,
    authenticity: str | None,
    source_origin: str | None,
) -> str:
    auth = _normalize_authenticity(authenticity=authenticity, source_origin=source_origin)
    if auth != "REFERENCE":
        return caption
    c = (caption or "").strip()
    if not c:
        return "【참고도】"
    if any(token in c for token in ["참고", "REFERENCE", "공식 도면 아님"]):
        return caption
    return f"【참고도】{c}"


def resolve_figure(case: Case, spec: FigureSpec, *, asset_search_dirs: list[Path] | None = None) -> FigureData:
    # Prefer exact asset_id match when present (most deterministic).
    for a in case.assets:
        if a.asset_id == spec.id:
            resolved = _resolve_existing_path(a.file_path, asset_search_dirs)
            source_origin = str(getattr(a, "source_origin", "") or "").strip()
            authenticity = str(getattr(a, "authenticity", "") or "").strip()
            return FigureData(
                file_path=str(resolved) if resolved else None,
                caption=_apply_reference_caption_prefix(
                    a.caption.text_or_placeholder(spec.caption),
                    authenticity=authenticity,
                    source_origin=source_origin,
                ),
                source_ids=a.source_ids or a.caption.src or ["S-TBD"],
                width_mm=getattr(a, "width_mm", None),
                crop=(getattr(a, "crop", None) or "").strip() or None,
                gen_method=_apply_authenticity_hint(
                    authenticity=authenticity,
                    source_origin=source_origin,
                    gen_method=_apply_pdf_page_hint(
                        file_path=a.file_path,
                        gen_method=(getattr(a, "gen_method", None) or "").strip() or None,
                    ),
                ),
            )

    # Otherwise select the first asset matching the type (best-effort).
    candidates = [a for a in case.assets if a.type == spec.asset_type]

    # VP photo selection: FIG-VP-01 -> asset.viewpoint == "VP-01"
    if spec.asset_type == "photo" and spec.id.startswith("FIG-VP-") and candidates:
        vp = spec.id.replace("FIG-", "", 1)
        for a in candidates:
            if (a.viewpoint.t or "").strip() == vp:
                candidates = [a]
                break

    # Prefer an asset whose file actually exists; fall back to the first one otherwise.
    if len(candidates) > 1:
        existing: list[Any] = []
        missing: list[Any] = []
        for a in candidates:
            fp = (a.file_path or "").strip()
            if fp and _resolve_existing_path(fp, asset_search_dirs):
                existing.append(a)
            else:
                missing.append(a)
        candidates = [*existing, *missing]

    asset = candidates[0] if candidates else None
    if asset is None:
        caption = f"【첨부 필요】{spec.caption}"
        return FigureData(file_path=None, caption=caption, source_ids=["S-TBD"])

    caption = asset.caption.text_or_placeholder(spec.caption)
    source_origin2 = str(getattr(asset, "source_origin", "") or "").strip()
    authenticity2 = str(getattr(asset, "authenticity", "") or "").strip()
    caption = _apply_reference_caption_prefix(
        caption,
        authenticity=authenticity2,
        source_origin=source_origin2,
    )
    source_ids = asset.source_ids or asset.caption.src or ["S-TBD"]

    resolved = _resolve_existing_path(asset.file_path, asset_search_dirs)

    # Optional: validate scale metadata if present
    if spec.constraints and "map_scale_range" in spec.constraints:
        rng = str(spec.constraints.get("map_scale_range") or "").strip()
        scale_value = None
        if asset.model_extra:
            scale_value = asset.model_extra.get("scale")
        scale = _parse_scale(scale_value)
        if rng and scale is None:
            caption = f"{caption} (축척 표기 확인 필요)"

    return FigureData(
        file_path=str(resolved) if resolved else None,
        caption=caption,
        source_ids=source_ids,
        width_mm=getattr(asset, "width_mm", None),
        crop=(getattr(asset, "crop", None) or "").strip() or None,
        gen_method=_apply_authenticity_hint(
            authenticity=authenticity2,
            source_origin=source_origin2,
            gen_method=_apply_pdf_page_hint(
                file_path=asset.file_path,
                gen_method=(getattr(asset, "gen_method", None) or "").strip() or None,
            ),
        ),
    )


def build_figure_map(case: Case, specs: list[FigureSpec], *, asset_search_dirs: list[Path] | None = None) -> dict[str, FigureData]:
    return {s.id: resolve_figure(case, s, asset_search_dirs=asset_search_dirs) for s in specs}
