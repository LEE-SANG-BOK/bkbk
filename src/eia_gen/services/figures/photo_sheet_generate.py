from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path
from typing import Any

from eia_gen.models.case import Asset, Case
from eia_gen.models.fields import TextField, normalize_source_ids
from eia_gen.services.figures.callout_composite import CalloutItem, compose_callout_composite
from eia_gen.services.figures.derived_evidence import record_derived_evidence
from eia_gen.spec.models import FigureSpec


_GRID_RE = re.compile(r"(?i)(?:^|\b)grid\s*[:=]\s*(2|4|6)(?:\b|$)")


def _tfv(x: Any) -> str:
    if isinstance(x, dict):
        if "t" in x:
            return str(x.get("t") or "").strip()
        if "v" in x:
            return str(x.get("v") or "").strip()
    t = getattr(x, "t", None)
    if isinstance(t, str):
        return t.strip()
    v = getattr(x, "v", None)
    if v is not None:
        return str(v).strip()
    return str(x or "").strip()


def _extract_src_ids(row: dict[str, Any]) -> list[str]:
    ids: list[str] = []
    for v in row.values():
        if isinstance(v, dict):
            src = v.get("src")
            if isinstance(src, list):
                ids.extend([str(s).strip() for s in src if str(s).strip()])
    return normalize_source_ids(ids)


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _asset_file_exists(file_path: str, base_dir: Path) -> bool:
    fp = (file_path or "").strip()
    if not fp:
        return False
    p = Path(fp)
    if p.is_absolute():
        return p.exists()
    return p.exists() or (base_dir / p).exists()


def _parse_grid_override(asset: Asset | None) -> int | None:
    if asset is None:
        return None
    raw = str(getattr(asset, "gen_method", "") or "").strip()
    if not raw:
        return None
    m = _GRID_RE.search(raw)
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


def ensure_photo_sheets_from_attachments(
    *,
    case: Case,
    figure_specs: list[FigureSpec],
    case_xlsx: Path,
    derived_dir: Path | None = None,
    style_path: Path | None = None,
    font_path: str | None = None,
) -> dict[str, Any]:
    """Best-effort photo_sheet(CALLOUT_COMPOSITE) generation from ATTACHMENTS manifest.

    Grouping rule (SSOT): ATTACHMENTS.related_fig_id == FigureSpec.id, evidence_type == "사진".
    """

    base_dir = case_xlsx.parent.resolve()
    derived_root = (derived_dir or (base_dir / "attachments" / "derived")).resolve()
    out_dir = derived_root / "figures" / "callout"
    out_dir.mkdir(parents=True, exist_ok=True)

    attachments_manifest = (case.model_extra or {}).get("attachments_manifest")
    if not isinstance(attachments_manifest, list):
        return {"generated": [], "skipped": [], "errors": {}}

    generated: list[str] = []
    skipped: list[str] = []
    errors: dict[str, str] = {}

    for spec in figure_specs:
        if spec.asset_type != "photo_sheet":
            continue

        fig_id = spec.id
        asset = next((a for a in case.assets if a.asset_id == fig_id), None)

        # Respect existing non-callout assets (manual override).
        if asset is None:
            for a in case.assets:
                if a.type != spec.asset_type:
                    continue
                if _asset_file_exists(a.file_path, base_dir):
                    skipped.append(fig_id)
                    break
            else:
                pass
            if fig_id in skipped:
                continue
        else:
            existing = (asset.file_path or "").replace("\\", "/").strip()
            is_owned_path = existing.startswith("attachments/derived/figures/callout/")
            if existing and _asset_file_exists(existing, base_dir) and not is_owned_path:
                skipped.append(fig_id)
                continue

        photos: list[dict[str, Any]] = []
        for r in attachments_manifest:
            if not isinstance(r, dict):
                continue
            if _tfv(r.get("related_fig_id")) != fig_id:
                continue
            if _tfv(r.get("evidence_type")) != "사진":
                continue
            fp = _tfv(r.get("file_path")).replace("\\", "/")
            if not fp:
                continue
            src_path = Path(fp).expanduser()
            if not src_path.is_absolute():
                src_path = (base_dir / src_path).resolve()
            if not src_path.exists():
                continue
            photos.append(
                {
                    "evidence_id": _tfv(r.get("evidence_id")),
                    "path": src_path,
                    "title": _tfv(r.get("title")),
                    "note": _tfv(r.get("note")),
                    "src_ids": _extract_src_ids(r),
                }
            )

        photos.sort(key=lambda x: (x.get("evidence_id") or "", str(x.get("path") or "")))
        if len(photos) < 2:
            skipped.append(fig_id)
            continue

        grid_override = _parse_grid_override(asset)
        max_slots = {2: 2, 4: 4, 6: 6}.get(grid_override or 0, 6)
        truncated = False
        if len(photos) > max_slots:
            photos = photos[:max_slots]
            truncated = True
        if len(photos) > 6:
            photos = photos[:6]
            truncated = True

        out_path = out_dir / f"{fig_id}.png"
        stored_rel = f"attachments/derived/figures/callout/{fig_id}.png"

        title = str(getattr(asset, "title", "") or "").strip() if asset is not None else ""
        title = title or spec.caption

        items: list[CalloutItem] = []
        for p in photos:
            caption = (p.get("title") or "").strip()
            if not caption:
                note = (p.get("note") or "").strip()
                caption = note if note and not note.lower().startswith("sha256=") else Path(str(p["path"])).name
            items.append(CalloutItem(path=p["path"], caption=caption))

        try:
            recipe = compose_callout_composite(
                items=items,
                out_path=out_path,
                title=title,
                grid=grid_override,
                style_path=style_path,
                font_path=font_path,
            )
        except Exception as e:
            errors[fig_id] = str(e)
            continue

        # Update (or create) the asset entry so resolve_figure picks it up.
        src_ids: list[str] = []
        for p in photos:
            src_ids.extend(p.get("src_ids") or [])
        src_ids = normalize_source_ids(src_ids)
        if not src_ids:
            src_ids = ["S-TBD"]

        if asset is not None:
            if not _asset_file_exists(asset.file_path, base_dir) or asset.file_path.replace("\\", "/").strip() == stored_rel:
                asset.file_path = stored_rel
            if getattr(asset, "caption", None) is not None and asset.caption.is_empty():
                asset.caption.t = spec.caption
            if not getattr(asset, "source_ids", None) or all(s == "S-TBD" for s in (asset.source_ids or [])):
                asset.source_ids = src_ids
        else:
            case.assets.append(
                Asset(
                    asset_id=fig_id,
                    type=spec.asset_type,
                    file_path=stored_rel,
                    caption=TextField(t=spec.caption),
                    source_ids=src_ids,
                )
            )

        sha = _sha256_file(out_path)
        input_eids = [p.get("evidence_id") or "" for p in photos if (p.get("evidence_id") or "").strip()]
        note_obj: dict[str, Any] = {
            "kind": "CALLOUT_COMPOSITE",
            "sha256": sha,
            "grid": recipe.get("grid"),
            "inputs": [
                {
                    "evidence_id": p.get("evidence_id") or "",
                    "file_path": str(p.get("path") or ""),
                    "caption": items[i].caption,
                }
                for i, p in enumerate(photos)
            ],
        }
        if truncated:
            note_obj["truncated"] = True
        if input_eids:
            note_obj["input_evidence_ids"] = input_eids

        record_derived_evidence(
            case,
            derived_path=out_path,
            related_fig_id=fig_id,
            report_anchor=fig_id,
            src_ids=src_ids,
            evidence_type="derived_png",
            title=title,
            note=json.dumps(note_obj, ensure_ascii=False, sort_keys=True),
            used_in=fig_id,
            case_dir=base_dir,
        )

        generated.append(fig_id)

    return {"generated": generated, "skipped": skipped, "errors": errors}
