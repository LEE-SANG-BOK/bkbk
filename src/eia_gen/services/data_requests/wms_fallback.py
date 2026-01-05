from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from pyproj import CRS

from eia_gen.services.data_requests.wms import compute_bbox
from eia_gen.services.data_requests.xlsx_io import apply_rows_to_sheet, read_location_hint, read_sheet_dicts


_WMS_FILENAME_LAYER_RE = re.compile(
    r"_(?P<layer>[A-Z0-9_]+)(?P<suffix>__(?:FALLBACK|PLACEHOLDER)__)?\.png$",
    flags=re.IGNORECASE,
)


def _parse_int(v: Any, default: int) -> int:
    if v is None:
        return default
    if isinstance(v, (int, float)):
        return int(v)
    s = str(v).strip()
    if not s:
        return default
    try:
        return int(s)
    except Exception:
        try:
            return int(float(s))
        except Exception:
            return default


def _parse_epsg(v: Any, default: int) -> int:
    if v is None:
        return default
    if isinstance(v, (int, float)):
        return int(v)
    s = str(v).strip().upper()
    if not s:
        return default
    if s.startswith("EPSG:"):
        s = s.split("EPSG:", 1)[1].strip()
    try:
        return int(s)
    except Exception:
        return default


def _parse_wms_note(note: str) -> tuple[int | None, tuple[float, float, float, float] | None]:
    """Parse `srs=EPSG:xxxx bbox=minx,miny,maxx,maxy` from ATTACHMENTS.note.

    Supports:
    - JSON note: {"srs":"EPSG:3857","bbox":[...]}
    - Legacy: "srs=EPSG:3857 bbox=1,2,3,4 ..."
    """
    s = str(note or "")

    if s.strip().startswith("{") and s.strip().endswith("}"):
        try:
            obj = json.loads(s)
        except Exception:
            obj = None
        if isinstance(obj, dict):
            srs = str(obj.get("srs") or obj.get("out_srs") or "").strip()
            epsg = _parse_epsg(srs, default=0) or None
            raw_bbox = obj.get("bbox")
            if isinstance(raw_bbox, str):
                parts = [p.strip() for p in raw_bbox.split(",") if p.strip()]
                if len(parts) == 4:
                    try:
                        nums = [float(p) for p in parts]
                        return epsg, (nums[0], nums[1], nums[2], nums[3])
                    except Exception:
                        return epsg, None
            if isinstance(raw_bbox, (list, tuple)) and len(raw_bbox) == 4:
                try:
                    nums = [float(p) for p in raw_bbox]
                    return epsg, (nums[0], nums[1], nums[2], nums[3])
                except Exception:
                    return epsg, None
            return epsg, None

    epsg = None
    bbox = None
    for tok in s.split():
        if tok.startswith("srs="):
            epsg = _parse_epsg(tok.split("=", 1)[1], default=0) or None
        if tok.startswith("bbox="):
            raw = tok.split("=", 1)[1]
            parts = [p.strip() for p in raw.split(",") if p.strip()]
            if len(parts) == 4:
                try:
                    nums = [float(p) for p in parts]
                    bbox = (nums[0], nums[1], nums[2], nums[3])
                except Exception:
                    bbox = None
    return epsg, bbox


def _bbox_close(
    a: tuple[float, float, float, float],
    b: tuple[float, float, float, float],
    *,
    epsg: int,
    tol_projected_m: float = 3.0,
    tol_deg: float = 1e-4,
) -> bool:
    tol = tol_deg if (not CRS.from_epsg(int(epsg)).is_projected) else tol_projected_m
    for x, y in zip(a, b, strict=True):
        if abs(float(x) - float(y)) > tol:
            return False
    return True


def _extract_layer_key_from_file_name(path: str) -> tuple[str | None, str]:
    name = Path(path).name
    m = _WMS_FILENAME_LAYER_RE.search(name)
    if not m:
        return None, ""
    layer = str(m.group("layer") or "").strip().upper()
    suffix = str(m.group("suffix") or "").strip().upper()
    return (layer or None), suffix


def _is_decodable_image(path: Path) -> bool:
    try:
        from PIL import Image

        with Image.open(path) as im:
            im.verify()
        return True
    except Exception:
        return False


@dataclass(frozen=True)
class WmsEvidenceCandidate:
    rel_path: str
    abs_path: Path
    layer_key: str
    epsg: int | None
    bbox: tuple[float, float, float, float] | None
    data_origin: str
    evidence_id: str
    used_in: str
    note: str


@dataclass(frozen=True)
class WireWmsFallbackReport:
    updated_req_ids: list[str]
    enabled_req_ids: list[str]
    skipped: dict[str, str]


def wire_wms_fallbacks_for_workbook(
    *,
    wb,
    case_dir: Path,
    enable_when_wired: bool = False,
    allow_layer_aliases: bool = True,
    force: bool = False,
) -> WireWmsFallbackReport:
    """Autofill WMS fallback_file_path from local evidences in ATTACHMENTS.

    Safety defaults:
    - Only uses evidences whose note contains a bbox/srs matching the request's computed bbox.
    - Skips PLACEHOLDER images.
    - Prefers OFFICIAL_DB/CLIENT_PROVIDED evidences.

    When `force=True`, bbox/srs matching is not required (unsafe).
    """
    case_dir = case_dir.resolve()

    # 1) Build candidates from ATTACHMENTS.
    atts = read_sheet_dicts(wb, "ATTACHMENTS")
    candidates_by_layer: dict[str, list[WmsEvidenceCandidate]] = {}
    for a in atts:
        rel = str(a.get("file_path") or "").strip().replace("\\", "/")
        if not rel or "attachments/evidence/wms/" not in rel.replace("\\", "/"):
            continue
        layer, suffix = _extract_layer_key_from_file_name(rel)
        if not layer:
            continue
        if "PLACEHOLDER" in suffix:
            continue

        data_origin = str(a.get("data_origin") or "").strip().upper()
        if data_origin not in {"OFFICIAL_DB", "CLIENT_PROVIDED"}:
            continue

        abs_path = Path(rel)
        if not abs_path.is_absolute():
            abs_path = (case_dir / abs_path).resolve()
        if not abs_path.exists() or abs_path.stat().st_size <= 0:
            continue
        if not _is_decodable_image(abs_path):
            continue

        note = str(a.get("note") or "").strip()
        epsg, bbox = _parse_wms_note(note)

        c = WmsEvidenceCandidate(
            rel_path=rel,
            abs_path=abs_path,
            layer_key=layer,
            epsg=epsg,
            bbox=bbox,
            data_origin=data_origin,
            evidence_id=str(a.get("evidence_id") or "").strip(),
            used_in=str(a.get("used_in") or "").strip(),
            note=note,
        )
        candidates_by_layer.setdefault(layer, []).append(c)

    # 2) Resolve WMS requests and wire fallbacks.
    loc = read_location_hint(wb)
    center_lon = loc.get("center_lon")
    center_lat = loc.get("center_lat")
    input_epsg = int(loc.get("epsg") or 4326)
    boundary_file = str(loc.get("boundary_file") or "")

    drr = read_sheet_dicts(wb, "DATA_REQUESTS")
    if not drr:
        return WireWmsFallbackReport(updated_req_ids=[], enabled_req_ids=[], skipped={})

    updated_req_ids: list[str] = []
    enabled_req_ids: list[str] = []
    skipped: dict[str, str] = {}
    updates: list[dict[str, Any]] = []

    def _iter_layer_keys(layer_key: str) -> list[str]:
        k = (layer_key or "").strip().upper()
        if not allow_layer_aliases:
            return [k]
        if k.startswith("ECO_NATURE_"):
            return [k, "ECO_NATURE_DATAGO_OPEN", "ECO_NATURE_MCEE_2015"]
        return [k]

    for row in drr:
        req_id = str(row.get("req_id") or "").strip()
        if not req_id:
            continue
        connector = str(row.get("connector") or "").strip().upper()
        if connector != "WMS":
            continue

        try:
            params = json.loads(str(row.get("params_json") or "") or "{}")
        except Exception:
            params = {}
        if not isinstance(params, dict):
            params = {}

        layer_key = str(params.get("layer_key") or params.get("layer") or "").strip().upper()
        if not layer_key:
            skipped[req_id] = "missing params.layer_key"
            continue

        # Already wired and readable → keep.
        fb = str(
            params.get("fallback_file_path")
            or params.get("fallback_path")
            or params.get("fallback_image")
            or ""
        ).strip()
        if fb:
            fb_abs = Path(fb).expanduser()
            if not fb_abs.is_absolute():
                fb_abs = (case_dir / fb_abs).resolve()
            if fb_abs.exists():
                continue

        out_srs = str(params.get("srs") or "EPSG:3857").strip() or "EPSG:3857"
        bbox_mode = str(params.get("bbox_mode") or "AUTO").strip() or "AUTO"
        radius_m = _parse_int(params.get("radius_m"), 1000)
        out_epsg = _parse_epsg(out_srs, 3857)

        expected_bbox = None
        try:
            expected_bbox = compute_bbox(
                case_dir=case_dir,
                boundary_file=boundary_file,
                center_lon=center_lon,
                center_lat=center_lat,
                input_epsg=input_epsg,
                out_srs=out_srs,
                bbox_mode=bbox_mode,
                radius_m=radius_m,
            )
        except Exception as e:
            skipped[req_id] = f"bbox compute failed: {e}"
            continue

        best: WmsEvidenceCandidate | None = None
        for lk in _iter_layer_keys(layer_key):
            for cand in candidates_by_layer.get(lk, []):
                if not force:
                    if cand.epsg is None or cand.bbox is None:
                        continue
                    if cand.epsg != out_epsg:
                        continue
                    if expected_bbox is None:
                        continue
                    if not _bbox_close(cand.bbox, expected_bbox, epsg=out_epsg):
                        continue

                if best is None or cand.abs_path.stat().st_mtime > best.abs_path.stat().st_mtime:
                    best = cand

            if best is not None:
                break

        if best is None:
            skipped[req_id] = f"no matching local WMS evidence for layer={layer_key}"
            continue

        # Write fallback_file_path using a stable, case-relative path.
        rel_path = best.rel_path
        try:
            rel_path = str(best.abs_path.relative_to(case_dir).as_posix())
        except Exception:
            rel_path = best.rel_path

        params["fallback_file_path"] = rel_path
        row2: dict[str, Any] = {"req_id": req_id, "params_json": json.dumps(params, ensure_ascii=False)}

        if enable_when_wired:
            enabled_v = row.get("enabled")
            enabled = None
            if isinstance(enabled_v, bool):
                enabled = enabled_v
            elif enabled_v is not None:
                enabled_s = str(enabled_v).strip().upper()
                if enabled_s in {"TRUE", "T", "Y", "YES", "1"}:
                    enabled = True
                elif enabled_s in {"FALSE", "F", "N", "NO", "0"}:
                    enabled = False

            if enabled is False:
                row2["enabled"] = True
                enabled_req_ids.append(req_id)

        updates.append(row2)
        updated_req_ids.append(req_id)

    if updates:
        apply_rows_to_sheet(
            wb,
            sheet_name="DATA_REQUESTS",
            rows=updates,
            merge_strategy="UPSERT_KEYS",
            upsert_keys=["req_id"],
        )

    return WireWmsFallbackReport(updated_req_ids=updated_req_ids, enabled_req_ids=enabled_req_ids, skipped=skipped)
