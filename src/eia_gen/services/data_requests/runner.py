from __future__ import annotations

import json
import hashlib
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from eia_gen.services.data_requests.airkorea import evidence_bytes as airkorea_evidence_bytes
from eia_gen.services.data_requests.airkorea import fetch_air_baseline
from eia_gen.services.data_requests.geocode import evidence_bytes as geocode_evidence_bytes
from eia_gen.services.data_requests.geocode import geocode_address
from eia_gen.services.data_requests.kosis import KosisMapping, KosisQuery
from eia_gen.services.data_requests.kosis import evidence_bytes as kosis_evidence_bytes
from eia_gen.services.data_requests.kosis import build_env_base_socio_rows, fetch_kosis_series, resolve_kosis_dataset
from eia_gen.services.data_requests.kma_asos import evidence_bytes as kma_asos_evidence_bytes
from eia_gen.services.data_requests.kma_asos import fetch_asos_daily_precip_stats
from eia_gen.services.data_requests.kma_stations import fetch_asos_station_catalog, pick_nearest_asos_stations
from eia_gen.services.data_requests.models import Evidence
from eia_gen.services.data_requests.auto_gis import (
    WmsOverlayInput,
    overlay_from_geojson,
    overlay_from_wms_evidence,
    zoning_breakdown_from_parcels,
)
from eia_gen.services.data_requests.nier_water import evidence_bytes as nier_evidence_bytes
from eia_gen.services.data_requests.nier_water import fetch_eia_ivstg_water_quality
from eia_gen.services.data_requests.wms import compute_bbox, fetch_wms
from eia_gen.services.data_requests.xlsx_io import (
    apply_rows_to_sheet,
    append_attachment,
    read_data_requests,
    read_sheet_dicts,
    read_location_hint,
    upsert_attachment_by_used_in,
    update_request_run,
)

from eia_gen.services.figures.materialize import MaterializeOptions, materialize_figure_image

from eia_gen.services.data_requests.sanitize import redact_text


def _sha1_bytes(b: bytes) -> str:
    return hashlib.sha1(b).hexdigest()


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _note_json(meta: dict[str, Any]) -> str:
    # Stable + compact (Excel-friendly).
    try:
        return json.dumps(meta, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
    except Exception:
        # Best-effort fallback to a readable repr (should be rare).
        return str(meta)


def _evidence_id(req_id: str) -> str:
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    return f"EV-{req_id}-{ts}"


def _write_placeholder_png(out_path: Path, *, title: str, lines: list[str]) -> None:
    try:
        from PIL import Image, ImageDraw
    except Exception as e:
        raise ValueError(f"Pillow is required to write placeholder images ({e})") from None

    out_path.parent.mkdir(parents=True, exist_ok=True)
    w, h = 1400, 900
    img = Image.new("RGB", (w, h), (255, 255, 255))
    d = ImageDraw.Draw(img)

    # frame
    d.rectangle([0, 0, w - 1, h - 1], outline=(0, 0, 0), width=2)

    y = 36
    d.text((36, y), str(title), fill=(0, 0, 0))
    y += 40
    for line in lines:
        if not line:
            continue
        d.text((36, y), str(line), fill=(0, 0, 0))
        y += 28

    img.save(out_path, format="PNG", optimize=True)


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
    """Parse `srs=EPSG:xxxx bbox=minx,miny,maxx,maxy` from ATTACHMENTS.note."""
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


@dataclass(frozen=True)
class RunResult:
    executed: int
    skipped: int
    warnings: list[str]
    evidences: list[Evidence]


def run_data_requests(
    *,
    wb,
    case_dir: Path,
    wms_layers_config: Path,
    cache_config: Path,
) -> RunResult:
    reqs = read_data_requests(wb)
    warnings: list[str] = []
    evidences: list[Evidence] = []
    executed = 0
    skipped = 0

    def _is_once_complete(req) -> bool:
        """Best-effort: treat ONCE requests as complete only when evidence exists and is readable.

        This prevents stale/invalid evidences (e.g., HTML saved as .png from an auth error)
        from permanently blocking reruns.
        """
        if (req.run_mode or "").upper() != "ONCE":
            return False
        if not req.last_run_at:
            return False
        if not req.last_evidence_ids:
            return False

        # If this request depends on other DATA_REQUESTS results, re-run when upstream refreshed.
        try:
            if (req.connector or "").strip().upper() == "AUTO_GIS":
                op = str((req.params or {}).get("operation") or "").strip().upper()
                if op == "OVERLAY_FROM_WMS_EVIDENCE":
                    items = (req.params or {}).get("items") or []
                    from_ids = []
                    if isinstance(items, list):
                        for it in items:
                            if not isinstance(it, dict):
                                continue
                            fid = str(it.get("from_req_id") or "").strip()
                            if fid:
                                from_ids.append(fid)

                    if from_ids:
                        try:
                            req_dt = datetime.fromisoformat(str(req.last_run_at).strip())
                        except Exception:
                            req_dt = None

                        if req_dt is not None:
                            cur_reqs = read_data_requests(wb)
                            by_id = {r.req_id: r for r in cur_reqs}
                            for fid in from_ids:
                                up = by_id.get(fid)
                                if not up or not up.last_run_at:
                                    continue
                                try:
                                    up_dt = datetime.fromisoformat(str(up.last_run_at).strip())
                                except Exception:
                                    continue
                                if up_dt > req_dt:
                                    return False
        except Exception:
            # Don't block execution on dependency parsing errors.
            pass

        atts = read_sheet_dicts(wb, "ATTACHMENTS")
        by_ev = {str(a.get("evidence_id") or "").strip(): a for a in atts}

        for ev_id in req.last_evidence_ids:
            ev_id = str(ev_id or "").strip()
            if not ev_id:
                return False
            a = by_ev.get(ev_id)
            if not a:
                return False
            rel = str(a.get("file_path") or "").strip()
            if not rel:
                return False
            if "__PLACEHOLDER__" in rel:
                return False
            p = Path(rel)
            if not p.is_absolute():
                p = (case_dir / p).resolve()
            if not p.exists() or p.stat().st_size <= 0:
                return False
            if "__PLACEHOLDER__" in p.name:
                return False

            # If evidence.note indicates it's a placeholder, do not treat the request as complete.
            note = str(a.get("note") or "").strip()
            if note:
                if "__PLACEHOLDER__" in note:
                    return False
                try:
                    obj = json.loads(note)
                except Exception:
                    obj = None
                if isinstance(obj, dict):
                    kind = str(obj.get("kind") or "").strip().upper()
                    if "PLACEHOLDER" in kind:
                        return False

            # If it looks like an image, verify it's actually decodable.
            if p.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp"}:
                try:
                    from PIL import Image

                    with Image.open(p) as im:
                        im.verify()
                except Exception:
                    return False

        return True

    for req in reqs:
        if not req.enabled or req.run_mode.upper() == "NEVER":
            skipped += 1
            continue
        # ONCE: skip only when evidence is still healthy.
        if _is_once_complete(req):
            skipped += 1
            continue
        if req.run_mode.upper() == "ONCE" and req.last_run_at:
            warnings.append(f"[{req.req_id}] ONCE request has missing/invalid evidence; re-running")

        attempted = False
        try:
            # Refresh LOCATION hints per request so earlier requests (e.g. GEOCODE) can
            # influence later ones within the same run.
            loc = read_location_hint(wb)
            params = dict(req.params or {})
            connector = (req.connector or "").strip().upper()

            if connector == "GEOCODE":
                attempted = True
                executed += 1

                provider = str(params.get("provider") or "AUTO").strip()

                address = str(params.get("address") or "").strip()
                if not address:
                    loc_rows = read_sheet_dicts(wb, "LOCATION")
                    if loc_rows:
                        address = str(loc_rows[0].get("address_road") or "").strip() or str(
                            loc_rows[0].get("address_jibeon") or ""
                        ).strip()
                if not address:
                    raise ValueError("Missing address for GEOCODE (LOCATION.address_road/address_jibeon or params_json.address)")

                res = geocode_address(address=address, provider=provider)

                ev_id = _evidence_id(req.req_id)
                ev_rel = Path("attachments/evidence/api") / f"{ev_id}_geocode.json"
                ev_abs = (case_dir / ev_rel).resolve()
                ev_abs.parent.mkdir(parents=True, exist_ok=True)
                ev_bytes = geocode_evidence_bytes(res.evidence_json)
                ev_abs.write_bytes(ev_bytes)

                # Update LOCATION sheet (preserve existing text fields).
                loc_rows = read_sheet_dicts(wb, "LOCATION")
                base = dict(loc_rows[0]) if loc_rows else {}
                base["center_lat"] = res.lat
                base["center_lon"] = res.lon
                base["crs"] = str(base.get("crs") or "EPSG:4326")
                base["src_id"] = req.src_id or str(base.get("src_id") or "S-TBD")

                sheet_warn = apply_rows_to_sheet(
                    wb,
                    sheet_name=req.output_sheet.strip() or "LOCATION",
                    rows=[base],
                    merge_strategy=req.merge_strategy or "REPLACE_SHEET",
                    upsert_keys=req.upsert_keys,
                )
                warnings.extend([f"[{req.req_id}] {w}" for w in sheet_warn])

                ev = Evidence(
                    evidence_id=ev_id,
                    evidence_type="기타",
                    title=f"GEOCODE:{res.provider}",
                    file_path=str(ev_rel).replace("\\", "/"),
                    used_in=f"DATA_REQUESTS:{req.req_id}",
                    data_origin="OFFICIAL_DB" if res.provider.upper() == "VWORLD" else "LITERATURE",
                    src_id=req.src_id or "S-TBD",
                    note=_note_json(
                        {
                            "connector": "GEOCODE",
                            "req_id": req.req_id,
                            "retrieved_at": _now_iso(),
                            "provider": res.provider,
                            "address": res.address,
                            "request_url": (res.evidence_json.get("request") or {}).get("url") if isinstance(res.evidence_json, dict) else "",
                            "request_params": (res.evidence_json.get("request") or {}).get("params") if isinstance(res.evidence_json, dict) else {},
                            "hash_sha1": _sha1_bytes(ev_bytes),
                        }
                    ),
                )
                append_attachment(wb, ev)
                update_request_run(wb, req_id=req.req_id, evidence_ids=[ev_id])
                evidences.append(ev)

            elif connector == "PDF_PAGE":
                attempted = True
                executed += 1

                raw_path = str(params.get("pdf_path") or params.get("file_path") or "").strip()
                if not raw_path:
                    raise ValueError("Missing params.pdf_path for PDF_PAGE")

                pdf_path = Path(raw_path)
                if not pdf_path.is_absolute():
                    pdf_path = (case_dir / pdf_path).resolve()
                if not pdf_path.exists():
                    raise ValueError(f"PDF not found: {pdf_path}")

                page_1based = _parse_int(params.get("page") or params.get("page_1based"), 0)
                if page_1based <= 0:
                    raise ValueError("Missing params.page (1-based) for PDF_PAGE")

                crop = str(params.get("crop") or "").strip() or None

                width_mm = None
                try:
                    if params.get("width_mm") not in (None, ""):
                        width_mm = float(params.get("width_mm"))
                except Exception:
                    width_mm = None

                dpi = _parse_int(params.get("dpi"), 250)
                max_width_px = _parse_int(params.get("max_width_px"), 2600)

                ev_id = _evidence_id(req.req_id)
                out_dir = case_dir / "attachments/evidence/pdf"

                try:
                    out_path = materialize_figure_image(
                        pdf_path,
                        MaterializeOptions(
                            out_dir=out_dir,
                            fig_id=ev_id,
                            gen_method=f"PDF_PAGE:{page_1based}",
                            crop=crop,
                            width_mm=width_mm,
                            target_dpi=dpi,
                            max_width_px=max_width_px,
                        ),
                    )
                except ImportError as e:
                    raise ValueError(f"PyMuPDF(fitz) is required for PDF_PAGE. ({e})")

                try:
                    ev_rel = out_path.relative_to(case_dir)
                except Exception:
                    ev_rel = out_path

                ev_bytes = out_path.read_bytes()

                title = str(params.get("title") or f"PDF_PAGE:{pdf_path.name} p{page_1based}").strip()
                data_origin = str(params.get("data_origin") or "LITERATURE").strip() or "LITERATURE"

                ev = Evidence(
                    evidence_id=ev_id,
                    evidence_type="기타",
                    title=title,
                    file_path=str(ev_rel).replace("\\", "/"),
                    used_in=f"DATA_REQUESTS:{req.req_id}",
                    data_origin=data_origin,
                    src_id=req.src_id or "S-TBD",
                    note=_note_json(
                        {
                            "connector": "PDF_PAGE",
                            "req_id": req.req_id,
                            "retrieved_at": _now_iso(),
                            "pdf_path": str(pdf_path),
                            "page_1based": int(page_1based),
                            "dpi": int(dpi),
                            "crop": crop or "",
                            "width_mm": width_mm,
                            "max_width_px": int(max_width_px),
                            "hash_sha1": _sha1_bytes(ev_bytes),
                        }
                    ),
                )
                append_attachment(wb, ev)
                update_request_run(wb, req_id=req.req_id, evidence_ids=[ev_id])
                evidences.append(ev)

            elif connector == "WMS":
                attempted = True
                executed += 1
                layer_key = str(params.get("layer_key") or params.get("layer") or "").strip()
                if not layer_key:
                    raise ValueError("missing params.layer_key")

                out_srs = str(params.get("srs") or "EPSG:3857")
                bbox_mode = str(params.get("bbox_mode") or "AUTO")
                width = _parse_int(params.get("width") or params.get("image_width"), 2048)
                height = _parse_int(params.get("height") or params.get("image_height"), 2048)
                radius_m = _parse_int(params.get("radius_m"), 1000)

                # Optional fallback: user-provided official image when WMS is blocked (key/approval/outage).
                fallback_file_path = str(
                    params.get("fallback_file_path") or params.get("fallback_path") or params.get("fallback_image") or ""
                ).strip()
                fb_abs: Path | None = None
                if fallback_file_path:
                    fb_abs = Path(fallback_file_path).expanduser()
                    if not fb_abs.is_absolute():
                        fb_abs = (case_dir / fb_abs).resolve()
                    if not fb_abs.exists():
                        fb_abs = None

                ev_id = _evidence_id(req.req_id)
                used_in = f"DATA_REQUESTS:{req.req_id}"
                ev_dir_rel = Path("attachments/evidence/wms")
                ev_dir_abs = (case_dir / ev_dir_rel).resolve()
                ev_dir_abs.mkdir(parents=True, exist_ok=True)

                bbox: tuple[float, float, float, float] | None = None
                fetched = None
                try:
                    bbox = compute_bbox(
                        case_dir=case_dir,
                        boundary_file=str(loc.get("boundary_file") or ""),
                        center_lon=loc.get("center_lon"),
                        center_lat=loc.get("center_lat"),
                        input_epsg=int(loc.get("epsg") or 4326),
                        out_srs=out_srs,
                        bbox_mode=bbox_mode,
                        radius_m=radius_m,
                    )

                    fetched = fetch_wms(
                        layer_key=layer_key,
                        bbox=bbox,
                        width=width,
                        height=height,
                        out_srs=out_srs,
                        wms_layers_config=wms_layers_config,
                        cache_config=cache_config,
                    )

                    ev_rel = ev_dir_rel / f"{ev_id}_{layer_key}.png"
                    ev_abs = (case_dir / ev_rel).resolve()
                    ev_abs.write_bytes(fetched.bytes_)

                    ev = Evidence(
                        evidence_id=ev_id,
                        evidence_type="기타",
                        title=f"WMS:{layer_key}",
                        file_path=str(ev_rel).replace("\\", "/"),
                        used_in=used_in,
                        data_origin="OFFICIAL_DB",
                        src_id=req.src_id or "S-TBD",
                        note=_note_json(
                            {
                                "kind": "WMS",
                                "connector": "WMS",
                                "req_id": req.req_id,
                                "retrieved_at": _now_iso(),
                                "layer_key": layer_key,
                                "srs": out_srs,
                                "bbox": [float(x) for x in (bbox or (0, 0, 0, 0))],
                                "cache_hit": bool(getattr(fetched, "cache_hit", False)),
                                "request_url": str(getattr(fetched, "request_url", "") or ""),
                                "request_params": getattr(fetched, "request_params", None),
                                "content_type": str(getattr(fetched, "content_type", "") or ""),
                                "hash_sha1": _sha1_bytes(fetched.bytes_),
                            }
                        ),
                    )
                    upsert_attachment_by_used_in(wb, ev)
                    update_request_run(wb, req_id=req.req_id, evidence_ids=[ev_id])
                    evidences.append(ev)
                except Exception as e:
                    err = redact_text(str(e))

                    # If a user-provided fallback exists, use it instead of a placeholder.
                    if fb_abs is not None:
                        try:
                            from PIL import Image
                            import io

                            with Image.open(fb_abs) as im:
                                buf = io.BytesIO()
                                im.convert("RGB").save(buf, format="PNG", optimize=True)
                                fb_bytes = buf.getvalue()
                        except Exception as ee:
                            fb_bytes = b""
                            warnings.append(
                                f"[{req.req_id}] WMS fallback_image invalid: {redact_text(str(ee))}"
                            )

                        if fb_bytes:
                            ev_rel = ev_dir_rel / f"{ev_id}_{layer_key}__FALLBACK__.png"
                            ev_abs = (case_dir / ev_rel).resolve()
                            ev_abs.write_bytes(fb_bytes)

                            ev = Evidence(
                                evidence_id=ev_id,
                                evidence_type="기타",
                                title=f"WMS:{layer_key} (fallback image)",
                                file_path=str(ev_rel).replace("\\", "/"),
                                used_in=used_in,
                                data_origin="CLIENT_PROVIDED",
                                src_id=req.src_id or "S-TBD",
                                note=_note_json(
                                    {
                                        "kind": "WMS_FALLBACK_LOCAL_FILE",
                                        "connector": "WMS",
                                        "req_id": req.req_id,
                                        "retrieved_at": _now_iso(),
                                        "layer_key": layer_key,
                                        "srs": out_srs,
                                        "bbox": [float(x) for x in bbox] if bbox else None,
                                        "fallback_file_path": str(fb_abs),
                                        "wms_error": err,
                                        "hash_sha1": _sha1_bytes(fb_bytes),
                                    }
                                ),
                            )
                            upsert_attachment_by_used_in(wb, ev)
                            update_request_run(wb, req_id=req.req_id, evidence_ids=[ev_id])
                            evidences.append(ev)
                            warnings.append(f"[{req.req_id}] WMS failed; used fallback_image instead ({layer_key})")
                            continue

                    # Final fallback: generate a placeholder PNG so the pipeline doesn't silently drop evidences.
                    ev_rel = ev_dir_rel / f"{ev_id}_{layer_key}__PLACEHOLDER__.png"
                    ev_abs = (case_dir / ev_rel).resolve()
                    _write_placeholder_png(
                        ev_abs,
                        title=f"WMS fetch failed (placeholder): {layer_key}",
                        lines=[
                            f"srs: {out_srs}",
                            f"bbox_mode: {bbox_mode}",
                            f"radius_m: {radius_m}",
                            f"error: {err}",
                            "fix: set DATA_REQUESTS.params_json.fallback_file_path to an official screenshot/image,",
                            "     or resolve API key/approval and re-run.",
                        ],
                    )
                    b0 = ev_abs.read_bytes()

                    ev = Evidence(
                        evidence_id=ev_id,
                        evidence_type="기타",
                        title=f"WMS:{layer_key} (placeholder)",
                        file_path=str(ev_rel).replace("\\", "/"),
                        used_in=used_in,
                        data_origin="MODEL_OUTPUT",
                        src_id=req.src_id or "S-TBD",
                        note=_note_json(
                            {
                                "kind": "WMS_PLACEHOLDER",
                                "connector": "WMS",
                                "req_id": req.req_id,
                                "retrieved_at": _now_iso(),
                                "layer_key": layer_key,
                                "srs": out_srs,
                                "bbox": [float(x) for x in bbox] if bbox else None,
                                "bbox_mode": bbox_mode,
                                "radius_m": int(radius_m),
                                "wms_error": err,
                                "hash_sha1": _sha1_bytes(b0),
                            }
                        ),
                    )
                    upsert_attachment_by_used_in(wb, ev)
                    update_request_run(wb, req_id=req.req_id, evidence_ids=[ev_id])
                    evidences.append(ev)
                    warnings.append(f"[{req.req_id}] WMS failed; wrote placeholder evidence ({layer_key})")
                    continue

            elif connector == "AUTO_GIS":
                attempted = True
                executed += 1
                op = str(params.get("operation") or "").strip().upper()
                if not op:
                    if (req.output_sheet or "").strip().upper() == "ZONING_BREAKDOWN":
                        op = "ZONING_BREAKDOWN_FROM_PARCELS"
                    elif (req.output_sheet or "").strip().upper() == "ZONING_OVERLAY":
                        op = "OVERLAY_FROM_GEOJSON"

                if op == "ZONING_BREAKDOWN_FROM_PARCELS":
                    parcels = read_sheet_dicts(wb, "PARCELS")
                    out = zoning_breakdown_from_parcels(parcels_rows=parcels, req_id=req.req_id)
                    warnings.extend([f"[{req.req_id}] {w}" for w in out.warnings])

                    ev_id = _evidence_id(req.req_id)
                    ev_rel = Path("attachments/evidence/calc") / f"{ev_id}_{out.evidence_filename}"
                    ev_abs = (case_dir / ev_rel).resolve()
                    ev_abs.parent.mkdir(parents=True, exist_ok=True)
                    ev_abs.write_bytes(out.evidence_bytes)

                    # inject evidence_id into rows when target sheet supports it
                    rows = []
                    for r in out.rows:
                        rr = dict(r)
                        rr["evidence_id"] = ev_id
                        rows.append(rr)

                    sheet = req.output_sheet.strip() or "ZONING_BREAKDOWN"
                    sheet_warn = apply_rows_to_sheet(
                        wb,
                        sheet_name=sheet,
                        rows=rows,
                        merge_strategy=req.merge_strategy or "REPLACE_SHEET",
                        upsert_keys=req.upsert_keys,
                    )
                    warnings.extend([f"[{req.req_id}] {w}" for w in sheet_warn])

                    ev = Evidence(
                        evidence_id=ev_id,
                        evidence_type="계산서",
                        title="AUTO_GIS:ZONING_BREAKDOWN",
                        file_path=str(ev_rel).replace("\\", "/"),
                        used_in=f"DATA_REQUESTS:{req.req_id}",
                        data_origin="MODEL_OUTPUT",
                        src_id=req.src_id or "S-TBD",
                        note=_note_json(
                            {
                                "connector": "AUTO_GIS",
                                "req_id": req.req_id,
                                "retrieved_at": _now_iso(),
                                "operation": "ZONING_BREAKDOWN_FROM_PARCELS",
                                "hash_sha1": _sha1_bytes(out.evidence_bytes),
                            }
                        ),
                    )
                    append_attachment(wb, ev)
                    update_request_run(wb, req_id=req.req_id, evidence_ids=[ev_id])
                    evidences.append(ev)

                elif op == "OVERLAY_FROM_GEOJSON":
                    overlays = params.get("overlays") or []
                    if not isinstance(overlays, list):
                        overlays = []
                    metric_epsg = _parse_int(params.get("metric_epsg"), 5186)
                    boundary_file = str(params.get("boundary_file") or loc.get("boundary_file") or "").strip()
                    boundary_epsg = _parse_int(params.get("boundary_epsg"), int(loc.get("epsg") or 4326))

                    out = overlay_from_geojson(
                        case_dir=case_dir,
                        boundary_file=boundary_file,
                        boundary_epsg=boundary_epsg,
                        overlays=overlays,
                        metric_epsg=metric_epsg,
                        req_id=req.req_id,
                    )
                    warnings.extend([f"[{req.req_id}] {w}" for w in out.warnings])

                    ev_id = _evidence_id(req.req_id)
                    ev_rel = Path("attachments/evidence/gis") / f"{ev_id}_{out.evidence_filename}"
                    ev_abs = (case_dir / ev_rel).resolve()
                    ev_abs.parent.mkdir(parents=True, exist_ok=True)
                    ev_abs.write_bytes(out.evidence_bytes)

                    # include evidence id in basis for traceability (sheet has no evidence_id col)
                    rows = []
                    for r in out.rows:
                        rr = dict(r)
                        basis = str(rr.get("basis") or "").strip()
                        rr["basis"] = f"{basis} evidence={ev_id}".strip()
                        rows.append(rr)

                    sheet = req.output_sheet.strip() or "ZONING_OVERLAY"
                    sheet_warn = apply_rows_to_sheet(
                        wb,
                        sheet_name=sheet,
                        rows=rows,
                        merge_strategy=req.merge_strategy or "UPSERT_KEYS",
                        upsert_keys=req.upsert_keys or ["overlay_id"],
                    )
                    warnings.extend([f"[{req.req_id}] {w}" for w in sheet_warn])

                    ev = Evidence(
                        evidence_id=ev_id,
                        evidence_type="계산서",
                        title="AUTO_GIS:OVERLAY",
                        file_path=str(ev_rel).replace("\\", "/"),
                        used_in=f"DATA_REQUESTS:{req.req_id}",
                        data_origin="MODEL_OUTPUT",
                        src_id=req.src_id or "S-TBD",
                        note=_note_json(
                            {
                                "connector": "AUTO_GIS",
                                "req_id": req.req_id,
                                "retrieved_at": _now_iso(),
                                "operation": "OVERLAY_FROM_GEOJSON",
                                "metric_epsg": int(metric_epsg),
                                "hash_sha1": _sha1_bytes(out.evidence_bytes),
                            }
                        ),
                    )
                    append_attachment(wb, ev)
                    update_request_run(wb, req_id=req.req_id, evidence_ids=[ev_id])
                    evidences.append(ev)

                elif op == "OVERLAY_FROM_WMS_EVIDENCE":
                    items = params.get("items") or params.get("overlays") or []
                    if not isinstance(items, list):
                        items = []

                    alpha_threshold = _parse_int(params.get("alpha_threshold"), 10)
                    analysis_max_size = _parse_int(params.get("analysis_max_size"), 512)
                    distance_sample_stride = _parse_int(params.get("distance_sample_stride"), 4)
                    distance_max_points = _parse_int(params.get("distance_max_points"), 5000)
                    metric_epsg = _parse_int(params.get("metric_epsg"), 5186)
                    radius_m = _parse_int(params.get("radius_m"), 1000)

                    boundary_file = str(params.get("boundary_file") or loc.get("boundary_file") or "").strip()
                    boundary_epsg = _parse_epsg(params.get("boundary_epsg"), int(loc.get("epsg") or 4326))

                    # Resolve WMS evidences from ATTACHMENTS and/or upstream DATA_REQUESTS.
                    atts = read_sheet_dicts(wb, "ATTACHMENTS")
                    att_by_ev = {str(a.get("evidence_id") or "").strip(): a for a in atts}

                    resolved: list[WmsOverlayInput] = []
                    for it in items:
                        if not isinstance(it, dict):
                            continue

                        overlay_id = str(it.get("overlay_id") or it.get("id") or "").strip()
                        category = str(it.get("category") or "").strip()
                        designation_name = str(it.get("designation_name") or it.get("name") or "").strip()
                        src_id = str(it.get("src_id") or req.src_id or "S-TBD").strip() or "S-TBD"

                        evidence_id = str(it.get("evidence_id") or "").strip()
                        from_req_id = str(it.get("from_req_id") or "").strip()
                        if not evidence_id and from_req_id:
                            for rr in read_data_requests(wb):
                                if rr.req_id == from_req_id and rr.last_evidence_ids:
                                    evidence_id = rr.last_evidence_ids[0]
                                    break

                        if not overlay_id:
                            warnings.append(f"[{req.req_id}] OVERLAY_FROM_WMS_EVIDENCE item missing overlay_id")
                            continue

                        att = att_by_ev.get(evidence_id)
                        if not att:
                            warnings.append(
                                f"[{req.req_id}] [{overlay_id}] missing WMS evidence in ATTACHMENTS: evidence_id={evidence_id or '-'}"
                            )
                            continue

                        rel_path = str(att.get("file_path") or "").strip()
                        if not rel_path:
                            warnings.append(f"[{req.req_id}] [{overlay_id}] missing file_path for evidence_id={evidence_id}")
                            continue

                        img_path = Path(rel_path)
                        if not img_path.is_absolute():
                            img_path = (case_dir / img_path).resolve()

                        note = str(att.get("note") or "").strip()
                        epsg2, bbox = _parse_wms_note(note)
                        if bbox is None:
                            warnings.append(f"[{req.req_id}] [{overlay_id}] missing bbox in evidence note: {evidence_id}")
                            continue
                        epsg_img = epsg2 or _parse_epsg(it.get("epsg"), 3857)

                        basis = str(it.get("basis") or "").strip() or f"AUTO_GIS(WMS evidence={evidence_id})"
                        resolved.append(
                            WmsOverlayInput(
                                overlay_id=overlay_id,
                                category=category,
                                designation_name=designation_name,
                                image_path=img_path,
                                image_bbox=bbox,
                                image_epsg=epsg_img,
                                src_id=src_id,
                                basis=basis,
                            )
                        )

                    out = overlay_from_wms_evidence(
                        case_dir=case_dir,
                        boundary_file=boundary_file,
                        boundary_epsg=boundary_epsg,
                        center_lon=loc.get("center_lon"),
                        center_lat=loc.get("center_lat"),
                        center_epsg=int(loc.get("epsg") or 4326),
                        radius_m=radius_m,
                        items=resolved,
                        req_id=req.req_id,
                        metric_epsg=metric_epsg,
                        alpha_threshold=alpha_threshold,
                        analysis_max_size=analysis_max_size,
                        distance_sample_stride=distance_sample_stride,
                        distance_max_points=distance_max_points,
                    )
                    warnings.extend([f"[{req.req_id}] {w}" for w in out.warnings])

                    ev_id = _evidence_id(req.req_id)
                    ev_rel = Path("attachments/evidence/gis") / f"{ev_id}_{out.evidence_filename}"
                    ev_abs = (case_dir / ev_rel).resolve()
                    ev_abs.parent.mkdir(parents=True, exist_ok=True)
                    ev_abs.write_bytes(out.evidence_bytes)

                    # include evidence id in basis for traceability (sheet has no evidence_id col)
                    rows = []
                    for r in out.rows:
                        rr = dict(r)
                        basis = str(rr.get("basis") or "").strip()
                        rr["basis"] = f"{basis} evidence={ev_id}".strip()
                        rows.append(rr)

                    sheet = req.output_sheet.strip() or "ZONING_OVERLAY"
                    sheet_warn = apply_rows_to_sheet(
                        wb,
                        sheet_name=sheet,
                        rows=rows,
                        merge_strategy=req.merge_strategy or "UPSERT_KEYS",
                        upsert_keys=req.upsert_keys or ["overlay_id"],
                    )
                    warnings.extend([f"[{req.req_id}] {w}" for w in sheet_warn])

                    ev = Evidence(
                        evidence_id=ev_id,
                        evidence_type="계산서",
                        title="AUTO_GIS:WMS_OVERLAY",
                        file_path=str(ev_rel).replace("\\", "/"),
                        used_in=f"DATA_REQUESTS:{req.req_id}",
                        data_origin="MODEL_OUTPUT",
                        src_id=req.src_id or "S-TBD",
                        note=_note_json(
                            {
                                "connector": "AUTO_GIS",
                                "req_id": req.req_id,
                                "retrieved_at": _now_iso(),
                                "operation": "OVERLAY_FROM_WMS_EVIDENCE",
                                "metric_epsg": int(metric_epsg),
                                "hash_sha1": _sha1_bytes(out.evidence_bytes),
                            }
                        ),
                    )
                    append_attachment(wb, ev)
                    update_request_run(wb, req_id=req.req_id, evidence_ids=[ev_id])
                    evidences.append(ev)

                else:
                    warnings.append(f"[{req.req_id}] AUTO_GIS unknown operation: {op!r}")

            elif connector == "AIRKOREA":
                attempted = True
                executed += 1

                lon = params.get("center_lon", loc.get("center_lon"))
                lat = params.get("center_lat", loc.get("center_lat"))
                if lon is None or lat is None:
                    raise ValueError("Missing LOCATION.center_lon/center_lat for AIRKOREA")
                lon_f = float(lon)
                lat_f = float(lat)

                station_override = str(params.get("station_name") or "").strip()
                data_term = str(params.get("data_term") or "MONTH").strip() or "MONTH"
                num_rows = _parse_int(params.get("num_rows"), 200)

                baseline = fetch_air_baseline(
                    center_lon=lon_f,
                    center_lat=lat_f,
                    station_name_override=station_override,
                    data_term=data_term,
                    num_rows=num_rows,
                )

                ev_id = _evidence_id(req.req_id)
                ev_rel = Path("attachments/evidence/api") / f"{ev_id}_airkorea.json"
                ev_abs = (case_dir / ev_rel).resolve()
                ev_abs.parent.mkdir(parents=True, exist_ok=True)
                ev_bytes = airkorea_evidence_bytes(baseline.evidence_json)
                ev_abs.write_bytes(ev_bytes)

                sheet_rows = []
                unit_map = {"PM10": "µg/m3", "PM2.5": "µg/m3", "O3": "ppm"}
                for pol, v in baseline.values.items():
                    sheet_rows.append(
                        {
                            "air_id": f"{ev_id}-{pol}",
                            "station_name": baseline.station_name,
                            "station_distance_km": baseline.station_distance_km,
                            "period_start": baseline.period_start,
                            "period_end": baseline.period_end,
                            "pollutant": pol,
                            "value_avg": round(float(v), 4),
                            "unit": unit_map.get(pol, ""),
                            "data_origin": "OFFICIAL_DB",
                            "src_id": req.src_id or "S-TBD",
                            "evidence_id": ev_id,
                        }
                    )

                sheet = req.output_sheet.strip() or "ENV_BASE_AIR"
                sheet_warn = apply_rows_to_sheet(
                    wb,
                    sheet_name=sheet,
                    rows=sheet_rows,
                    merge_strategy=req.merge_strategy or "REPLACE_SHEET",
                    upsert_keys=req.upsert_keys,
                )
                warnings.extend([f"[{req.req_id}] {w}" for w in sheet_warn])

                ev = Evidence(
                    evidence_id=ev_id,
                    evidence_type="측정원시자료",
                    title=f"AIRKOREA:{baseline.station_name}",
                    file_path=str(ev_rel).replace("\\", "/"),
                    used_in=f"DATA_REQUESTS:{req.req_id}",
                    data_origin="OFFICIAL_DB",
                    src_id=req.src_id or "S-TBD",
                    note=_note_json(
                        {
                            "connector": "AIRKOREA",
                            "req_id": req.req_id,
                            "retrieved_at": _now_iso(),
                            "data_term": data_term,
                            "station_name": baseline.station_name,
                            "request_url": (baseline.evidence_json.get("measure_request") or {}).get("url")
                            if isinstance(baseline.evidence_json, dict)
                            else "",
                            "request_params": (baseline.evidence_json.get("measure_request") or {}).get("params")
                            if isinstance(baseline.evidence_json, dict)
                            else {},
                            "hash_sha1": _sha1_bytes(ev_bytes),
                        }
                    ),
                )
                append_attachment(wb, ev)
                update_request_run(wb, req_id=req.req_id, evidence_ids=[ev_id])
                evidences.append(ev)

            elif connector == "KOSIS":
                attempted = True
                executed += 1

                admin_code = str(params.get("admin_code") or "").strip()
                if not admin_code:
                    raise ValueError("Missing params.admin_code for KOSIS")

                admin_name = str(params.get("admin_name") or "").strip()
                if not admin_name:
                    loc_rows = read_sheet_dicts(wb, "LOCATION")
                    if loc_rows:
                        a1 = str(loc_rows[0].get("admin_si") or "").strip()
                        a2 = str(loc_rows[0].get("admin_eupmyeon") or "").strip()
                        admin_name = (a1 + (" " + a2 if a2 else "")).strip()

                dataset_keys: list[str] = []
                raw_keys = params.get("dataset_keys")
                if isinstance(raw_keys, str):
                    dataset_keys = [k.strip() for k in raw_keys.split(",") if k.strip()]
                elif isinstance(raw_keys, list):
                    dataset_keys = [str(k or "").strip() for k in raw_keys if str(k or "").strip()]
                dk = str(params.get("dataset_key") or "").strip()
                if dk and dk not in dataset_keys:
                    dataset_keys = [dk] + dataset_keys

                computed_rows: list[dict[str, Any]] = []
                ev_json: dict[str, Any] = {}
                reqs_meta: list[dict[str, Any]] = []

                if dataset_keys:
                    # Resolve via SSOT `config/kosis_datasets.yaml` to avoid per-case hardcoding.
                    cfg_override = str(params.get("kosis_datasets_config") or params.get("datasets_config") or "").strip()
                    if cfg_override:
                        cfg_path = Path(cfg_override)
                        if not cfg_path.is_absolute():
                            cfg_path = (wms_layers_config.parent / cfg_path).resolve()
                    else:
                        cfg_path = (wms_layers_config.parent / "kosis_datasets.yaml").resolve()

                    def _year(v: Any) -> str:
                        s = str(v or "").strip()
                        if not s:
                            return ""
                        return s[:4] if len(s) >= 4 else s

                    from datetime import datetime as _dt

                    today = _dt.now()
                    default_end_year = today.year - 1
                    default_start_year = default_end_year - 4

                    start_year = _year(params.get("start_year") or params.get("start_yr")) or str(default_start_year)
                    end_year = _year(params.get("end_year") or params.get("end_yr") or params.get("year")) or str(default_end_year)

                    ctx = {
                        "admin_code": admin_code,
                        "admin_name": admin_name,
                        "start_year": start_year,
                        "end_year": end_year,
                        "year": end_year,
                    }

                    by_year: dict[str, dict[str, Any]] = {}
                    merge_warnings: list[str] = []
                    dataset_runs: list[dict[str, Any]] = []

                    for dataset_key in dataset_keys:
                        q, mappings, ds_meta = resolve_kosis_dataset(
                            dataset_key=dataset_key,
                            config_path=cfg_path,
                            context=ctx,
                        )
                        fetched = fetch_kosis_series(q=q)
                        part_rows = build_env_base_socio_rows(
                            items=fetched.items,
                            mappings=mappings,
                            admin_code=admin_code,
                            admin_name=admin_name,
                        )
                        dataset_runs.append(
                            {
                                "dataset": ds_meta,
                                "fetch": fetched.evidence_json,
                                "output_rows": part_rows,
                            }
                        )
                        req_obj = (fetched.evidence_json.get("request") or {}) if isinstance(fetched.evidence_json, dict) else {}
                        reqs_meta.append(
                            {
                                "dataset_key": ds_meta.get("dataset_key"),
                                "url": req_obj.get("url"),
                                "params": req_obj.get("params"),
                            }
                        )
                        for r in part_rows:
                            year = str(r.get("year") or "").strip()
                            if not year:
                                continue
                            cur = by_year.setdefault(
                                year,
                                {
                                    "admin_code": admin_code,
                                    "admin_name": admin_name,
                                    "year": year,
                                },
                            )
                            for k, v in r.items():
                                if k in {"admin_code", "admin_name", "year"}:
                                    continue
                                if v is None:
                                    continue
                                if k in cur and cur.get(k) not in {None, ""} and cur.get(k) != v:
                                    merge_warnings.append(f"conflict year={year} col={k} {cur.get(k)!r}!={v!r}")
                                cur[k] = v

                    computed_rows = [by_year[y] for y in sorted(by_year.keys())]
                    if not computed_rows:
                        raise ValueError("KOSIS returned no mappable rows (check dataset_keys/config)")

                    ev_json = {
                        "generated_at": _now_iso(),
                        "dataset_keys": dataset_keys,
                        "datasets": dataset_runs,
                        "computed": {
                            "admin_code": admin_code,
                            "admin_name": admin_name,
                            "output_rows": computed_rows,
                            "merge_warnings": merge_warnings,
                        },
                    }
                else:
                    # Backward-compat: allow fully-specified query_params+mappings in the case.xlsx.
                    query_params = params.get("query_params")
                    if not isinstance(query_params, dict):
                        raise ValueError("KOSIS params.query_params must be an object")

                    mappings_raw = params.get("mappings")
                    if not isinstance(mappings_raw, list) or not mappings_raw:
                        raise ValueError(
                            "KOSIS params.mappings is required (list of {output_col, match_itm_id/match_itm_nm_contains})"
                        )
                    mappings: list[KosisMapping] = []
                    for m in mappings_raw:
                        if not isinstance(m, dict):
                            continue
                        output_col = str(m.get("output_col") or m.get("col") or "").strip()
                        if not output_col:
                            continue
                        mappings.append(
                            KosisMapping(
                                output_col=output_col,
                                match_itm_id=str(m.get("match_itm_id") or m.get("itm_id") or "").strip(),
                                match_itm_nm_contains=str(m.get("match_itm_nm_contains") or m.get("itm_nm_contains") or "").strip(),
                            )
                        )
                    if not mappings:
                        raise ValueError("KOSIS params.mappings has no valid entries")

                    fetched = fetch_kosis_series(q=KosisQuery(query_params=query_params))
                    computed_rows = build_env_base_socio_rows(
                        items=fetched.items,
                        mappings=mappings,
                        admin_code=admin_code,
                        admin_name=admin_name,
                    )
                    if not computed_rows:
                        raise ValueError("KOSIS returned no mappable rows (check mappings/query_params)")

                    ev_json = dict(fetched.evidence_json)
                    ev_json["computed"] = dict(ev_json.get("computed") or {})
                    ev_json["computed"].update(
                        {
                            "admin_code": admin_code,
                            "admin_name": admin_name,
                            "output_rows": computed_rows,
                            "mappings": [m.__dict__ for m in mappings],
                        }
                    )

                ev_id = _evidence_id(req.req_id)
                ev_rel = Path("attachments/evidence/api") / f"{ev_id}_kosis.json"
                ev_abs = (case_dir / ev_rel).resolve()
                ev_abs.parent.mkdir(parents=True, exist_ok=True)

                ev_bytes = kosis_evidence_bytes(ev_json)
                ev_abs.write_bytes(ev_bytes)

                # Fill ENV_BASE_SOCIO
                sheet_rows: list[dict[str, Any]] = []
                for r in computed_rows:
                    year = str(r.get("year") or "").strip()
                    if not year:
                        continue
                    row: dict[str, Any] = {
                        "socio_id": f"SOC-{year}",
                        "admin_code": admin_code,
                        "admin_name": admin_name,
                        "year": year,
                        "data_origin": "OFFICIAL_DB",
                        "src_id": req.src_id or "S-TBD",
                        "evidence_id": ev_id,
                    }
                    for col in ("population_total", "households", "housing_total"):
                        if col in r and r.get(col) is not None:
                            row[col] = r.get(col)
                    sheet_rows.append(row)

                sheet_warn = apply_rows_to_sheet(
                    wb,
                    sheet_name=req.output_sheet.strip() or "ENV_BASE_SOCIO",
                    rows=sheet_rows,
                    merge_strategy=req.merge_strategy or "REPLACE_SHEET",
                    upsert_keys=req.upsert_keys,
                )
                warnings.extend([f"[{req.req_id}] {w}" for w in sheet_warn])

                note_request_url = ""
                note_request_params: Any = {}
                if dataset_keys:
                    if reqs_meta:
                        note_request_url = str(reqs_meta[0].get("url") or "").strip()
                    note_request_params = {"datasets": reqs_meta}
                else:
                    if isinstance(ev_json, dict):
                        req_obj = ev_json.get("request") or {}
                        if isinstance(req_obj, dict):
                            note_request_url = str(req_obj.get("url") or "").strip()
                            note_request_params = req_obj.get("params") or {}

                ev = Evidence(
                    evidence_id=ev_id,
                    evidence_type="기타",
                    title=f"KOSIS:{admin_code}",
                    file_path=str(ev_rel).replace("\\", "/"),
                    used_in=f"DATA_REQUESTS:{req.req_id}",
                    data_origin="OFFICIAL_DB",
                    src_id=req.src_id or "S-TBD",
                    note=_note_json(
                        {
                            "connector": "KOSIS",
                            "req_id": req.req_id,
                            "retrieved_at": _now_iso(),
                            "admin_code": admin_code,
                            "dataset_keys": dataset_keys,
                            "request_url": note_request_url,
                            "request_params": note_request_params,
                            "hash_sha1": _sha1_bytes(ev_bytes),
                        }
                    ),
                )
                append_attachment(wb, ev)
                update_request_run(wb, req_id=req.req_id, evidence_ids=[ev_id])
                evidences.append(ev)

            elif connector == "NIER_WATER":
                attempted = True
                executed += 1

                mgt_no = str(params.get("mgt_no") or params.get("mgtNo") or "").strip()
                if not mgt_no:
                    raise ValueError("Missing params.mgt_no (EIA mgtNo) for NIER_WATER")

                ivstg_spot_nm = str(params.get("ivstg_spot_nm") or params.get("ivstgSpotNm") or "").strip() or None
                res = fetch_eia_ivstg_water_quality(mgt_no=mgt_no, ivstg_spot_nm=ivstg_spot_nm)
                if not res.stations:
                    raise ValueError("NIER_WATER returned no stations")

                ev_id = _evidence_id(req.req_id)
                ev_rel = Path("attachments/evidence/api") / f"{ev_id}_nier_ivstg.json"
                ev_abs = (case_dir / ev_rel).resolve()
                ev_abs.parent.mkdir(parents=True, exist_ok=True)
                ev_bytes = nier_evidence_bytes(res.evidence_json)
                ev_abs.write_bytes(ev_bytes)

                # Distance (best-effort): compute from LOCATION center to ivstg points (EPSG:5179)
                center_lon = loc.get("center_lon")
                center_lat = loc.get("center_lat")
                center_epsg = int(loc.get("epsg") or 4326)
                sx = sy = None
                if center_lon is not None and center_lat is not None:
                    try:
                        from pyproj import CRS, Transformer

                        t = Transformer.from_crs(CRS.from_epsg(center_epsg), CRS.from_epsg(5179), always_xy=True)
                        sx, sy = t.transform(float(center_lon), float(center_lat))
                    except Exception:
                        sx, sy = None, None

                sheet_rows: list[dict[str, Any]] = []
                for idx, st in enumerate(res.stations, start=1):
                    distance_m = None
                    if sx is not None and sy is not None and st.x_5179 is not None and st.y_5179 is not None:
                        try:
                            dx = float(st.x_5179) - float(sx)
                            dy = float(st.y_5179) - float(sy)
                            distance_m = (dx * dx + dy * dy) ** 0.5
                        except Exception:
                            distance_m = None

                    # We output one row per parameter to match ENV_BASE_WATER long format.
                    for param, val in st.metrics_mgL.items():
                        unit = "" if param == "PH" else "mg/L"
                        sheet_rows.append(
                            {
                                "water_id": f"WAT-{idx:03d}",
                                "waterbody_name": st.name or "조사지점",
                                "relation": st.address,
                                "distance_m": distance_m,
                                "parameter": param,
                                "value": val,
                                "unit": unit,
                                "sampling_date": "",
                                "data_origin": "OFFICIAL_DB",
                                "src_id": req.src_id or "S-TBD",
                                "evidence_id": ev_id,
                            }
                        )

                sheet_warn = apply_rows_to_sheet(
                    wb,
                    sheet_name=req.output_sheet.strip() or "ENV_BASE_WATER",
                    rows=sheet_rows,
                    merge_strategy=req.merge_strategy or "REPLACE_SHEET",
                    upsert_keys=req.upsert_keys,
                )
                warnings.extend([f"[{req.req_id}] {w}" for w in sheet_warn])

                ev = Evidence(
                    evidence_id=ev_id,
                    evidence_type="측정원시자료",
                    title=f"NIER_WATER_IVSTG:{mgt_no}",
                    file_path=str(ev_rel).replace("\\", "/"),
                    used_in=f"DATA_REQUESTS:{req.req_id}",
                    data_origin="OFFICIAL_DB",
                    src_id=req.src_id or "S-TBD",
                    note=_note_json(
                        {
                            "connector": "NIER_WATER",
                            "req_id": req.req_id,
                            "retrieved_at": _now_iso(),
                            "mgt_no": mgt_no,
                            "request_url": (res.evidence_json.get("request") or {}).get("url") if isinstance(res.evidence_json, dict) else "",
                            "request_params": (res.evidence_json.get("request") or {}).get("params") if isinstance(res.evidence_json, dict) else {},
                            "hash_sha1": _sha1_bytes(ev_bytes),
                        }
                    ),
                )
                append_attachment(wb, ev)
                update_request_run(wb, req_id=req.req_id, evidence_ids=[ev_id])
                evidences.append(ev)

            elif connector == "KMA_ASOS":
                attempted = True
                executed += 1

                station_id = str(params.get("stn_id") or params.get("stnIds") or "").strip()
                if not station_id:
                    cands = params.get("station_candidates") or []
                    if isinstance(cands, list) and cands:
                        first = cands[0]
                        if isinstance(first, dict):
                            station_id = str(first.get("station_id") or first.get("stn_id") or "").strip()
                        else:
                            station_id = str(first).strip()

                # As a last resort, auto-pick nearest station at runtime (requires coords).
                # Prefer cached catalog CSV, fallback to online fetch.
                auto_pick_err: str | None = None
                if not station_id and loc.get("center_lon") is not None and loc.get("center_lat") is not None:
                    try:
                        from eia_gen.services.data_requests.kma_stations import (
                            fetch_asos_station_catalog,
                            load_asos_station_catalog_csv,
                            pick_nearest_asos_stations,
                            write_asos_station_catalog_csv,
                        )

                        stations_path = wms_layers_config.parent / "stations" / "kma_asos_stations.csv"
                        stations = load_asos_station_catalog_csv(stations_path)
                        if not stations:
                            cat = fetch_asos_station_catalog()
                            stations = cat.stations
                            try:
                                write_asos_station_catalog_csv(stations_path, stations)
                            except Exception:
                                pass

                        cands2 = pick_nearest_asos_stations(
                            center_lon=float(loc.get("center_lon")),
                            center_lat=float(loc.get("center_lat")),
                            stations=stations,
                            top_n=1,
                        )
                        if cands2:
                            station_id = str(cands2[0].get("station_id") or "").strip()
                    except Exception as e:
                        auto_pick_err = str(e)
                if not station_id:
                    raise ValueError(
                        "Missing params_json.stn_id for KMA_ASOS (set stn_id, or run `eia-gen fetch-kma-asos-stations` "
                        "to enable nearest-station auto selection)"
                        + (f" (auto-pick failed: {auto_pick_err})" if auto_pick_err else "")
                    )

                # Default: recent 5 years (daily) — best-effort baseline.
                def _dt(v: Any) -> str:
                    s = str(v or "").strip()
                    return s

                end_dt = _dt(params.get("end_dt"))
                start_dt = _dt(params.get("start_dt"))
                if not end_dt or not start_dt:
                    # KMA ASOS daily is typically available up to *yesterday*.
                    base = datetime.now() - timedelta(days=1)
                    end_dt = end_dt or base.strftime("%Y%m%d")
                    if not start_dt:
                        try:
                            start_dt = base.replace(year=base.year - 5).strftime("%Y%m%d")
                        except ValueError:
                            # Feb 29 edge-case: fall back to ~5 years.
                            start_dt = (base - timedelta(days=365 * 5)).strftime("%Y%m%d")

                stats = fetch_asos_daily_precip_stats(
                    station_id=station_id,
                    start_dt=start_dt,
                    end_dt=end_dt,
                )
                # Keep labels/evidence consistent with any internal date adjustments.
                start_dt = stats.start_dt
                end_dt = stats.end_dt

                ev_id = _evidence_id(req.req_id)
                ev_rel = Path("attachments/evidence/api") / f"{ev_id}_kma_asos.json"
                ev_abs = (case_dir / ev_rel).resolve()
                ev_abs.parent.mkdir(parents=True, exist_ok=True)
                ev_bytes = kma_asos_evidence_bytes(stats.evidence_json)
                ev_abs.write_bytes(ev_bytes)

                # Apply to DRR_HYDRO_RAIN (best-effort: use source_basis as station label).
                src_basis = str(params.get("source_basis") or "").strip() or f"ASOS({station_id}) {start_dt}~{end_dt}"
                rp = params.get("return_period_yr")
                return_period_yr = rp if (rp is not None and str(rp).strip() != "") else None

                rows = []
                if stats.max_1h_rain_mm is not None:
                    rows.append(
                        {
                            "rain_id": f"{ev_id}-1H",
                            "source_basis": src_basis,
                            "return_period_yr": return_period_yr,
                            "duration_hr": 1,
                            "rainfall_mm": round(float(stats.max_1h_rain_mm), 2),
                            "intensity_formula": str(params.get("intensity_formula") or "").strip(),
                            "temporal_dist": str(params.get("temporal_dist") or "").strip(),
                            "data_origin": "OFFICIAL_DB",
                            "src_id": req.src_id or "S-TBD",
                            "evidence_id": ev_id,
                        }
                    )
                if stats.max_24h_rain_mm is not None:
                    rows.append(
                        {
                            "rain_id": f"{ev_id}-24H",
                            "source_basis": src_basis,
                            "return_period_yr": return_period_yr,
                            "duration_hr": 24,
                            "rainfall_mm": round(float(stats.max_24h_rain_mm), 2),
                            "intensity_formula": str(params.get("intensity_formula") or "").strip(),
                            "temporal_dist": str(params.get("temporal_dist") or "").strip(),
                            "data_origin": "OFFICIAL_DB",
                            "src_id": req.src_id or "S-TBD",
                            "evidence_id": ev_id,
                        }
                    )
                if not rows:
                    raise ValueError("KMA_ASOS computed no rainfall stats (missing sumRn in response?)")

                sheet = req.output_sheet.strip() or "DRR_HYDRO_RAIN"
                sheet_warn = apply_rows_to_sheet(
                    wb,
                    sheet_name=sheet,
                    rows=rows,
                    merge_strategy=req.merge_strategy or "REPLACE_SHEET",
                    upsert_keys=req.upsert_keys,
                )
                warnings.extend([f"[{req.req_id}] {w}" for w in sheet_warn])

                ev = Evidence(
                    evidence_id=ev_id,
                    evidence_type="측정원시자료",
                    title=f"KMA_ASOS:{station_id}",
                    file_path=str(ev_rel).replace("\\", "/"),
                    used_in=f"DATA_REQUESTS:{req.req_id}",
                    data_origin="OFFICIAL_DB",
                    src_id=req.src_id or "S-TBD",
                    note=_note_json(
                        {
                            "connector": "KMA_ASOS",
                            "req_id": req.req_id,
                            "retrieved_at": _now_iso(),
                            "station_id": station_id,
                            "start_dt": start_dt,
                            "end_dt": end_dt,
                            "request_url": (stats.evidence_json.get("request") or {}).get("url") if isinstance(stats.evidence_json, dict) else "",
                            "request_params": (stats.evidence_json.get("request") or {}).get("params") if isinstance(stats.evidence_json, dict) else {},
                            "hash_sha1": _sha1_bytes(ev_bytes),
                        }
                    ),
                )
                append_attachment(wb, ev)
                update_request_run(wb, req_id=req.req_id, evidence_ids=[ev_id])
                evidences.append(ev)

            else:
                warnings.append(f"[{req.req_id}] connector not implemented: {req.connector}")
                skipped += 1
                continue

        except Exception as e:
            if not attempted:
                executed += 1
            warnings.append(f"[{req.req_id}] {req.connector} failed: {redact_text(str(e))}")
            continue

    return RunResult(executed=executed, skipped=skipped, warnings=warnings, evidences=evidences)
