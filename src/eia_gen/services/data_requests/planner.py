from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import yaml

from eia_gen.services.data_requests.models import DataRequest
from eia_gen.services.data_requests.xlsx_io import read_sheet_dicts


def _load_yaml(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    return yaml.safe_load(path.read_text(encoding="utf-8")) or {}

def _effective_empty(rows: list[dict[str, Any]], *, any_of: list[str]) -> bool:
    """Treat a sheet as empty when all rows lack meaningful values for key columns."""
    if not rows:
        return True
    keys = [k for k in any_of if k]
    if not keys:
        return len(rows) == 0
    for r in rows:
        for k in keys:
            v = r.get(k)
            if v is None:
                continue
            if isinstance(v, str) and not v.strip():
                continue
            return False
    return True



def plan_default_data_requests(*, wms_layers_config: Path) -> list[DataRequest]:
    """Create a minimal WMS-focused DATA_REQUESTS plan.

    This planner is intentionally conservative:
    - Adds only a small fixed set of WMS layers (eco + flood + landslide).
    - Marks rows disabled when a required API key env var is missing.
    """
    cfg = _load_yaml(wms_layers_config)
    layers: dict[str, Any] = cfg.get("layers") or {}
    providers: dict[str, Any] = cfg.get("providers") or {}

    wanted = [
        # Prefer the explicit data.go.kr ecologyzmp key (legacy alias remains supported).
        "ECO_NATURE_DATAGO_OPEN",
        "FLOOD_TRACE",
        "LANDSLIDE_RISK",
    ]

    out: list[DataRequest] = []
    priority = 50
    for key in wanted:
        layer = layers.get(key) or {}
        provider_name = str(layer.get("provider") or "").strip()
        provider = providers.get(provider_name) or {}

        srs = str(layer.get("srs") or provider.get("default_srs") or "EPSG:3857")
        src_ids = layer.get("default_src_ids") or []
        src_id = str(src_ids[0]).strip() if src_ids else "S-TBD"

        enabled = True
        note = f"WMS layer_key={key} provider={provider_name} srs={srs}"

        # ECO_NATURE layers vary by provider/key approval; keep them opt-in by default.
        # Users can flip `enabled=TRUE` in DATA_REQUESTS when their endpoint/key is verified.
        if key.upper().startswith("ECO_NATURE"):
            enabled = False
            note += " (disabled: opt-in; verify eco WMS endpoint/key first)"

        # Disable when provider requires API key and it's missing.
        auth = provider.get("auth") or {}
        env_var = str(auth.get("env_var") or "").strip() if auth else ""
        if env_var and not os.environ.get(env_var, "").strip():
            enabled = False
            note += f" (disabled: missing env {env_var})"

        params = {
            "layer_key": key,
            "bbox_mode": "AUTO",  # use boundary_file when present, else point+radius
            "srs": srs,
            "width": 2048,
            "height": 2048,
            "radius_m": 1000,
            # Optional: when WMS is blocked (approval/key/outage), set this to a local official image
            # and the runner will use it as a fallback evidence.
            "fallback_file_path": "",
        }

        out.append(
            DataRequest(
                req_id=f"REQ-WMS-{key}",
                enabled=enabled,
                priority=priority,
                connector="WMS",
                purpose="OVERLAY",
                src_id=src_id,
                params_json=json.dumps(params, ensure_ascii=False),
                params=params,
                output_sheet="ATTACHMENTS",
                merge_strategy="APPEND",
                upsert_keys=[],
                # Default to ONCE to avoid duplicating evidences on repeated runs.
                # (User can switch to AUTO when refresh is desired.)
                run_mode="ONCE",
                last_run_at="",
                last_evidence_ids=[],
                note=note,
            )
        )
        priority += 1

    return out


def plan_data_requests_for_workbook(*, wb, wms_layers_config: Path) -> list[DataRequest]:
    """Plan DATA_REQUESTS based on workbook contents (v0).

    Current behavior:
    - Always includes the default WMS evidences plan.
    - Adds AIRKOREA baseline when ENV_BASE_AIR is empty.
    - Adds KMA_ASOS baseline when DRR_HYDRO_RAIN is empty (disabled until stn_id is set).
    - Adds AUTO_GIS zoning breakdown when PARCELS has zoning values and ZONING_BREAKDOWN is empty.
    - Adds KOSIS/NIER_WATER rows as disabled guidance when baseline sheets are empty.
    """
    plan = plan_default_data_requests(wms_layers_config=wms_layers_config)

    # Helper: find a planned request by id (for wiring downstream requests).
    by_req_id = {r.req_id: r for r in plan}

    # LOCATION geocode (best-effort): when coords are missing but address exists.
    loc_rows = read_sheet_dicts(wb, "LOCATION")
    has_address = False
    has_coords = False
    if loc_rows:
        lon = loc_rows[0].get("center_lon")
        lat = loc_rows[0].get("center_lat")
        has_coords = not (lon is None or lat is None or str(lon).strip() == "" or str(lat).strip() == "")

        addr = str(loc_rows[0].get("address_road") or "").strip() or str(loc_rows[0].get("address_jibeon") or "").strip()
        has_address = bool(addr)

    if not has_coords and has_address:
        plan.append(
            DataRequest(
                req_id="REQ-GEOCODE-LOCATION",
                enabled=True,
                priority=10,
                connector="GEOCODE",
                purpose="OVERLAY",
                src_id="S-03",
                params_json=json.dumps({"provider": "AUTO"}, ensure_ascii=False),
                params={"provider": "AUTO"},
                output_sheet="LOCATION",
                merge_strategy="REPLACE_SHEET",
                upsert_keys=[],
                run_mode="ONCE",
                last_run_at="",
                last_evidence_ids=[],
                note="GEOCODE LOCATION.center_lat/center_lon from address (AUTO: VWORLD if key exists else NOMINATIM)",
            )
        )

    # AIRKOREA baseline (best-effort): nearest station → recent mean values
    env_air = read_sheet_dicts(wb, "ENV_BASE_AIR")
    if _effective_empty(env_air, any_of=["station_name", "pollutant", "value_avg"]):
        enabled = True
        note = "AIRKOREA baseline → ENV_BASE_AIR (nearest station, dataTerm=MONTH)"

        # Disable when API key is missing.
        if not os.environ.get("AIRKOREA_API_KEY", "").strip() and not os.environ.get("DATA_GO_KR_SERVICE_KEY", "").strip():
            enabled = False
            note += " (disabled: missing env AIRKOREA_API_KEY or DATA_GO_KR_SERVICE_KEY)"

        if not loc_rows:
            enabled = False
            note += " (disabled: missing LOCATION sheet row)"
        elif not has_coords and not has_address:
            enabled = False
            note += " (disabled: missing LOCATION coords and address)"
        elif not has_coords and has_address:
            note += " (will use GEOCODE first if planned)"

        params = {"data_term": "MONTH", "num_rows": 200}
        plan.append(
            DataRequest(
                req_id="REQ-AIRKOREA-ENV_BASE_AIR",
                enabled=enabled,
                priority=30,
                connector="AIRKOREA",
                purpose="AIR_BASELINE",
                # Recommend users keep a dedicated AirKorea source in sources.yaml (e.g. S-03).
                src_id="S-03",
                params_json=json.dumps(params, ensure_ascii=False),
                params=params,
                output_sheet="ENV_BASE_AIR",
                merge_strategy="REPLACE_SHEET",
                upsert_keys=[],
                run_mode="ONCE",
                last_run_at="",
                last_evidence_ids=[],
                note=note,
            )
        )

    # KMA ASOS rainfall (best-effort): needs station id. We still create the row to guide users.
    drr_rain = read_sheet_dicts(wb, "DRR_HYDRO_RAIN")
    if _effective_empty(drr_rain, any_of=["rainfall_mm", "duration_hr", "return_period_yr"]):
        enabled = True
        note = "KMA_ASOS daily precip → DRR_HYDRO_RAIN (auto-pick nearest ASOS station when possible)"

        has_key = bool(os.environ.get("KMA_API_KEY", "").strip() or os.environ.get("DATA_GO_KR_SERVICE_KEY", "").strip())
        if not has_key:
            enabled = False
            note += " (disabled: missing env KMA_API_KEY or DATA_GO_KR_SERVICE_KEY)"

        # Planner does not call external APIs. When a local station catalog exists and
        # coords are already present, we can prefill candidates. Otherwise the runner
        # will auto-pick at runtime (after GEOCODE if needed).
        station_candidates: list[dict[str, Any]] = []
        station_id = ""
        station_name = ""

        if not loc_rows:
            enabled = False
            note += " (disabled: missing LOCATION sheet row)"
        elif not has_coords and not has_address:
            enabled = False
            note += " (disabled: missing LOCATION coords and address)"
        elif has_coords:
            try:
                from eia_gen.services.data_requests.kma_stations import load_asos_station_catalog_csv, pick_nearest_asos_stations

                stations_path = wms_layers_config.parent / "stations" / "kma_asos_stations.csv"
                stations = load_asos_station_catalog_csv(stations_path)
                if stations:
                    station_candidates = pick_nearest_asos_stations(
                        center_lon=float(lon),
                        center_lat=float(lat),
                        stations=stations,
                        top_n=3,
                    )
                    if station_candidates:
                        station_id = str(station_candidates[0].get("station_id") or "").strip()
                        station_name = str(station_candidates[0].get("station_name") or "").strip()
                else:
                    note += " (hint: run `eia-gen fetch-kma-asos-stations` for faster station selection)"
            except Exception:
                pass

        from datetime import datetime as _dt

        today = _dt.now()
        params = {
            "stn_id": station_id,
            "start_dt": today.replace(year=today.year - 5).strftime("%Y%m%d"),
            "end_dt": today.strftime("%Y%m%d"),
            "station_candidates": station_candidates,
        }
        if station_id:
            label = station_id if not station_name else f"{station_id} {station_name}"
            params["source_basis"] = f"ASOS({label}) {params['start_dt']}~{params['end_dt']}"
        elif not enabled:
            params["stn_id"] = ""
        plan.append(
            DataRequest(
                req_id="REQ-KMA-ASOS-DRR_HYDRO_RAIN",
                enabled=enabled,
                priority=31,
                connector="KMA_ASOS",
                purpose="DRR_RAINFALL",
                # Recommend users keep a dedicated KMA source in sources.yaml.
                src_id="S-03",
                params_json=json.dumps(params, ensure_ascii=False),
                params=params,
                output_sheet="DRR_HYDRO_RAIN",
                merge_strategy="REPLACE_SHEET",
                upsert_keys=[],
                run_mode="ONCE",
                last_run_at="",
                last_evidence_ids=[],
                note=note,
            )
        )

    # KOSIS socio stats (guidance): dataset_key -> SSOT config (`config/kosis_datasets.yaml`).
    env_socio = read_sheet_dicts(wb, "ENV_BASE_SOCIO")
    if _effective_empty(env_socio, any_of=["year", "population_total", "households", "housing_total"]):
        enabled = False
        note = "KOSIS socio stats → ENV_BASE_SOCIO (fill params_json.admin_code then enable; dataset_keys via config/kosis_datasets.yaml)"
        if not os.environ.get("KOSIS_API_KEY", "").strip():
            note += " (disabled: missing env KOSIS_API_KEY)"

        from datetime import datetime as _dt

        today = _dt.now()
        default_end_year = today.year - 1
        default_start_year = default_end_year - 4

        params = {
            "admin_code": "",
            "admin_name": "",
            "start_year": str(default_start_year),
            "end_year": str(default_end_year),
            # Default keys (override per project if needed).
            "dataset_keys": ["POP_HOUSEHOLDS", "HOUSING_STATUS"],
        }
        plan.append(
            DataRequest(
                req_id="REQ-KOSIS-ENV_BASE_SOCIO",
                enabled=enabled,
                priority=32,
                connector="KOSIS",
                purpose="SOCIO_STATS",
                src_id="S-03",
                params_json=json.dumps(params, ensure_ascii=False),
                params=params,
                output_sheet="ENV_BASE_SOCIO",
                merge_strategy="REPLACE_SHEET",
                upsert_keys=[],
                run_mode="ONCE",
                last_run_at="",
                last_evidence_ids=[],
                note=note,
            )
        )

    # NIER water baseline (guidance): requires EIA mgtNo.
    env_water = read_sheet_dicts(wb, "ENV_BASE_WATER")
    if _effective_empty(env_water, any_of=["waterbody_name", "parameter", "value"]):
        enabled = False
        note = "NIER_WATER ivstg → ENV_BASE_WATER (requires params_json.mgt_no then enable)"
        if not os.environ.get("DATA_GO_KR_SERVICE_KEY", "").strip() and not os.environ.get("NIER_WATER_API_KEY", "").strip():
            note += " (disabled: missing env DATA_GO_KR_SERVICE_KEY or NIER_WATER_API_KEY)"

        params = {
            "mgt_no": "",
            "ivstg_spot_nm": "",
        }
        plan.append(
            DataRequest(
                req_id="REQ-NIER-WATER-ENV_BASE_WATER",
                enabled=enabled,
                priority=33,
                connector="NIER_WATER",
                purpose="WATER_BASELINE",
                src_id="S-03",
                params_json=json.dumps(params, ensure_ascii=False),
                params=params,
                output_sheet="ENV_BASE_WATER",
                merge_strategy="REPLACE_SHEET",
                upsert_keys=[],
                run_mode="ONCE",
                last_run_at="",
                last_evidence_ids=[],
                note=note,
            )
        )

    # ZONING_OVERLAY (best-effort): compute O/X + distance from WMS evidence rasters.
    # - Runs after WMS requests so it can read their last_evidence_ids.
    # - Only enables when at least one relevant WMS request is enabled.
    wms_flood = by_req_id.get("REQ-WMS-FLOOD_TRACE")
    wms_slide = by_req_id.get("REQ-WMS-LANDSLIDE_RISK")
    wms_overlay_items: list[dict[str, Any]] = []
    if wms_flood:
        wms_overlay_items.append(
            {
                "overlay_id": "HZ-FLOOD_TRACE",
                "category": "DISASTER",
                "designation_name": "침수흔적도(생활안전지도)",
                "from_req_id": wms_flood.req_id,
                "src_id": wms_flood.src_id,
            }
        )
    if wms_slide:
        wms_overlay_items.append(
            {
                "overlay_id": "HZ-LANDSLIDE_RISK",
                "category": "DISASTER",
                "designation_name": "산사태위험지도(생활안전지도)",
                "from_req_id": wms_slide.req_id,
                "src_id": wms_slide.src_id,
            }
        )

    if wms_overlay_items:
        enabled = bool((wms_flood and wms_flood.enabled) or (wms_slide and wms_slide.enabled))
        note = "AUTO_GIS: ZONING_OVERLAY from WMS evidences (non-transparent pixel heuristic)"
        if not loc_rows:
            enabled = False
            note += " (disabled: missing LOCATION sheet row)"
        elif not has_coords and not has_address:
            enabled = False
            note += " (disabled: missing LOCATION coords and address)"
        elif not has_coords and has_address:
            note += " (will use GEOCODE first if planned)"
        if not enabled:
            note += " (hint: set SAFEMAP_API_KEY to enable FLOOD/LANDSLIDE WMS)"

        params = {
            "operation": "OVERLAY_FROM_WMS_EVIDENCE",
            "items": wms_overlay_items,
            "metric_epsg": 5186,
            "analysis_max_size": 512,
            "alpha_threshold": 10,
            "radius_m": 1000,
        }
        plan.append(
            DataRequest(
                req_id="REQ-AUTO-GIS-ZONING-OVERLAY-WMS",
                enabled=enabled,
                priority=60,
                connector="AUTO_GIS",
                purpose="OVERLAY",
                # Prefer the underlying WMS source id to avoid unknown-source warnings.
                src_id=(
                    (wms_flood.src_id if wms_flood else "")
                    or (wms_slide.src_id if wms_slide else "")
                    or "S-TBD"
                ),
                params_json=json.dumps(params, ensure_ascii=False),
                params=params,
                output_sheet="ZONING_OVERLAY",
                merge_strategy="UPSERT_KEYS",
                upsert_keys=["overlay_id"],
                run_mode="ONCE",
                last_run_at="",
                last_evidence_ids=[],
                note=note,
            )
        )

    zoning_breakdown = read_sheet_dicts(wb, "ZONING_BREAKDOWN")
    parcels = read_sheet_dicts(wb, "PARCELS")
    has_zoning = any(str(r.get("zoning") or "").strip() for r in parcels)
    if _effective_empty(zoning_breakdown, any_of=["zoning", "area_m2"]) and has_zoning:
        src_candidates = [str(r.get("src_id") or "").strip() for r in parcels if str(r.get("src_id") or "").strip()]
        src_id = src_candidates[0] if src_candidates else "S-TBD"

        params = {"operation": "ZONING_BREAKDOWN_FROM_PARCELS"}
        plan.append(
            DataRequest(
                req_id="REQ-AUTO-GIS-ZONING-BREAKDOWN",
                enabled=True,
                priority=40,
                connector="AUTO_GIS",
                purpose="OVERLAY",
                src_id=src_id,
                params_json=json.dumps(params, ensure_ascii=False),
                params=params,
                output_sheet="ZONING_BREAKDOWN",
                merge_strategy="REPLACE_SHEET",
                upsert_keys=[],
                run_mode="ONCE",
                last_run_at="",
                last_evidence_ids=[],
                note="AUTO_GIS: aggregate PARCELS.zoning into ZONING_BREAKDOWN (evidence: calc csv)",
            )
        )

    plan.sort(key=lambda x: (x.priority, x.req_id))
    return plan
