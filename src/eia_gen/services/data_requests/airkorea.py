from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import datetime
from typing import Any

import httpx
from pyproj import CRS, Transformer

from eia_gen.services.data_requests.sanitize import strip_secrets_from_params
from eia_gen.services.data_requests.data_go_kr import build_url


AIRKOREA_BASE = "http://apis.data.go.kr/B552584"


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _service_key() -> str:
    # Prefer a dedicated key env var, fallback to shared data.go.kr service key.
    for env in ("AIRKOREA_API_KEY", "DATA_GO_KR_SERVICE_KEY"):
        v = os.environ.get(env, "").strip()
        if v:
            return v
    return ""


def _as_float(v: Any) -> float | None:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s or s in {"-", "NA", "N/A"}:
        return None
    try:
        return float(s)
    except Exception:
        return None


def _parse_dt(s: str) -> datetime | None:
    s = (s or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    return None


def _to_tm5181(lon: float, lat: float) -> tuple[float, float]:
    # AirKorea `getNearbyMsrstnList` expects "TM" coordinates (commonly used by Kakao/Daum),
    # which correspond to EPSG:5181 in most practical conversions.
    t = Transformer.from_crs(CRS.from_epsg(4326), CRS.from_epsg(5181), always_xy=True)
    x, y = t.transform(lon, lat)
    return float(x), float(y)


@dataclass(frozen=True)
class AirBaseline:
    station_name: str
    station_distance_km: float | None
    period_start: str
    period_end: str
    values: dict[str, float]  # PM10/PM25/O3
    evidence_json: dict[str, Any]


def fetch_air_baseline(
    *,
    center_lon: float,
    center_lat: float,
    station_name_override: str = "",
    data_term: str = "MONTH",
    num_rows: int = 200,
    timeout_sec: int = 20,
) -> AirBaseline:
    """Fetch a best-effort baseline from AirKorea public API.

    Strategy (v0):
    - Resolve nearest measurement station (getNearbyMsrstnList) unless station_name_override is set.
    - Pull recent observations (getMsrstnAcctoRltmMesureDnsty) with dataTerm=MONTH (default).
    - Compute mean values for PM10 / PM2.5 / O3 across returned items.

    Notes:
    - This is not an "annual official report" aggregator; it is a reproducible automatic baseline
      with stored evidence to be refined later.
    """
    key = _service_key()
    if not key:
        raise ValueError("Missing AirKorea API key (AIRKOREA_API_KEY or DATA_GO_KR_SERVICE_KEY)")

    station_name = station_name_override.strip()
    station_distance_km: float | None = None

    station_req: dict[str, Any] = {}
    station_resp: dict[str, Any] | None = None
    if not station_name:
        tm_x, tm_y = _to_tm5181(center_lon, center_lat)
        station_params = {
            "serviceKey": key,
            "returnType": "json",
            "tmX": f"{tm_x:.3f}",
            "tmY": f"{tm_y:.3f}",
        }
        station_req = {
            "url": f"{AIRKOREA_BASE}/MsrstnInfoInqireSvc/getNearbyMsrstnList",
            "params": strip_secrets_from_params(station_params),
        }
        with httpx.Client(timeout=timeout_sec, follow_redirects=True) as client:
            try:
                url = build_url(base_url=station_req["url"], service_key=key, params=station_req["params"], key_param="serviceKey")
                r = client.get(url)
                r.raise_for_status()
                station_resp = r.json()
            except httpx.HTTPStatusError as e:
                raise ValueError(f"AirKorea station request failed: HTTP {e.response.status_code}") from None
            except httpx.HTTPError:
                raise ValueError("AirKorea station request failed") from None

        items = ((station_resp or {}).get("response") or {}).get("body") or {}
        items = items.get("items") or []
        if not items:
            raise ValueError("AirKorea: getNearbyMsrstnList returned no stations")
        first = items[0] or {}
        station_name = str(first.get("stationName") or "").strip()
        station_distance_km = _as_float(first.get("tm"))  # typically km
        if not station_name:
            raise ValueError("AirKorea: stationName missing in nearest station response")

    meas_params = {
        "serviceKey": key,
        "returnType": "json",
        "numOfRows": int(num_rows),
        "pageNo": 1,
        "stationName": station_name,
        "dataTerm": data_term,
        "ver": "1.3",
    }
    meas_req = {
        "url": f"{AIRKOREA_BASE}/ArpltnInforInqireSvc/getMsrstnAcctoRltmMesureDnsty",
        "params": strip_secrets_from_params(meas_params),
    }

    with httpx.Client(timeout=timeout_sec, follow_redirects=True) as client:
        try:
            url = build_url(base_url=meas_req["url"], service_key=key, params=meas_req["params"], key_param="serviceKey")
            r = client.get(url)
            r.raise_for_status()
            meas_resp = r.json()
        except httpx.HTTPStatusError as e:
            raise ValueError(f"AirKorea measurement request failed: HTTP {e.response.status_code}") from None
        except httpx.HTTPError:
            raise ValueError("AirKorea measurement request failed") from None

    body = ((meas_resp or {}).get("response") or {}).get("body") or {}
    items = body.get("items") or []
    if not isinstance(items, list) or not items:
        raise ValueError("AirKorea: measurement API returned no items")

    # compute means
    def _mean(key: str) -> float | None:
        vals: list[float] = []
        for it in items:
            v = _as_float((it or {}).get(key))
            if v is None:
                continue
            vals.append(v)
        if not vals:
            return None
        return sum(vals) / float(len(vals))

    pm10 = _mean("pm10Value")
    pm25 = _mean("pm25Value")
    o3 = _mean("o3Value")

    # period range from items' dataTime
    dts = [_parse_dt(str((it or {}).get("dataTime") or "")) for it in items]
    dts = [d for d in dts if d is not None]
    if dts:
        d0 = min(dts)
        d1 = max(dts)
        period_start = d0.strftime("%Y-%m-%d")
        period_end = d1.strftime("%Y-%m-%d")
    else:
        period_start = ""
        period_end = ""

    values: dict[str, float] = {}
    if pm10 is not None:
        values["PM10"] = float(pm10)
    if pm25 is not None:
        values["PM2.5"] = float(pm25)
    if o3 is not None:
        values["O3"] = float(o3)

    evidence = {
        "generated_at": _now_iso(),
        "inputs": {
            "center_lon": center_lon,
            "center_lat": center_lat,
            "station_name_override": station_name_override,
            "data_term": data_term,
            "num_rows": num_rows,
        },
        "station_request": station_req,
        "station_response": station_resp,
        "measure_request": meas_req,
        "measure_response": meas_resp,
        "computed": {
            "station_name": station_name,
            "station_distance_km": station_distance_km,
            "period_start": period_start,
            "period_end": period_end,
            "values": values,
        },
    }

    return AirBaseline(
        station_name=station_name,
        station_distance_km=station_distance_km,
        period_start=period_start,
        period_end=period_end,
        values=values,
        evidence_json=evidence,
    )


def evidence_bytes(evidence: dict[str, Any]) -> bytes:
    return json.dumps(evidence, ensure_ascii=False, indent=2).encode("utf-8")
