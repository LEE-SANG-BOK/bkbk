from __future__ import annotations

import csv
import json
import os
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import httpx

from eia_gen.services.data_requests.data_go_kr import build_url
from pyproj import Geod


KMA_ASOS_STATIONS_URL = "http://apis.data.go.kr/1360000/AsosInfoService/getAsosStnInfo"


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _service_key() -> str:
    # Prefer a dedicated key env var, fallback to shared data.go.kr service key.
    for env in ("KMA_API_KEY", "DATA_GO_KR_SERVICE_KEY"):
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


def _extract_items(resp: dict[str, Any]) -> list[dict[str, Any]]:
    body = ((resp or {}).get("response") or {}).get("body") or {}
    items = body.get("items")
    if isinstance(items, dict) and "item" in items:
        items = items.get("item")
    if isinstance(items, list):
        return [dict(x or {}) for x in items]
    if isinstance(items, dict):
        return [dict(items)]
    return []


@dataclass(frozen=True)
class AsosStation:
    station_id: str
    station_name: str
    lat: float
    lon: float


@dataclass(frozen=True)
class AsosStationCatalog:
    stations: list[AsosStation]
    evidence_json: dict[str, Any]


def fetch_asos_station_catalog(*, timeout_sec: int = 25, num_rows: int = 100, max_pages: int = 50) -> AsosStationCatalog:
    """Fetch KMA ASOS station catalog (id/name/lat/lon) with pagination.

    Notes:
    - `num_rows` is treated as page size.
    - We retry a few times on transient 5xx errors.
    - Evidence excludes API keys.
    """
    key = _service_key()
    if not key:
        raise ValueError("Missing KMA API key (KMA_API_KEY or DATA_GO_KR_SERVICE_KEY)")

    page_size = max(1, int(num_rows))
    all_items: list[dict[str, Any]] = []
    pages: list[dict[str, Any]] = []

    def _total_count(resp: dict[str, Any]) -> int | None:
        body = ((resp or {}).get("response") or {}).get("body") or {}
        v = body.get("totalCount")
        try:
            if v is None or str(v).strip() == "":
                return None
            return int(str(v).strip())
        except Exception:
            return None

    def _result_code(resp: dict[str, Any]) -> str:
        header = ((resp or {}).get("response") or {}).get("header") or {}
        return str(header.get("resultCode") or "").strip()

    def _result_msg(resp: dict[str, Any]) -> str:
        header = ((resp or {}).get("response") or {}).get("header") or {}
        return str(header.get("resultMsg") or "").strip()

    with httpx.Client(timeout=timeout_sec, follow_redirects=True) as client:
        total: int | None = None
        page_no = 1
        while page_no <= max_pages:
            params = {
                "pageNo": page_no,
                "numOfRows": page_size,
                "dataType": "JSON",
            }
            url = build_url(base_url=KMA_ASOS_STATIONS_URL, service_key=key, params=params, key_param="serviceKey")

            # best-effort retries for transient server errors
            resp: dict[str, Any] | None = None
            last_status: int | None = None
            for attempt in range(4):
                try:
                    r = client.get(url)
                    last_status = getattr(r, "status_code", None)
                    r.raise_for_status()
                    resp = r.json()
                    break
                except httpx.HTTPStatusError as e:
                    last_status = e.response.status_code
                    if 500 <= int(last_status or 0) < 600 and attempt < 3:
                        time.sleep(0.6 * (2**attempt))
                        continue
                    raise ValueError(f"KMA ASOS station catalog fetch failed: HTTP {e.response.status_code}") from None
                except httpx.HTTPError:
                    if attempt < 3:
                        time.sleep(0.6 * (2**attempt))
                        continue
                    raise ValueError("KMA ASOS station catalog fetch failed") from None

            if not isinstance(resp, dict):
                raise ValueError("KMA ASOS station catalog returned invalid JSON")

            code = _result_code(resp)
            if code and code not in {"00", "0", "200"}:
                raise ValueError(f"KMA ASOS station catalog API error: resultCode={code} resultMsg={_result_msg(resp)}")

            items = _extract_items(resp)
            if not items:
                break

            all_items.extend(items)
            pages.append({"pageNo": page_no, "numOfRows": page_size, "http_status": last_status})

            total = total if total is not None else _total_count(resp)
            if total is not None and page_no * page_size >= total:
                break
            if len(items) < page_size:
                break

            page_no += 1

    if not all_items:
        raise ValueError("KMA ASOS station info returned no items")

    stations: list[AsosStation] = []
    for it in all_items:
        sid = str(it.get("stnId") or it.get("stationId") or it.get("stn_id") or "").strip()
        name = str(it.get("stnNm") or it.get("stationName") or it.get("stn_nm") or "").strip()
        lat = _as_float(it.get("lat"))
        lon = _as_float(it.get("lon"))
        if not sid or lat is None or lon is None:
            continue
        stations.append(AsosStation(station_id=sid, station_name=name, lat=float(lat), lon=float(lon)))

    if not stations:
        raise ValueError("KMA ASOS station info had no usable lat/lon fields")

    evidence = {
        "generated_at": _now_iso(),
        "request": {
            "url": KMA_ASOS_STATIONS_URL,
            "params": {
                "dataType": "JSON",
                "numOfRows": page_size,
                "max_pages": max_pages,
            },
            "pages": pages,
        },
        "response": {
            "item_count": len(all_items),
            "total_count": total,
        },
        "parsed": {"station_count": len(stations)},
    }

    return AsosStationCatalog(stations=stations, evidence_json=evidence)


def load_asos_station_catalog_csv(path: Path) -> list[AsosStation]:
    if not path.exists():
        return []
    rows: list[AsosStation] = []
    with path.open("r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for r in reader:
            sid = str(r.get("station_id") or "").strip()
            name = str(r.get("station_name") or "").strip()
            lat = _as_float(r.get("lat"))
            lon = _as_float(r.get("lon"))
            if not sid or lat is None or lon is None:
                continue
            rows.append(AsosStation(station_id=sid, station_name=name, lat=float(lat), lon=float(lon)))
    return rows


def write_asos_station_catalog_csv(path: Path, stations: list[AsosStation]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["station_id", "station_name", "lat", "lon"])
        writer.writeheader()
        for s in sorted(stations, key=lambda x: (x.station_id, x.station_name)):
            writer.writerow(
                {
                    "station_id": s.station_id,
                    "station_name": s.station_name,
                    "lat": f"{s.lat:.8f}",
                    "lon": f"{s.lon:.8f}",
                }
            )


def evidence_bytes(evidence: dict[str, Any]) -> bytes:
    return json.dumps(evidence, ensure_ascii=False, indent=2).encode("utf-8")


def pick_nearest_asos_stations(
    *,
    center_lon: float,
    center_lat: float,
    stations: list[AsosStation],
    top_n: int = 3,
) -> list[dict[str, Any]]:
    """Pick nearest ASOS stations using geodesic distance (WGS84)."""
    if not stations:
        return []

    geod = Geod(ellps="WGS84")
    scored: list[tuple[float, AsosStation]] = []
    for s in stations:
        try:
            _, _, dist_m = geod.inv(center_lon, center_lat, s.lon, s.lat)
            if dist_m is None:
                continue
            scored.append((float(dist_m), s))
        except Exception:
            continue

    scored.sort(key=lambda x: x[0])
    out: list[dict[str, Any]] = []
    for dist_m, s in scored[: max(1, int(top_n))]:
        out.append(
            {
                "station_id": s.station_id,
                "station_name": s.station_name,
                "distance_km": round(dist_m / 1000.0, 3),
            }
        )
    return out

