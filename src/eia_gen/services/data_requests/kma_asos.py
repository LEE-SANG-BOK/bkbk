from __future__ import annotations

import json
import os
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Any

import httpx

from eia_gen.services.data_requests.data_go_kr import build_url


KMA_ASOS_DAILY_URL = "http://apis.data.go.kr/1360000/AsosDalyInfoService/getWthrDataList"


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


def _result_code(resp: dict[str, Any]) -> str:
    header = ((resp or {}).get("response") or {}).get("header") or {}
    return str(header.get("resultCode") or "").strip()


def _result_msg(resp: dict[str, Any]) -> str:
    header = ((resp or {}).get("response") or {}).get("header") or {}
    return str(header.get("resultMsg") or "").strip()


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
class AsosPrecipStats:
    station_id: str
    start_dt: str  # YYYYMMDD
    end_dt: str  # YYYYMMDD
    total_rain_mm: float | None
    max_24h_rain_mm: float | None
    max_1h_rain_mm: float | None
    evidence_json: dict[str, Any]


def fetch_asos_daily_precip_stats(
    *,
    station_id: str,
    start_dt: str,
    end_dt: str,
    timeout_sec: int = 25,
    num_rows: int = 9999,
) -> AsosPrecipStats:
    """Fetch ASOS daily data and compute precipitation summary (best-effort).

    Notes:
    - Uses KMA ASOS daily API (data.go.kr).
    - Computes max 24h rainfall from daily sum (sumRn).
    - Attempts max 1h rainfall from common fields when present (best-effort).
    """
    key = _service_key()
    if not key:
        raise ValueError("Missing KMA API key (KMA_API_KEY or DATA_GO_KR_SERVICE_KEY)")

    requested_start_dt = str(start_dt or "").strip()
    requested_end_dt = str(end_dt or "").strip()

    stn = str(station_id or "").strip()
    if not stn:
        raise ValueError("Missing station_id (params_json.stn_id)")

    # KMA ASOS daily API typically serves data up to *yesterday*.
    # If a caller asks for today, the API may return resultCode=99 ("전날 자료까지 제공...").
    # To keep AUTO runs stable, clamp end_dt to yesterday when it exceeds yesterday.
    def _parse_ymd(s: str) -> datetime | None:
        ss = str(s or "").strip()
        if not ss:
            return None
        try:
            return datetime.strptime(ss, "%Y%m%d")
        except Exception:
            return None

    max_end_dt = (datetime.now() - timedelta(days=1)).strftime("%Y%m%d")
    end_parsed = _parse_ymd(requested_end_dt)
    max_end_parsed = _parse_ymd(max_end_dt)
    if end_parsed is not None and max_end_parsed is not None and end_parsed > max_end_parsed:
        end_dt = max_end_dt
    else:
        end_dt = requested_end_dt
    start_dt = requested_start_dt

    # Basic guard: ensure start<=end when both parse.
    end_parsed2 = _parse_ymd(end_dt)
    start_parsed2 = _parse_ymd(start_dt)
    if start_parsed2 is not None and end_parsed2 is not None and start_parsed2 > end_parsed2:
        raise ValueError(f"KMA ASOS daily API invalid date range: start_dt={start_dt} end_dt={end_dt}")

    # NOTE: This API can reject large date ranges with resultCode=99, even if pagination exists.
    # To keep requests reliable, we chunk the date range into <= 1,000-day blocks.
    page_size = max(1, min(int(num_rows), 999))
    all_items: list[dict[str, Any]] = []
    last_resp: dict[str, Any] | None = None
    ranges: list[tuple[str, str]] = []
    range_runs: list[dict[str, Any]] = []

    def _total_count(resp: dict[str, Any]) -> int | None:
        body = ((resp or {}).get("response") or {}).get("body") or {}
        v = body.get("totalCount")
        try:
            if v is None or str(v).strip() == "":
                return None
            return int(str(v).strip())
        except Exception:
            return None

    # Prepare chunk ranges.
    if start_parsed2 is not None and end_parsed2 is not None:
        cur = start_parsed2
        while cur <= end_parsed2:
            chunk_end = min(cur + timedelta(days=999), end_parsed2)  # inclusive <= 1000 days
            ranges.append((cur.strftime("%Y%m%d"), chunk_end.strftime("%Y%m%d")))
            cur = chunk_end + timedelta(days=1)
    else:
        ranges.append((start_dt, end_dt))

    def _fetch_range(
        client: httpx.Client,
        *,
        range_start_dt: str,
        range_end_dt: str,
    ) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int | None, dict[str, Any] | None]:
        pages: list[dict[str, Any]] = []
        items_all: list[dict[str, Any]] = []
        total: int | None = None
        last_resp: dict[str, Any] | None = None

        page_no = 1
        while page_no <= 50:
            params = {
                "pageNo": page_no,
                "numOfRows": page_size,
                "dataType": "JSON",
                "dataCd": "ASOS",
                "dateCd": "DAY",
                "startDt": range_start_dt,
                "endDt": range_end_dt,
                "stnIds": stn,
            }

            url = build_url(base_url=KMA_ASOS_DAILY_URL, service_key=key, params=params, key_param="serviceKey")

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
                    raise ValueError(f"KMA ASOS daily API failed: HTTP {e.response.status_code}") from None
                except httpx.HTTPError:
                    if attempt < 3:
                        time.sleep(0.6 * (2**attempt))
                        continue
                    raise ValueError("KMA ASOS daily API failed") from None

            if not isinstance(resp, dict):
                raise ValueError("KMA ASOS daily API returned invalid JSON")

            last_resp = resp
            code = _result_code(resp)
            if code and code not in {"00", "0", "200"}:
                raise ValueError(f"KMA ASOS daily API error: resultCode={code} resultMsg={_result_msg(resp)}")

            chunk = _extract_items(resp)
            if not chunk:
                break

            items_all.extend(chunk)
            pages.append({"pageNo": page_no, "numOfRows": page_size, "http_status": last_status})

            total = total if total is not None else _total_count(resp)
            if total is not None and page_no * page_size >= total:
                break
            if len(chunk) < page_size:
                break
            page_no += 1

        return items_all, pages, total, last_resp

    with httpx.Client(timeout=timeout_sec, follow_redirects=True) as client:
        for range_start_dt, range_end_dt in ranges:
            items_chunk, pages, total, last = _fetch_range(
                client,
                range_start_dt=range_start_dt,
                range_end_dt=range_end_dt,
            )
            if items_chunk:
                all_items.extend(items_chunk)
            last_resp = last or last_resp
            range_runs.append(
                {
                    "startDt": range_start_dt,
                    "endDt": range_end_dt,
                    "total_count": total,
                    "pages": pages,
                }
            )

    if not all_items:
        raise ValueError("KMA ASOS daily API returned no items")

    items = all_items

    # Daily total rainfall typically in sumRn
    totals: list[float] = []
    max1h_candidates: list[float] = []
    max1h_keys = [
        "maxRn",  # some datasets
        "maxRnHrmt",  # maximum hourly rainfall time (sometimes paired)
        "maxRn1hr",
        "max1hrRn",
        "maxRnHr",
        "maxRnHr1",
    ]

    for it in items:
        v = _as_float(it.get("sumRn"))
        if v is not None:
            totals.append(v)

        for k in max1h_keys:
            vv = _as_float(it.get(k))
            if vv is not None:
                max1h_candidates.append(vv)

    total_rain = sum(totals) if totals else None
    max_24h = max(totals) if totals else None
    max_1h = max(max1h_candidates) if max1h_candidates else None

    evidence = {
        "generated_at": _now_iso(),
        "request": {
            "url": KMA_ASOS_DAILY_URL,
            "params": {"stnIds": stn, "startDt": start_dt, "endDt": end_dt, "page_size": page_size},
            "requested": {"startDt": requested_start_dt, "endDt": requested_end_dt},
            "ranges": range_runs,
        },
        "response": {"items_count": len(items), "last_page": last_resp},
        "computed": {
            "station_id": stn,
            "start_dt": start_dt,
            "end_dt": end_dt,
            "total_rain_mm": total_rain,
            "max_24h_rain_mm": max_24h,
            "max_1h_rain_mm": max_1h,
        },
    }

    return AsosPrecipStats(
        station_id=stn,
        start_dt=start_dt,
        end_dt=end_dt,
        total_rain_mm=total_rain,
        max_24h_rain_mm=max_24h,
        max_1h_rain_mm=max_1h,
        evidence_json=evidence,
    )


def evidence_bytes(evidence: dict[str, Any]) -> bytes:
    return json.dumps(evidence, ensure_ascii=False, indent=2).encode("utf-8")
