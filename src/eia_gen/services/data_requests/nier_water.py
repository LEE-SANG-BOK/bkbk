from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import datetime
from typing import Any

import httpx

from eia_gen.services.data_requests.data_go_kr import build_url


NIER_EIA_WATER_IVSTG_URL = "https://apis.data.go.kr/1480523/WaterqualityServices/getIvstg"


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _service_key() -> str:
    # Prefer a dedicated key env var, fallback to shared data.go.kr service key.
    for env in ("NIER_WATER_API_KEY", "DATA_GO_KR_SERVICE_KEY"):
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


def _as_list(v: Any) -> list[dict[str, Any]]:
    if v is None:
        return []
    if isinstance(v, list):
        return [dict(x or {}) for x in v if isinstance(x, dict)]
    if isinstance(v, dict):
        return [dict(v)]
    return []


@dataclass(frozen=True)
class NierIvstgStation:
    name: str
    address: str
    x_5179: float | None
    y_5179: float | None
    metrics_mgL: dict[str, float]


@dataclass(frozen=True)
class NierIvstgResult:
    mgt_no: str
    stations: list[NierIvstgStation]
    evidence_json: dict[str, Any]


def fetch_eia_ivstg_water_quality(
    *,
    mgt_no: str,
    ivstg_spot_nm: str | None = None,
    timeout_sec: int = 25,
) -> NierIvstgResult:
    """Fetch NIER "환경영향평가 수질정보" (ivstg) and extract core water quality metrics.

    Notes:
    - This API requires `mgtNo` (환경영향평가 사업코드). For non-EIA projects, prefer
      manual baselines or other water monitoring APIs.
    - We only extract a small common subset (BOD/COD/SS/TN/TP/DO/PH) to fit ENV_BASE_WATER.
    """
    key = _service_key()
    if not key:
        raise ValueError("Missing NIER water API key (NIER_WATER_API_KEY or DATA_GO_KR_SERVICE_KEY)")

    mgt = str(mgt_no or "").strip()
    if not mgt:
        raise ValueError("Missing mgt_no (params_json.mgt_no)")

    params: dict[str, Any] = {
        "mgtNo": mgt,
        "type": "json",
    }
    if ivstg_spot_nm and str(ivstg_spot_nm).strip():
        params["ivstgSpotNm"] = str(ivstg_spot_nm).strip()

    url = build_url(base_url=NIER_EIA_WATER_IVSTG_URL, service_key=key, params=params, key_param="ServiceKey")

    with httpx.Client(timeout=timeout_sec, follow_redirects=True) as client:
        try:
            r = client.get(url)
            r.raise_for_status()
            resp = r.json()
        except httpx.HTTPStatusError as e:
            raise ValueError(f"NIER ivstg API failed: HTTP {e.response.status_code}") from None
        except httpx.HTTPError:
            raise ValueError("NIER ivstg API failed") from None

    outer = (resp or {}).get("response") if isinstance(resp, dict) else None
    if not isinstance(outer, dict):
        outer = resp if isinstance(resp, dict) else {}

    header = outer.get("header") if isinstance(outer, dict) else None
    if isinstance(header, dict):
        code = str(header.get("resultCode") or "").strip()
        if code and code not in {"00", "0", "200"}:
            raise ValueError(f"NIER ivstg API error: resultCode={code} resultMsg={header.get('resultMsg')}")

    body = outer.get("body") if isinstance(outer, dict) else None
    if not isinstance(body, dict):
        body = {}

    stations: list[NierIvstgStation] = []
    for gb in _as_list(body.get("ivstgGbs")):
        for s in _as_list(gb.get("ivstgs")):
            name = str(s.get("ivstgSpotNm") or "").strip()
            address = str(s.get("adres") or "").strip()
            x = _as_float(s.get("xcnts"))
            y = _as_float(s.get("ydnts"))
            odrs = s.get("odrs") if isinstance(s.get("odrs"), dict) else {}

            metrics: dict[str, float] = {}
            mapping = {
                "BOD": "bodVal",
                "COD": "codVal",
                "SS": "ssVal",
                "TN": "tnVal",
                "TP": "tpVal",
                "DO": "doVal",
                "PH": "phVal",
            }
            for k, src_key in mapping.items():
                v = _as_float((odrs or {}).get(src_key))
                if v is None:
                    continue
                metrics[k] = float(v)

            # Skip stations with no usable metrics.
            if not metrics and not name and not address:
                continue

            stations.append(
                NierIvstgStation(
                    name=name,
                    address=address,
                    x_5179=x,
                    y_5179=y,
                    metrics_mgL=metrics,
                )
            )

    evidence = {
        "generated_at": _now_iso(),
        "request": {
            "url": NIER_EIA_WATER_IVSTG_URL,
            "params": {k: v for k, v in params.items() if k != "ServiceKey"},
        },
        "response": resp,
        "computed": {
            "mgt_no": mgt,
            "station_count": len(stations),
            "fields": ["BOD", "COD", "SS", "TN", "TP", "DO", "PH"],
        },
    }

    return NierIvstgResult(mgt_no=mgt, stations=stations, evidence_json=evidence)


def evidence_bytes(evidence: dict[str, Any]) -> bytes:
    return json.dumps(evidence, ensure_ascii=False, indent=2).encode("utf-8")

