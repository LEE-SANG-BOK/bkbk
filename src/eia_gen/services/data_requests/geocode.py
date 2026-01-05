from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import datetime
from typing import Any

import httpx


NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
VWORLD_GEOCODE_URL = "https://api.vworld.kr/req/address"


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


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


@dataclass(frozen=True)
class GeocodeResult:
    provider: str
    address: str
    lat: float
    lon: float
    evidence_json: dict[str, Any]


def _vworld_key() -> str:
    return os.environ.get("VWORLD_API_KEY", "").strip()


def geocode_address(*, address: str, provider: str = "AUTO", timeout_sec: int = 20) -> GeocodeResult:
    """Geocode an address into WGS84 lon/lat.

    Supported providers:
    - AUTO: prefer VWORLD when `VWORLD_API_KEY` exists, else NOMINATIM
    - VWORLD: VWorld address geocoder (requires `VWORLD_API_KEY`)
    - NOMINATIM: OpenStreetMap Nominatim (no key; best-effort)
    """
    addr = str(address or "").strip()
    if not addr:
        raise ValueError("Missing address")

    prov = str(provider or "AUTO").strip().upper()
    if prov == "AUTO":
        prov = "VWORLD" if _vworld_key() else "NOMINATIM"

    if prov == "VWORLD":
        key = _vworld_key()
        if not key:
            raise ValueError("Missing VWORLD_API_KEY")

        params = {
            "service": "address",
            "request": "getcoord",
            "version": "2.0",
            "crs": "epsg:4326",
            "address": addr,
            "format": "json",
            "type": "road",
            "key": key,
        }
        with httpx.Client(timeout=timeout_sec, follow_redirects=True) as client:
            r = client.get(VWORLD_GEOCODE_URL, params=params)
            r.raise_for_status()
            resp = r.json()

        # best-effort parse
        point = (((resp or {}).get("response") or {}).get("result") or {}).get("point") or {}
        lon = _as_float(point.get("x"))
        lat = _as_float(point.get("y"))
        if lon is None or lat is None:
            raise ValueError("VWorld geocode returned no point")

        evidence = {
            "generated_at": _now_iso(),
            "provider": "VWORLD",
            "request": {"url": VWORLD_GEOCODE_URL, "params": {k: v for k, v in params.items() if k != "key"}},
            "response": resp,
            "result": {"lat": lat, "lon": lon},
        }
        return GeocodeResult(provider="VWORLD", address=addr, lat=float(lat), lon=float(lon), evidence_json=evidence)

    if prov == "NOMINATIM":
        params = {"q": addr, "format": "jsonv2", "limit": 1}
        headers = {
            # Nominatim requires a valid User-Agent.
            "User-Agent": "eia-gen/0.1 (local, for EIA/DIA drafting)",
        }
        with httpx.Client(timeout=timeout_sec, follow_redirects=True, headers=headers) as client:
            r = client.get(NOMINATIM_URL, params=params)
            r.raise_for_status()
            resp = r.json()

        if not isinstance(resp, list) or not resp:
            raise ValueError("Nominatim returned no results")
        first = resp[0] or {}
        lat = _as_float(first.get("lat"))
        lon = _as_float(first.get("lon"))
        if lon is None or lat is None:
            raise ValueError("Nominatim result missing lat/lon")

        evidence = {
            "generated_at": _now_iso(),
            "provider": "NOMINATIM",
            "request": {"url": NOMINATIM_URL, "params": params},
            "response": resp,
            "result": {"lat": lat, "lon": lon},
        }
        return GeocodeResult(provider="NOMINATIM", address=addr, lat=float(lat), lon=float(lon), evidence_json=evidence)

    raise ValueError(f"Unsupported provider: {provider}")


def evidence_bytes(evidence: dict[str, Any]) -> bytes:
    return json.dumps(evidence, ensure_ascii=False, indent=2).encode("utf-8")

