from __future__ import annotations

import hashlib
import json
import io
import re
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import httpx
import yaml
from pyproj import CRS, Transformer

from eia_gen.services.data_requests.sanitize import redact_text
from eia_gen.services.data_requests.sanitize import strip_secrets_from_params


@dataclass(frozen=True)
class WmsFetchResult:
    bytes_: bytes
    content_type: str
    request_url: str
    request_params: dict[str, Any]
    cache_hit: bool


def _sha1(s: str) -> str:
    return hashlib.sha1(s.encode("utf-8")).hexdigest()


def _now() -> datetime:
    return datetime.now()


_HTML_ALERT_RE = re.compile(r"(?i)alert\(\s*['\"]([^'\"]+)['\"]\s*\)")
_HTML_TITLE_RE = re.compile(r"(?is)<title>\s*([^<]{1,120})\s*</title>")


def _extract_html_error_hint(text: str) -> str:
    """Best-effort extract a concise hint from HTML error bodies (never includes secrets)."""
    s = str(text or "")
    if not s:
        return ""
    m = _HTML_ALERT_RE.search(s)
    if m:
        return m.group(1).strip()
    m = _HTML_TITLE_RE.search(s)
    if m:
        return m.group(1).strip()
    return ""


def _load_yaml(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    return yaml.safe_load(path.read_text(encoding="utf-8")) or {}


def _parse_epsg(srs: str) -> int:
    s = str(srs or "").strip().upper()
    if s.startswith("EPSG:"):
        s = s.split("EPSG:", 1)[1].strip()
    return int(s)


def _transform_bbox(
    bbox: tuple[float, float, float, float], from_epsg: int, to_epsg: int
) -> tuple[float, float, float, float]:
    if from_epsg == to_epsg:
        return bbox
    t = Transformer.from_crs(CRS.from_epsg(from_epsg), CRS.from_epsg(to_epsg), always_xy=True)
    minx, miny, maxx, maxy = bbox
    corners = [
        (minx, miny),
        (minx, maxy),
        (maxx, miny),
        (maxx, maxy),
    ]
    xs: list[float] = []
    ys: list[float] = []
    for x, y in corners:
        tx, ty = t.transform(x, y)
        xs.append(float(tx))
        ys.append(float(ty))
    return (min(xs), min(ys), max(xs), max(ys))


def _bbox_from_geojson(path: Path) -> tuple[float, float, float, float]:
    obj = json.loads(path.read_text(encoding="utf-8"))

    def _walk_coords(coords):
        if isinstance(coords, (int, float)):
            return
        if isinstance(coords, list) and coords and isinstance(coords[0], (int, float)):
            # [x, y, ...]
            yield coords
            return
        if isinstance(coords, list):
            for c in coords:
                yield from _walk_coords(c)

    coords_list = []
    if "features" in obj:
        for f in obj["features"]:
            geom = (f or {}).get("geometry") or {}
            coords_list.append(geom.get("coordinates"))
    elif "geometry" in obj:
        coords_list.append(obj["geometry"].get("coordinates"))
    else:
        coords_list.append(obj.get("coordinates"))

    minx = miny = float("inf")
    maxx = maxy = float("-inf")
    any_pt = False
    for coords in coords_list:
        for xy in _walk_coords(coords):
            if not isinstance(xy, list) or len(xy) < 2:
                continue
            x, y = float(xy[0]), float(xy[1])
            any_pt = True
            minx = min(minx, x)
            miny = min(miny, y)
            maxx = max(maxx, x)
            maxy = max(maxy, y)
    if not any_pt:
        raise ValueError(f"Invalid/empty GeoJSON coordinates: {path}")
    return (minx, miny, maxx, maxy)


def compute_bbox(
    *,
    case_dir: Path,
    boundary_file: str,
    center_lon: float | None,
    center_lat: float | None,
    input_epsg: int,
    out_srs: str,
    bbox_mode: str,
    radius_m: int,
) -> tuple[float, float, float, float]:
    # output EPSG
    out_epsg = _parse_epsg(out_srs)

    if bbox_mode.upper() in {"BOUNDARY", "AUTO"} and boundary_file.strip():
        p = Path(boundary_file)
        if not p.is_absolute():
            p = (case_dir / p).resolve()
        bbox = _bbox_from_geojson(p)
        return _transform_bbox(bbox, from_epsg=input_epsg, to_epsg=out_epsg)

    if center_lon is None or center_lat is None:
        raise ValueError("Missing center_lon/center_lat for POINT_RADIUS bbox")

    # Build bbox around point.
    if input_epsg == out_epsg == 4326:
        # Rough meter->degree conversion.
        rdeg = radius_m / 111_000.0
        return (center_lon - rdeg, center_lat - rdeg, center_lon + rdeg, center_lat + rdeg)

    in_crs = CRS.from_epsg(input_epsg)
    out_crs = CRS.from_epsg(out_epsg)
    t = Transformer.from_crs(in_crs, out_crs, always_xy=True)
    x, y = t.transform(center_lon, center_lat)

    # Assume projected CRS uses meters (true for common KR EPSGs like 3857/5186).
    if out_crs.is_projected:
        return (x - radius_m, y - radius_m, x + radius_m, y + radius_m)

    # If output is geographic (degrees), fall back to rough conversion around the (transformed) point.
    rdeg = radius_m / 111_000.0
    return (x - rdeg, y - rdeg, x + rdeg, y + rdeg)


def fetch_wms(
    *,
    layer_key: str,
    bbox: tuple[float, float, float, float],
    width: int,
    height: int,
    out_srs: str,
    wms_layers_config: Path,
    cache_config: Path,
    force_refresh: bool = False,
) -> WmsFetchResult:
    layers_cfg = _load_yaml(wms_layers_config)
    cache_cfg = _load_yaml(cache_config)

    layers = (layers_cfg.get("layers") or {})
    if layer_key not in layers:
        raise ValueError(f"Unknown WMS layer_key: {layer_key} (config: {wms_layers_config})")
    layer = layers[layer_key] or {}

    providers = (layers_cfg.get("providers") or {})
    provider_name = str(layer.get("provider") or "").strip()
    provider = providers.get(provider_name) or {}
    if not provider_name or not provider:
        raise ValueError(f"Invalid provider for layer {layer_key}: {provider_name!r}")

    base_url = str(provider.get("base_url") or "").strip()
    if not base_url:
        raise ValueError(f"Missing provider.base_url: {provider_name}")

    provider_type = str(provider.get("type") or "").strip().upper()
    omit_params = provider.get("omit_params") or []
    if isinstance(omit_params, str):
        omit_params = [omit_params]

    # Build request params
    params: dict[str, Any] = {}
    params.update(provider.get("default_params") or {})

    # Auth (query param key) if configured
    auth = provider.get("auth") or {}
    auth_key_param_candidates: list[str] = []
    auth_api_key: str = ""
    if auth:
        mode = str(auth.get("mode") or "").strip()
        if mode == "query_param":
            env_var = str(auth.get("env_var") or "").strip()
            import os

            auth_api_key = os.environ.get(env_var, "").strip() if env_var else ""
            if not auth_api_key:
                raise ValueError(f"Missing API key env var: {env_var}")
            key_param_candidates = auth.get("key_param_candidates") or []
            auth_key_param_candidates = [str(k).strip() for k in key_param_candidates if str(k).strip()]
            if not auth_key_param_candidates:
                auth_key_param_candidates = ["apiKey"]

    # Provider/layer specifics
    fmt = str(params.get("format") or layer.get("format") or "image/png")
    transparent = layer.get("transparent")
    if transparent is not None:
        params["transparent"] = "true" if bool(transparent) else "false"

    if provider_type == "OGC_WMS":
        params["layers"] = str(layer.get("layers") or "")
        params["styles"] = str(layer.get("styles") or "")
        version = str(params.get("version") or "1.1.1")
        # Use SRS for 1.1.1, CRS for 1.3.0
        if version.startswith("1.3"):
            params["crs"] = out_srs
        else:
            params["srs"] = out_srs
        params["bbox"] = ",".join(str(x) for x in bbox)
        params["width"] = int(width)
        params["height"] = int(height)
        params["format"] = fmt
    elif provider_type == "REST_IMAGE":
        # REST-style "WMS조회" endpoints: param mapping differs.
        pmap = provider.get("param_map") or {}
        params[pmap.get("layers", "layers")] = str(layer.get("layers") or "")
        params[pmap.get("srs", "srs")] = out_srs
        # Some endpoints require bbox order adjustment.
        order = str(provider.get("bbox_order") or "minx,miny,maxx,maxy")
        minx, miny, maxx, maxy = bbox
        order_map = {
            "minx": minx,
            "miny": miny,
            "maxx": maxx,
            "maxy": maxy,
        }
        bbox_vals = [order_map[k.strip()] for k in order.split(",")]
        params[pmap.get("bbox", "bbox")] = ",".join(str(x) for x in bbox_vals)
        params[pmap.get("width", "width")] = int(width)
        params[pmap.get("height", "height")] = int(height)
        params[pmap.get("format", "format")] = str(layer.get("format") or provider.get("default_format") or "png")
    else:
        raise ValueError(f"Unsupported provider.type: {provider_type} ({provider_name})")

    if omit_params:
        for k in omit_params:
            kk = str(k or "").strip()
            if not kk:
                continue
            params.pop(kk, None)

    # Cache lookup
    cache_enabled = bool(((cache_cfg.get("cache") or {}).get("enabled")) if cache_cfg else False)
    cache_root = Path(((cache_cfg.get("cache") or {}).get("root_dir")) or ".cache/maps")
    ttl_days = int(((cache_cfg.get("cache") or {}).get("ttl_days")) or 0)
    wms_tpl = ((cache_cfg.get("cache") or {}).get("wms") or {}).get("path_template") or "wms/{provider}/{layer}/{srs}/{hash}.png"

    cache_hit = False
    cache_path: Path | None = None
    if cache_enabled:
        # Hash key excludes volatile fields ordering.
        safe_params = strip_secrets_from_params(params)
        key_material = base_url + "\n" + json.dumps(sorted(safe_params.items()), ensure_ascii=False)
        h = _sha1(key_material)
        safe_srs = out_srs.replace(":", "_")
        rel = wms_tpl.format(provider=provider_name, layer=layer_key, srs=safe_srs, hash=h)
        cache_path = cache_root / rel
        if (not force_refresh) and cache_path.exists():
            if ttl_days <= 0:
                cache_hit = True
            else:
                age = _now() - datetime.fromtimestamp(cache_path.stat().st_mtime)
                if age <= timedelta(days=ttl_days):
                    cache_hit = True
        if cache_hit:
            b = cache_path.read_bytes()
            # Validate cached content: prevent HTML/XML being reused as "png".
            try:
                from PIL import Image

                with Image.open(io.BytesIO(b)) as im:
                    im.verify()
            except Exception:
                # Cache is corrupted/invalid → delete and treat as cache miss.
                try:
                    cache_path.unlink(missing_ok=True)
                except Exception:
                    pass
            else:
                return WmsFetchResult(
                    bytes_=b,
                    content_type=fmt,
                    request_url=base_url,
                    request_params=safe_params,
                    cache_hit=True,
                )

    timeout_sec = int(provider.get("timeout_sec") or 20)
    with httpx.Client(timeout=timeout_sec, follow_redirects=True) as client:
        def _fetch_once(request_params: dict[str, Any]) -> tuple[bytes, str]:
            try:
                r = client.get(base_url, params=request_params)
                r.raise_for_status()
                content_type = r.headers.get("content-type", fmt)
                b = r.content

                ct = str(content_type or "").lower()
                if "image" not in ct:
                    snippet = ""
                    try:
                        snippet = b[:800].decode("utf-8", errors="ignore")
                        snippet = " ".join(snippet.split())
                        snippet = redact_text(snippet)
                    except Exception:
                        snippet = ""
                    hint = _extract_html_error_hint(snippet)
                    detail = hint or snippet[:200]
                    raise ValueError(
                        f"WMS returned non-image content (layer={layer_key}, ct={content_type}) {detail}"
                    )

                # Also verify the bytes are decodable as an image (some servers mislabel content-type).
                try:
                    from PIL import Image

                    with Image.open(io.BytesIO(b)) as im:
                        im.verify()
                except Exception:
                    raise ValueError(f"WMS returned invalid image bytes (layer={layer_key})") from None

                return b, content_type
            except httpx.HTTPStatusError as e:
                ct = (e.response.headers.get("content-type") or "").split(";", 1)[0].strip()
                snippet = ""
                try:
                    b0 = e.response.content or b""
                    snippet = b0[:400].decode("utf-8", errors="ignore")
                    snippet = " ".join(snippet.split())[:200]
                except Exception:
                    snippet = ""
                if snippet:
                    snippet = f" {redact_text(snippet)}"
                raise ValueError(
                    f"WMS fetch failed: HTTP {e.response.status_code} (layer={layer_key}) ct={ct or 'unknown'}{snippet}"
                ) from None
            except httpx.HTTPError:
                raise ValueError(f"WMS fetch failed (layer={layer_key})") from None

        # Some providers accept only specific API key param spellings (e.g., apikey vs apiKey).
        # Try candidates in order until one returns a valid image.
        used_params = dict(params)
        best_err: Exception | None = None
        best_err_key_param: str = ""
        best_err_score: int = -1

        def _err_score(e: Exception) -> int:
            """Heuristic score to pick the most actionable error among auth-key attempts."""
            s = str(e or "")
            s_lower = s.lower()
            score = 0
            # SAFEMAP frequently returns this exact alert when the key is invalid.
            if "키 값이 맞지 않습니다" in s or "키값이 맞지 않습니다" in s:
                score = max(score, 100)
            if "http 401" in s_lower or "http 403" in s_lower or "forbidden" in s_lower or "unauthorized" in s_lower:
                score = max(score, 90)
            if "non-image content" in s_lower or "invalid image bytes" in s_lower:
                score = max(score, 30)
            if "http 404" in s_lower:
                score = max(score, 20)
            return score

        candidates = auth_key_param_candidates if (auth_api_key and auth_key_param_candidates) else []
        if candidates:
            for key_param in candidates:
                attempt = dict(params)
                attempt[str(key_param)] = auth_api_key
                try:
                    b, content_type = _fetch_once(attempt)
                    used_params = attempt
                    best_err = None
                    break
                except Exception as e:
                    score = _err_score(e)
                    if best_err is None or score > best_err_score:
                        best_err = e
                        best_err_score = score
                        best_err_key_param = str(key_param)
                    continue
            if best_err is not None:
                msg = str(best_err)
                if best_err_key_param:
                    msg = f"{msg} (auth key param={best_err_key_param})"
                raise ValueError(msg) from None
        else:
            # No auth or no candidates configured.
            b, content_type = _fetch_once(used_params)

    if cache_enabled and cache_path is not None:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_bytes(b)

    safe_params = strip_secrets_from_params(used_params)

    return WmsFetchResult(
        bytes_=b,
        content_type=content_type,
        request_url=base_url,
        request_params=safe_params,
        cache_hit=False,
    )
