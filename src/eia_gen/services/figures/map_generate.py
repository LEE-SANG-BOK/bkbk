from __future__ import annotations

import hashlib
import io
import json
import math
import os
import re
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import httpx
import yaml
from PIL import Image, ImageDraw, ImageFont
from pyproj import CRS, Transformer

from eia_gen.models.case import Case
from eia_gen.services.data_requests.wms import fetch_wms
from eia_gen.services.figures.derived_evidence import record_derived_evidence
from eia_gen.services.figures.geojson_utils import GeoFeature, bbox_of_features, iter_points, load_geojson_features
from eia_gen.services.figures.sample_style_v1 import (
    BOUNDARY_FILL,
    BOUNDARY_OUTLINE,
    LEGEND_BG,
    LEGEND_BORDER,
)
from eia_gen.spec.models import FigureSpec


_MAP_KIND_RE = re.compile(r"(?i)(?:^|\b)(MAP_BASE|MAP_BUFFER|MAP_OVERLAY)\b")


@dataclass(frozen=True)
class MapRecipe:
    kind: str  # MAP_BASE | MAP_BUFFER | MAP_OVERLAY
    basemap_provider: str | None
    basemap_layer: str | None
    zoom: int | None  # None=AUTO
    size: tuple[int, int]
    wms_layers: list[str]
    buffer_rings_m: list[int]
    out_srs: str  # basemap CRS (tiles) == EPSG:3857


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _sha256_file(p: Path) -> str:
    h = hashlib.sha256()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _sha1_bytes(b: bytes) -> str:
    return hashlib.sha1(b).hexdigest()


def _default_basemap_config_path() -> Path:
    # Best-effort repo-relative default.
    try:
        repo_root = Path(__file__).resolve().parents[4]
        cand = repo_root / "config" / "basemap.yaml"
        if cand.exists():
            return cand
    except Exception:
        pass
    return Path("config/basemap.yaml")


def _default_wms_layers_config_path() -> Path:
    try:
        repo_root = Path(__file__).resolve().parents[4]
        cand = repo_root / "config" / "wms_layers.yaml"
        if cand.exists():
            return cand
    except Exception:
        pass
    return Path("config/wms_layers.yaml")


def _default_cache_config_path() -> Path:
    try:
        repo_root = Path(__file__).resolve().parents[4]
        cand = repo_root / "config" / "cache.yaml"
        if cand.exists():
            return cand
    except Exception:
        pass
    return Path("config/cache.yaml")


def _load_yaml(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    obj = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    return obj if isinstance(obj, dict) else {}


def _parse_int(v: Any) -> int | None:
    try:
        return int(str(v).strip())
    except Exception:
        return None


def _parse_size(s: str) -> tuple[int, int] | None:
    m = re.match(r"^\s*(\d+)\s*[xX]\s*(\d+)\s*$", str(s or ""))
    if not m:
        return None
    try:
        w = int(m.group(1))
        h = int(m.group(2))
    except Exception:
        return None
    if w <= 0 or h <= 0:
        return None
    return (w, h)


def parse_map_recipe(gen_method: str | None) -> MapRecipe | None:
    s = str(gen_method or "").strip()
    if not s:
        return None

    m = _MAP_KIND_RE.search(s)
    if not m:
        return None
    kind = str(m.group(1)).upper()

    params: dict[str, str] = {}
    for token in re.split(r"[;\s]+", s):
        if "=" not in token:
            continue
        k, v = token.split("=", 1)
        kk = str(k or "").strip().lower()
        vv = str(v or "").strip()
        if kk:
            params[kk] = vv

    size = _parse_size(params.get("size", "")) or (1920, 1080)
    zoom = _parse_int(params.get("zoom"))

    basemap_provider = (params.get("basemap_provider") or params.get("provider") or "").strip() or None
    basemap_layer = (params.get("basemap_layer") or params.get("layer") or "").strip() or None

    wms_layers: list[str] = []
    if kind == "MAP_OVERLAY":
        raw = params.get("wms_layers") or params.get("wms_layer") or params.get("layer_key") or ""
        for part in re.split(r"[,;/]+", str(raw)):
            key = part.strip()
            if key:
                wms_layers.append(key)

    buffer_rings_m: list[int] = []
    if kind == "MAP_BUFFER":
        raw_rings = params.get("rings_m") or params.get("rings") or "300,500,1000"
        for part in re.split(r"[,;/]+", str(raw_rings)):
            n = _parse_int(part)
            if n is None:
                continue
            if n > 0:
                buffer_rings_m.append(int(n))
        # de-dup + stable order
        buffer_rings_m = sorted(set(buffer_rings_m))

    return MapRecipe(
        kind=kind,
        basemap_provider=basemap_provider,
        basemap_layer=basemap_layer,
        zoom=zoom,
        size=size,
        wms_layers=wms_layers,
        buffer_rings_m=buffer_rings_m,
        out_srs=str(params.get("out_srs") or "EPSG:3857").strip() or "EPSG:3857",
    )


def _augment_overlay_caption(caption: str, *, layer_titles: list[str]) -> str:
    c = (caption or "").strip()
    titles = [str(t or "").strip() for t in (layer_titles or []) if str(t or "").strip()]
    if not titles:
        return c
    # If the caption already contains all titles, do not duplicate.
    if c and all(t in c for t in titles):
        return c
    suffix = ", ".join(titles)
    if c:
        return f"{c} (중첩: {suffix})"
    return f"중첩도 (중첩: {suffix})"


def _pick_first_existing(base_dir: Path, candidates: list[str]) -> Path | None:
    for rel in candidates:
        p = (base_dir / rel).expanduser()
        if p.exists():
            return p
    return None


def _resolve_boundary_geojson_path(case: Case, base_dir: Path, geom_ref: str | None) -> Path | None:
    ref = (geom_ref or "").strip()
    if ref:
        p = Path(ref).expanduser()
        if not p.is_absolute():
            p = (base_dir / p).expanduser()
        if p.exists():
            return p

    boundary_file = None
    try:
        boundary_file = getattr(case.project_overview.location, "boundary_file", None)
    except Exception:
        boundary_file = None
    if isinstance(boundary_file, str) and boundary_file.strip():
        p2 = (base_dir / boundary_file.strip()).expanduser()
        if p2.exists():
            return p2

    return _pick_first_existing(
        base_dir,
        [
            "attachments/gis/site_boundary.geojson",
            "attachments/gis/boundary.geojson",
            "attachments/gis/site_boundary.json",
            "attachments/gis/boundary.json",
        ],
    )


def _clamp_lat(lat: float) -> float:
    return max(-85.05112878, min(85.05112878, float(lat)))


def _latlon_to_world_px(lon: float, lat: float, zoom: int) -> tuple[float, float]:
    lat = _clamp_lat(lat)
    lat_rad = math.radians(lat)
    n = 2.0**float(zoom)
    x = (float(lon) + 180.0) / 360.0 * 256.0 * n
    y = (1.0 - (math.log(math.tan(lat_rad) + (1.0 / math.cos(lat_rad))) / math.pi)) / 2.0 * 256.0 * n
    return (x, y)


def _world_px_to_latlon(x: float, y: float, zoom: int) -> tuple[float, float]:
    n = 2.0**float(zoom)
    lon = (float(x) / (256.0 * n)) * 360.0 - 180.0
    yy = float(y) / (256.0 * n)
    lat_rad = math.atan(math.sinh(math.pi * (1.0 - 2.0 * yy)))
    lat = math.degrees(lat_rad)
    return (float(lon), _clamp_lat(float(lat)))


def _bbox_lonlat_from_world_rect(
    world_rect: tuple[float, float, float, float],
    *,
    zoom: int,
) -> tuple[float, float, float, float]:
    x0, y0, x1, y1 = world_rect
    corners = [(x0, y0), (x0, y1), (x1, y0), (x1, y1)]
    lons: list[float] = []
    lats: list[float] = []
    for x, y in corners:
        lon, lat = _world_px_to_latlon(float(x), float(y), int(zoom))
        lons.append(float(lon))
        lats.append(float(lat))
    return (min(lons), min(lats), max(lons), max(lats))


def _meters_per_pixel(lat: float, zoom: int) -> float:
    # Web Mercator resolution
    lat_rad = math.radians(_clamp_lat(lat))
    return math.cos(lat_rad) * 2.0 * math.pi * 6378137.0 / (256.0 * (2.0**float(zoom)))


def _pick_zoom(
    bbox_lonlat: tuple[float, float, float, float],
    *,
    width_px: int,
    height_px: int,
    min_zoom: int,
    max_zoom: int,
    pad_frac: float = 0.08,
) -> int:
    min_lon, min_lat, max_lon, max_lat = bbox_lonlat
    # Expand bbox slightly for fit computation.
    for z in range(int(max_zoom), int(min_zoom) - 1, -1):
        x0, y0 = _latlon_to_world_px(min_lon, max_lat, z)
        x1, y1 = _latlon_to_world_px(max_lon, min_lat, z)
        span_x = abs(x1 - x0)
        span_y = abs(y1 - y0)
        if span_x <= width_px * (1.0 - pad_frac) and span_y <= height_px * (1.0 - pad_frac):
            return int(z)
    return int(min_zoom)


def _expand_world_rect_to_aspect(
    rect: tuple[float, float, float, float],
    *,
    target_w: int,
    target_h: int,
    margin_frac: float = 0.08,
) -> tuple[float, float, float, float]:
    x0, y0, x1, y1 = rect
    cx = (x0 + x1) / 2.0
    cy = (y0 + y1) / 2.0
    span_x = max(1.0, abs(x1 - x0))
    span_y = max(1.0, abs(y1 - y0))

    span_x *= 1.0 + float(margin_frac)
    span_y *= 1.0 + float(margin_frac)

    target_ratio = float(target_w) / float(target_h)
    cur_ratio = span_x / span_y
    if cur_ratio > target_ratio:
        span_y = span_x / target_ratio
    else:
        span_x = span_y * target_ratio

    return (cx - span_x / 2.0, cy - span_y / 2.0, cx + span_x / 2.0, cy + span_y / 2.0)


@dataclass
class _BasemapProvider:
    name: str
    type: str  # WMTS | XYZ
    url_template: str
    layer: str
    ext: str
    min_zoom: int
    max_zoom: int
    key_env: str | None = None
    key: str | None = None


def _resolve_basemap_provider(
    recipe: MapRecipe, *, basemap_cfg: dict[str, Any]
) -> tuple[_BasemapProvider | None, dict[str, Any]]:
    defaults = basemap_cfg.get("defaults") if isinstance(basemap_cfg.get("defaults"), dict) else {}
    providers = basemap_cfg.get("providers") if isinstance(basemap_cfg.get("providers"), dict) else {}

    want_provider = (recipe.basemap_provider or str(defaults.get("provider") or "")).strip() or ""
    want_layer = (recipe.basemap_layer or str(defaults.get("layer") or "")).strip() or ""

    chosen_name = want_provider or "vworld_wmts"
    chosen = providers.get(chosen_name) if isinstance(providers.get(chosen_name), dict) else None
    if not chosen:
        return (None, {"error": f"unknown basemap provider: {chosen_name}"})

    ptype = str(chosen.get("type") or "").strip().upper()
    url_template = str(chosen.get("url_template") or "").strip()
    if not url_template:
        return (None, {"error": f"basemap provider missing url_template: {chosen_name}"})

    meta: dict[str, Any] = {"provider": chosen_name, "provider_type": ptype}
    key_env = str(chosen.get("key_env") or "").strip() or None
    key = os.environ.get(key_env, "").strip() if key_env else ""
    if key_env:
        meta["key_env"] = key_env
        meta["key_present"] = bool(key)

    if ptype == "WMTS":
        if key_env and (not key):
            # If the user didn't explicitly request a WMTS provider but the default requires a key,
            # fall back to a public XYZ provider (best-effort) to avoid blank placeholders.
            explicit_provider = bool((recipe.basemap_provider or "").strip())
            if not explicit_provider:
                fb = providers.get("osm_tile") if isinstance(providers.get("osm_tile"), dict) else None
                if fb and str(fb.get("type") or "").strip().upper() == "XYZ":
                    fb_url = str(fb.get("url_template") or "").strip()
                    if fb_url:
                        meta = {
                            "provider": "osm_tile",
                            "provider_type": "XYZ",
                            "fallback_from_provider": chosen_name,
                            "fallback_reason": f"missing_key:{key_env}",
                        }
                        min_zoom = int(fb.get("min_zoom") or 0)
                        max_zoom = int(fb.get("max_zoom") or 19)
                        return (
                            _BasemapProvider(
                                name="osm_tile",
                                type="XYZ",
                                url_template=fb_url,
                                layer="",
                                ext="png",
                                min_zoom=min_zoom,
                                max_zoom=max_zoom,
                            ),
                            meta,
                        )
            return (None, {"error": f"Missing API key env var: {key_env} (provider={chosen_name})"})

        layers = chosen.get("layers") if isinstance(chosen.get("layers"), dict) else {}
        layer_name = want_layer or str(defaults.get("layer") or "Base")
        layer_cfg = layers.get(layer_name) if isinstance(layers.get(layer_name), dict) else None
        if not layer_cfg:
            return (None, {"error": f"unknown WMTS layer: {layer_name} (provider={chosen_name})"})
        ext = str(layer_cfg.get("ext") or "png").strip() or "png"
        min_zoom = int(layer_cfg.get("min_zoom") or 0)
        max_zoom = int(layer_cfg.get("max_zoom") or 19)
        return (
            _BasemapProvider(
                name=chosen_name,
                type="WMTS",
                url_template=url_template,
                layer=layer_name,
                ext=ext,
                min_zoom=min_zoom,
                max_zoom=max_zoom,
                key_env=key_env,
                key=key or None,
            ),
            meta,
        )

    if ptype == "XYZ":
        min_zoom = int(chosen.get("min_zoom") or 0)
        max_zoom = int(chosen.get("max_zoom") or 19)
        return (
            _BasemapProvider(
                name=chosen_name,
                type="XYZ",
                url_template=url_template,
                layer=want_layer or "",
                ext="png",
                min_zoom=min_zoom,
                max_zoom=max_zoom,
                key_env=key_env,
                key=key or None,
            ),
            meta,
        )

    return (None, {"error": f"unsupported basemap provider.type: {ptype} ({chosen_name})"})


def _tile_cache_path(
    *,
    cache_cfg: dict[str, Any],
    provider: _BasemapProvider,
    z: int,
    x: int,
    y: int,
) -> Path | None:
    cache = cache_cfg.get("cache") if isinstance(cache_cfg.get("cache"), dict) else {}
    if not bool(cache.get("enabled")):
        return None
    root = Path(str(cache.get("root_dir") or ".cache/maps")).expanduser()
    tpl = str(((cache.get("wmts") or {}) if isinstance(cache.get("wmts"), dict) else {}).get("path_template") or "")
    if not tpl:
        tpl = "wmts/{provider}/{layer}/{z}/{y}/{x}.{ext}"
    rel = tpl.format(provider=provider.name, layer=provider.layer or "default", z=int(z), y=int(y), x=int(x), ext=provider.ext)
    return root / rel


def _tile_cache_hit_ok(path: Path, *, ttl_days: int) -> bool:
    if not path.exists():
        return False
    if ttl_days <= 0:
        return True
    try:
        age = datetime.now() - datetime.fromtimestamp(path.stat().st_mtime)
        return age <= timedelta(days=int(ttl_days))
    except Exception:
        return False


def _fetch_tile_bytes(
    *,
    provider: _BasemapProvider,
    z: int,
    x: int,
    y: int,
    cache_cfg: dict[str, Any],
) -> tuple[bytes, bool, str]:
    cache = cache_cfg.get("cache") if isinstance(cache_cfg.get("cache"), dict) else {}
    ttl_days = int(cache.get("ttl_days") or 0)

    cache_path = _tile_cache_path(cache_cfg=cache_cfg, provider=provider, z=z, x=x, y=y)
    if cache_path is not None and _tile_cache_hit_ok(cache_path, ttl_days=ttl_days):
        return (cache_path.read_bytes(), True, str(cache_path))

    url = provider.url_template
    key = provider.key or ""
    if provider.type == "WMTS":
        if provider.key_env and not key:
            raise ValueError(f"Missing API key env var: {provider.key_env} (provider={provider.name})")
        url = url.format(key=key, layer=provider.layer, z=int(z), x=int(x), y=int(y), ext=provider.ext)
    else:
        url = url.format(z=int(z), x=int(x), y=int(y))

    with httpx.Client(timeout=20, follow_redirects=True) as client:
        r = client.get(url)
        r.raise_for_status()
        b = r.content

    if cache_path is not None:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_bytes(b)

    return (b, False, str(cache_path) if cache_path is not None else "")


def _compose_basemap(
    *,
    bbox_lonlat: tuple[float, float, float, float],
    size: tuple[int, int],
    provider: _BasemapProvider,
    zoom: int,
    cache_cfg: dict[str, Any],
) -> tuple[Image.Image, dict[str, Any]]:
    w, h = size
    min_lon, min_lat, max_lon, max_lat = bbox_lonlat
    x0, y0 = _latlon_to_world_px(min_lon, max_lat, zoom)
    x1, y1 = _latlon_to_world_px(max_lon, min_lat, zoom)

    rect = _expand_world_rect_to_aspect((x0, y0, x1, y1), target_w=w, target_h=h)
    rx0, ry0, rx1, ry1 = rect

    tile_size = 256
    tx0 = int(math.floor(rx0 / tile_size))
    ty0 = int(math.floor(ry0 / tile_size))
    tx1 = int(math.floor((rx1 - 1) / tile_size))
    ty1 = int(math.floor((ry1 - 1) / tile_size))

    tiles_w = max(1, tx1 - tx0 + 1)
    tiles_h = max(1, ty1 - ty0 + 1)

    mosaic = Image.new("RGB", (tiles_w * tile_size, tiles_h * tile_size), (235, 235, 235))

    cache_hits = 0
    tile_count = 0
    errors: list[str] = []
    for ty in range(ty0, ty1 + 1):
        for tx in range(tx0, tx1 + 1):
            tile_count += 1
            try:
                b, hit, _ = _fetch_tile_bytes(provider=provider, z=zoom, x=tx, y=ty, cache_cfg=cache_cfg)
                if hit:
                    cache_hits += 1
                with Image.open(io.BytesIO(b)) as im:
                    tile = im.convert("RGB")
            except Exception as e:
                errors.append(f"tile z={zoom} x={tx} y={ty}: {type(e).__name__}")
                tile = Image.new("RGB", (tile_size, tile_size), (220, 220, 220))
            mosaic.paste(tile, ((tx - tx0) * tile_size, (ty - ty0) * tile_size))

    crop_x0 = int(round(rx0 - tx0 * tile_size))
    crop_y0 = int(round(ry0 - ty0 * tile_size))
    crop_x1 = int(round(rx1 - tx0 * tile_size))
    crop_y1 = int(round(ry1 - ty0 * tile_size))

    cropped = mosaic.crop((crop_x0, crop_y0, crop_x1, crop_y1)).resize((w, h), resample=Image.Resampling.BILINEAR)
    meta = {
        "zoom": int(zoom),
        "world_rect_px": [float(rx0), float(ry0), float(rx1), float(ry1)],
        "tile_range": {"x0": tx0, "y0": ty0, "x1": tx1, "y1": ty1},
        "tile_count": int(tile_count),
        "cache_hits": int(cache_hits),
        "errors": errors,
    }
    return (cropped, meta)


def _draw_north_arrow(draw: ImageDraw.ImageDraw, x: int, y: int) -> None:
    draw.polygon([(x, y), (x - 14, y + 38), (x + 14, y + 38)], fill=(0, 0, 0, 220))
    draw.text((x - 6, y + 42), "N", fill=(0, 0, 0, 220))


def _draw_scale_bar(draw: ImageDraw.ImageDraw, *, lat: float, zoom: int, x: int, y: int) -> None:
    mpp = _meters_per_pixel(lat, zoom)
    candidates = [50, 100, 200, 500, 1000, 2000, 5000, 10000]
    target_px_min = 120
    target_px_max = 220
    chosen_m = 500
    chosen_px = int(round(chosen_m / max(1e-6, mpp)))
    for m in candidates:
        px = int(round(m / max(1e-6, mpp)))
        if target_px_min <= px <= target_px_max:
            chosen_m = m
            chosen_px = px
            break

    bar_h = 12
    draw.rectangle((x, y, x + chosen_px, y + bar_h), fill=(0, 0, 0, 220))
    draw.rectangle((x, y, x + chosen_px // 2, y + bar_h), fill=(255, 255, 255, 220))
    label = f"{chosen_m} m" if chosen_m < 1000 else f"{chosen_m // 1000} km"
    draw.text((x, y - 18), label, fill=(0, 0, 0, 220))


def _draw_legend(
    draw: ImageDraw.ImageDraw, w: int, h: int, entries: list[tuple[str, tuple[int, int, int, int]]]
) -> None:
    if not entries:
        return
    seen: set[str] = set()
    uniq: list[tuple[str, tuple[int, int, int, int]]] = []
    for label, color in entries:
        if label in seen:
            continue
        seen.add(label)
        uniq.append((label, color))

    box_w = 360
    box_h = 30 + 22 * len(uniq)
    x0 = w - box_w - 30
    y0 = h - box_h - 30
    draw.rounded_rectangle(
        (x0, y0, x0 + box_w, y0 + box_h), radius=10, fill=LEGEND_BG.rgba(), outline=LEGEND_BORDER.rgba()
    )
    draw.text((x0 + 12, y0 + 8), "범례", fill=(0, 0, 0, 220))
    for i, (label, color) in enumerate(uniq):
        yy = y0 + 30 + i * 22
        draw.rectangle((x0 + 12, yy + 4, x0 + 32, yy + 18), fill=color, outline=(0, 0, 0, 80))
        draw.text((x0 + 40, yy + 3), label, fill=(0, 0, 0, 220))


def _transformer(from_epsg: int, to_epsg: int) -> Transformer:
    return Transformer.from_crs(CRS.from_epsg(from_epsg), CRS.from_epsg(to_epsg), always_xy=True)


def _parse_epsg(srs: str) -> int:
    s = str(srs or "").strip().upper()
    if s.startswith("EPSG:"):
        s = s.split("EPSG:", 1)[1].strip()
    return int(s)


def _wms_bbox_from_lonlat(
    bbox_lonlat: tuple[float, float, float, float],
    *,
    out_srs: str,
) -> tuple[float, float, float, float]:
    out_epsg = _parse_epsg(out_srs)
    if out_epsg == 4326:
        return bbox_lonlat

    min_lon, min_lat, max_lon, max_lat = bbox_lonlat
    if out_epsg == 3857:
        min_lat = _clamp_lat(min_lat)
        max_lat = _clamp_lat(max_lat)

    t = _transformer(4326, out_epsg)
    corners = [
        (min_lon, min_lat),
        (min_lon, max_lat),
        (max_lon, min_lat),
        (max_lon, max_lat),
    ]
    xs: list[float] = []
    ys: list[float] = []
    for lon, lat in corners:
        x, y = t.transform(float(lon), float(lat))
        xs.append(float(x))
        ys.append(float(y))
    return (min(xs), min(ys), max(xs), max(ys))


def _bbox_lonlat_from_features(features: list[GeoFeature], *, input_epsg: int) -> tuple[float, float, float, float] | None:
    bbox = bbox_of_features(features)
    if bbox is None:
        return None
    minx, miny, maxx, maxy = bbox
    t = _transformer(input_epsg, 4326)
    corners = [(minx, miny), (minx, maxy), (maxx, miny), (maxx, maxy)]
    lons: list[float] = []
    lats: list[float] = []
    for x, y in corners:
        lon, lat = t.transform(float(x), float(y))
        lons.append(float(lon))
        lats.append(float(lat))
    return (min(lons), min(lats), max(lons), max(lats))


def _looks_like_lonlat_bbox(bbox: tuple[float, float, float, float]) -> bool:
    minx, miny, maxx, maxy = bbox
    if not (-180.0 <= minx <= 180.0 and -180.0 <= maxx <= 180.0):
        return False
    if not (-90.0 <= miny <= 90.0 and -90.0 <= maxy <= 90.0):
        return False
    return True


def _validate_lonlat_bbox(bbox: tuple[float, float, float, float]) -> tuple[float, float, float, float] | None:
    min_lon, min_lat, max_lon, max_lat = bbox
    if not (-180.0 <= min_lon <= 180.0 and -180.0 <= max_lon <= 180.0):
        return None
    if not (-90.0 <= min_lat <= 90.0 and -90.0 <= max_lat <= 90.0):
        return None
    if max_lon <= min_lon or max_lat <= min_lat:
        return None
    return (float(min_lon), float(min_lat), float(max_lon), float(max_lat))


def _center_lonlat(case: Case, *, input_epsg: int) -> tuple[float, float] | None:
    try:
        cx = case.project_overview.location.center_coord.lon.v
        cy = case.project_overview.location.center_coord.lat.v
    except Exception:
        return None
    if cx is None or cy is None:
        return None
    t = _transformer(input_epsg, 4326)
    lon, lat = t.transform(float(cx), float(cy))
    return (float(lon), float(lat))


def _draw_boundary(
    img: Image.Image,
    *,
    boundary: list[GeoFeature],
    input_epsg: int,
    zoom: int,
    world_rect: tuple[float, float, float, float],
) -> Image.Image:
    wx0, wy0, wx1, wy1 = world_rect
    out_w, out_h = img.size
    scale_x = out_w / max(1.0, (wx1 - wx0))
    scale_y = out_h / max(1.0, (wy1 - wy0))

    t = _transformer(input_epsg, 4326)

    def _pt(x: float, y: float) -> tuple[int, int]:
        lon, lat = t.transform(float(x), float(y))
        px, py = _latlon_to_world_px(float(lon), float(lat), zoom)
        ox = int(round((px - wx0) * scale_x))
        oy = int(round((py - wy0) * scale_y))
        return (ox, oy)

    # IMPORTANT: Draw semi-transparent fills on an overlay and alpha-composite.
    # Direct drawing on an RGBA image stores alpha in pixels, but later converting
    # to RGB (for stable export) drops alpha and can “white-out” the basemap.
    overlay = Image.new("RGBA", (out_w, out_h), (255, 255, 255, 0))
    od = ImageDraw.Draw(overlay, "RGBA")

    for feat in boundary:
        pts = [(_pt(x, y)) for x, y in iter_points(feat.geometry_type, feat.coordinates)]
        if len(pts) < 3:
            continue
        od.polygon(pts, fill=BOUNDARY_FILL.rgba())
        od.line(pts + [pts[0]], fill=BOUNDARY_OUTLINE.rgba(), width=6)

    return Image.alpha_composite(img, overlay)


def _buffer_circle_lonlat_points(
    *,
    center_lonlat: tuple[float, float],
    radius_m: int,
    n: int = 96,
) -> list[tuple[float, float]]:
    lon0, lat0 = center_lonlat
    # Build circle in EPSG:3857 (meters), then transform back to lon/lat.
    t_fwd = _transformer(4326, 3857)
    t_inv = _transformer(3857, 4326)
    x0, y0 = t_fwd.transform(float(lon0), float(lat0))
    pts: list[tuple[float, float]] = []
    for i in range(int(n)):
        ang = 2.0 * math.pi * (float(i) / float(n))
        x = float(x0) + float(radius_m) * math.cos(ang)
        y = float(y0) + float(radius_m) * math.sin(ang)
        lon, lat = t_inv.transform(x, y)
        pts.append((float(lon), float(lat)))
    return pts


def _draw_buffer_rings(
    draw: ImageDraw.ImageDraw,
    *,
    center_lonlat: tuple[float, float],
    rings_m: list[int],
    zoom: int,
    world_rect: tuple[float, float, float, float],
    out_size: tuple[int, int],
) -> list[tuple[str, tuple[int, int, int, int]]]:
    wx0, wy0, wx1, wy1 = world_rect
    out_w, out_h = out_size
    scale_x = out_w / max(1.0, (wx1 - wx0))
    scale_y = out_h / max(1.0, (wy1 - wy0))

    colors = [
        (255, 80, 80, 220),
        (255, 140, 80, 220),
        (255, 200, 80, 220),
        (120, 120, 255, 220),
    ]
    legend: list[tuple[str, tuple[int, int, int, int]]] = []

    for idx, r in enumerate(rings_m):
        pts_ll = _buffer_circle_lonlat_points(center_lonlat=center_lonlat, radius_m=int(r))
        pts_px: list[tuple[int, int]] = []
        for lon, lat in pts_ll:
            px, py = _latlon_to_world_px(lon, lat, zoom)
            ox = int(round((px - wx0) * scale_x))
            oy = int(round((py - wy0) * scale_y))
            pts_px.append((ox, oy))
        color = colors[idx % len(colors)]
        draw.line(pts_px + [pts_px[0]], fill=color, width=4)
        legend.append((f"영향권({int(r)}m)", color))
    return legend


def _render_map(
    *,
    case: Case,
    boundary: list[GeoFeature] | None,
    input_epsg: int,
    recipe: MapRecipe,
    out_path: Path,
    basemap_cfg: dict[str, Any],
    wms_cfg: dict[str, Any],
    wms_layers_config: Path,
    cache_cfg: dict[str, Any],
    cache_config: Path,
    title: str,
    source_origin: str | None,
) -> dict[str, Any]:
    provider, basemap_meta = _resolve_basemap_provider(recipe, basemap_cfg=basemap_cfg)
    if provider is None:
        raise ValueError(str(basemap_meta.get("error") or "invalid basemap provider"))

    bbox_lonlat = None
    if boundary:
        boundary_epsg = input_epsg
        try:
            raw_bbox = bbox_of_features(boundary)
            if raw_bbox is not None and _looks_like_lonlat_bbox(tuple(float(x) for x in raw_bbox)):  # type: ignore[arg-type]
                boundary_epsg = 4326
        except Exception:
            boundary_epsg = input_epsg

        bbox_lonlat = _bbox_lonlat_from_features(boundary, input_epsg=boundary_epsg)
        if bbox_lonlat is not None:
            bbox_lonlat = _validate_lonlat_bbox(bbox_lonlat)
        if bbox_lonlat is None:
            # Boundary exists but is unusable (EPSG mismatch / invalid coords) → ignore and fall back to center.
            boundary = None
    if bbox_lonlat is None:
        c = _center_lonlat(case, input_epsg=input_epsg)
        if c is None:
            raise ValueError("missing boundary geojson and center_coord")
        lon, lat = c
        # Default 1km bbox around center.
        t_fwd = _transformer(4326, 3857)
        t_inv = _transformer(3857, 4326)
        x0, y0 = t_fwd.transform(lon, lat)
        min_lon, min_lat = t_inv.transform(x0 - 1000.0, y0 - 1000.0)
        max_lon, max_lat = t_inv.transform(x0 + 1000.0, y0 + 1000.0)
        bbox_lonlat = _validate_lonlat_bbox((float(min_lon), float(min_lat), float(max_lon), float(max_lat)))
        if bbox_lonlat is None:
            raise ValueError("invalid lon/lat bbox computed from center_coord")

    w, h = recipe.size
    zoom = int(recipe.zoom) if recipe.zoom is not None else _pick_zoom(
        bbox_lonlat,
        width_px=w,
        height_px=h,
        min_zoom=provider.min_zoom,
        max_zoom=provider.max_zoom,
    )
    zoom = max(int(provider.min_zoom), min(int(zoom), int(provider.max_zoom)))

    # Basemap composition (tiles -> crop -> resize)
    base_img, base_meta = _compose_basemap(
        bbox_lonlat=bbox_lonlat,
        size=(w, h),
        provider=provider,
        zoom=zoom,
        cache_cfg=cache_cfg,
    )

    img = base_img.convert("RGBA")
    draw = ImageDraw.Draw(img, "RGBA")
    font = ImageFont.load_default()

    # Compute the exact world rect used by the basemap crop.
    world_rect = tuple(float(x) for x in base_meta.get("world_rect_px") or [0, 0, w, h])
    extent_lonlat = _bbox_lonlat_from_world_rect(world_rect, zoom=zoom)

    legend: list[tuple[str, tuple[int, int, int, int]]] = [("사업지 경계", BOUNDARY_FILL.rgba())]

    # Overlays (WMS) for MAP_OVERLAY
    overlays_meta: list[dict[str, Any]] = []
    if recipe.kind == "MAP_OVERLAY" and recipe.wms_layers:
        layers_cfg = wms_cfg.get("layers") if isinstance(wms_cfg.get("layers"), dict) else {}
        for layer_key in recipe.wms_layers:
            layer = layers_cfg.get(layer_key) if isinstance(layers_cfg.get(layer_key), dict) else None
            layer_title = str((layer or {}).get("title") or layer_key).strip()
            try:
                res = fetch_wms(
                    layer_key=layer_key,
                    bbox=_wms_bbox_from_lonlat(extent_lonlat, out_srs=recipe.out_srs),
                    width=w,
                    height=h,
                    out_srs=recipe.out_srs,
                    wms_layers_config=wms_layers_config,
                    cache_config=cache_config,
                )
                ov_sha1 = _sha1_bytes(res.bytes_)
                with Image.open(io.BytesIO(res.bytes_)) as ov:
                    ov_img = ov.convert("RGBA").resize((w, h), resample=Image.Resampling.BILINEAR)
                img = Image.alpha_composite(img, ov_img)
                legend.append((layer_title, (60, 60, 60, 180)))
                overlays_meta.append(
                    {
                        "layer_key": layer_key,
                        "title": layer_title,
                        "request_url": res.request_url,
                        "request_params": res.request_params,
                        "cache_hit": bool(res.cache_hit),
                        "hash_sha1": ov_sha1,
                        "default_src_ids": list((layer or {}).get("default_src_ids") or []) if isinstance(layer, dict) else [],
                    }
                )
            except Exception as e:
                overlays_meta.append(
                    {
                        "layer_key": layer_key,
                        "title": layer_title,
                        "error": f"{type(e).__name__}: {e}",
                    }
                )

    # Boundary and buffers
    if boundary:
        boundary_epsg2 = input_epsg
        try:
            raw_bbox2 = bbox_of_features(boundary)
            if raw_bbox2 is not None and _looks_like_lonlat_bbox(tuple(float(x) for x in raw_bbox2)):  # type: ignore[arg-type]
                boundary_epsg2 = 4326
        except Exception:
            boundary_epsg2 = input_epsg
        img = _draw_boundary(
            img,
            boundary=boundary,
            input_epsg=boundary_epsg2,
            zoom=zoom,
            world_rect=world_rect,
        )
        draw = ImageDraw.Draw(img, "RGBA")

    if recipe.kind == "MAP_BUFFER":
        c = _center_lonlat(case, input_epsg=input_epsg)
        if c is None and bbox_lonlat is not None:
            c = ((bbox_lonlat[0] + bbox_lonlat[2]) / 2.0, (bbox_lonlat[1] + bbox_lonlat[3]) / 2.0)
        if c is not None and recipe.buffer_rings_m:
            legend.extend(
                _draw_buffer_rings(
                    draw,
                    center_lonlat=c,
                    rings_m=recipe.buffer_rings_m,
                    zoom=zoom,
                    world_rect=world_rect,
                    out_size=(w, h),
                )
            )

    # Title box
    draw.rounded_rectangle((30, 30, 30 + 540, 30 + 48), radius=10, fill=(255, 255, 255, 210))
    draw.text((46, 44), str(title), fill=(0, 0, 0, 220), font=font)

    # North arrow + scale bar
    _draw_north_arrow(draw, w - 48, 36)
    lat0 = (extent_lonlat[1] + extent_lonlat[3]) / 2.0
    _draw_scale_bar(draw, lat=lat0, zoom=zoom, x=30, y=h - 60)
    _draw_legend(draw, w, h, legend)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    img.convert("RGB").save(out_path, format="PNG", optimize=True)
    out_sha = _sha256_file(out_path)

    note_obj: dict[str, Any] = {
        "kind": recipe.kind,
        "generated_at": _now_iso(),
        "source_origin": (source_origin or "").strip(),
        "out_sha256": out_sha,
        "size": {"width": int(w), "height": int(h)},
        "bbox_lonlat_input": [float(x) for x in bbox_lonlat],
        "bbox_lonlat": [float(x) for x in extent_lonlat],
        "out_srs": recipe.out_srs,
        "basemap": {
            **basemap_meta,
            **base_meta,
            "layer": provider.layer,
        },
    }
    if recipe.kind == "MAP_BUFFER":
        note_obj["buffer_rings_m"] = [int(x) for x in recipe.buffer_rings_m]
    if overlays_meta:
        note_obj["overlays"] = overlays_meta
        # Provide one canonical request_* for EVIDENCE_INDEX extraction (best-effort).
        for ov in overlays_meta:
            if isinstance(ov, dict) and ov.get("request_url") and ov.get("request_params"):
                note_obj["retrieved_at"] = note_obj["generated_at"]
                note_obj["request_url"] = ov.get("request_url")
                note_obj["request_params"] = ov.get("request_params")
                note_obj["hash_sha1"] = ov.get("hash_sha1") or ""
                break

    return {"note": json.dumps(note_obj, ensure_ascii=False, sort_keys=True)}


def ensure_figures_from_map_methods(
    *,
    case: Case,
    figure_specs: list[FigureSpec],
    case_xlsx: Path,
    out_dir: Path,
    basemap_config: Path | None = None,
    wms_layers_config: Path | None = None,
    cache_config: Path | None = None,
) -> dict[str, Any]:
    """Generate finished PNGs for FIGURES whose gen_method includes MAP_* recipes.

    This is best-effort and does not modify the input XLSX. Outputs are written under:
    - <case_dir>/attachments/derived/figures/maps/<FIG_ID>.png
    """
    base_dir = case_xlsx.parent.resolve()
    derived_dir = (base_dir / "attachments" / "derived").resolve()
    out_fig_dir = derived_dir / "figures" / "maps"
    out_fig_dir.mkdir(parents=True, exist_ok=True)

    basemap_config = (basemap_config or _default_basemap_config_path()).resolve()
    wms_layers_config = (wms_layers_config or _default_wms_layers_config_path()).resolve()
    cache_config = (cache_config or _default_cache_config_path()).resolve()

    basemap_cfg = _load_yaml(basemap_config)
    wms_cfg = _load_yaml(wms_layers_config)
    cache_cfg = _load_yaml(cache_config)

    want_ids = {s.id for s in figure_specs if getattr(s, "id", None)}
    caption_by_id: dict[str, str] = {s.id: (s.caption or "") for s in figure_specs if getattr(s, "id", None)}
    generated: list[str] = []
    skipped: list[str] = []
    errors: list[str] = []

    for a in getattr(case, "assets", []) or []:
        fig_id = str(getattr(a, "asset_id", "") or "").strip()
        if not fig_id or fig_id not in want_ids:
            continue

        recipe = parse_map_recipe(getattr(a, "gen_method", None))
        if recipe is None:
            continue

        fp = str(getattr(a, "file_path", "") or "").strip()
        if fp:
            p = Path(fp)
            if not p.is_absolute():
                p = (base_dir / p).expanduser()
            if p.exists():
                skipped.append(fig_id)
                continue

        # Load boundary if present.
        boundary_path = _resolve_boundary_geojson_path(case, base_dir, getattr(a, "geom_ref", None))
        boundary: list[GeoFeature] | None = None
        if boundary_path is not None:
            try:
                boundary = load_geojson_features(boundary_path)
            except Exception:
                boundary = None

        input_epsg = int(getattr(case.project_overview.location.center_coord, "epsg", 4326) or 4326)
        out_path = out_fig_dir / f"{fig_id}.png"
        stored_rel = str(out_path.relative_to(base_dir)).replace("\\", "/")

        title = (
            str(getattr(a, "title", "") or "").strip()
            or str(getattr(getattr(a, "caption", None), "t", "") or "").strip()
            or str(caption_by_id.get(fig_id) or "").strip()
            or fig_id
        )
        source_origin = str(getattr(a, "source_origin", "") or "").strip()

        if recipe.kind == "MAP_OVERLAY" and recipe.wms_layers:
            layers_cfg = wms_cfg.get("layers") if isinstance(wms_cfg.get("layers"), dict) else {}
            layer_titles: list[str] = []
            for lk in recipe.wms_layers:
                layer = layers_cfg.get(lk) if isinstance(layers_cfg.get(lk), dict) else {}
                layer_title = str(layer.get("title") or lk).strip()
                if layer_title and layer_title not in layer_titles:
                    layer_titles.append(layer_title)
            # Keep the map image title short; adjust caption text instead (traceable in source_register).
            try:
                a.caption.t = _augment_overlay_caption(a.caption.text_or_placeholder(title), layer_titles=layer_titles)
            except Exception:
                pass

        try:
            res = _render_map(
                case=case,
                boundary=boundary,
                input_epsg=input_epsg,
                recipe=recipe,
                out_path=out_path,
                basemap_cfg=basemap_cfg,
                wms_cfg=wms_cfg,
                wms_layers_config=wms_layers_config,
                cache_cfg=cache_cfg,
                cache_config=cache_config,
                title=title,
                source_origin=source_origin,
            )
        except Exception as e:
            errors.append(f"{fig_id}: {type(e).__name__}: {e}")
            continue

        basemap_provider_used = ""
        try:
            note_obj = json.loads(str(res.get("note") or "{}"))
            basemap_provider_used = str((note_obj.get("basemap") or {}).get("provider") or "").strip()
        except Exception:
            basemap_provider_used = ""

        # Update asset to point to generated file (portable path).
        a.file_path = stored_rel
        if not getattr(a, "source_ids", None) or all(str(s).strip() in {"", "S-TBD", "SRC-TBD"} for s in (a.source_ids or [])):
            # Best-effort: add common sources for basemap/WMS overlays.
            srcs: list[str] = []
            if recipe.kind in {"MAP_BASE", "MAP_BUFFER", "MAP_OVERLAY"}:
                if basemap_provider_used == "vworld_wmts":
                    srcs.append("SRC_VWORLD_WMTS")
                elif basemap_provider_used == "osm_tile":
                    srcs.append("SRC_OSM_TILE")
            if recipe.kind == "MAP_OVERLAY":
                layers_cfg = wms_cfg.get("layers") if isinstance(wms_cfg.get("layers"), dict) else {}
                for lk in recipe.wms_layers:
                    layer = layers_cfg.get(lk) if isinstance(layers_cfg.get(lk), dict) else {}
                    for sid in layer.get("default_src_ids") or []:
                        if isinstance(sid, str) and sid.strip():
                            srcs.append(sid.strip())
            a.source_ids = list(dict.fromkeys(srcs)) if srcs else ["S-TBD"]

        record_derived_evidence(
            case,
            derived_path=out_path,
            related_fig_id=fig_id,
            report_anchor=fig_id,
            src_ids=[str(s) for s in (getattr(a, "source_ids", None) or [])],
            evidence_type="derived_png",
            title=a.caption.text_or_placeholder(title),
            note=str(res.get("note") or ""),
            used_in=fig_id,
            case_dir=base_dir,
        )

        generated.append(fig_id)

    return {"generated": generated, "skipped": skipped, "errors": errors}
