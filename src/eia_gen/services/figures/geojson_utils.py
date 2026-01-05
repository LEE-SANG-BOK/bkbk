from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable


@dataclass(frozen=True)
class GeoFeature:
    geometry_type: str
    coordinates: Any
    properties: dict[str, Any]


def load_geojson_features(path: str | Path) -> list[GeoFeature]:
    p = Path(path)
    data = json.loads(p.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("geojson root must be object")
    if data.get("type") != "FeatureCollection":
        raise ValueError("geojson must be a FeatureCollection")
    raw_features = data.get("features") or []
    out: list[GeoFeature] = []
    for f in raw_features:
        if not isinstance(f, dict):
            continue
        geom = f.get("geometry")
        props = f.get("properties") or {}
        if not isinstance(props, dict):
            props = {}
        if not isinstance(geom, dict):
            continue
        gtype = str(geom.get("type") or "").strip()
        coords = geom.get("coordinates")
        if not gtype:
            continue
        out.append(GeoFeature(geometry_type=gtype, coordinates=coords, properties=props))
    return out


def iter_points(geometry_type: str, coordinates: Any) -> Iterable[tuple[float, float]]:
    """Yield (x,y) points from GeoJSON geometry coordinates."""
    gt = geometry_type

    def _iter(seq: Any) -> Iterable[tuple[float, float]]:
        if not isinstance(seq, list):
            return []
        # coordinate pair
        if len(seq) >= 2 and isinstance(seq[0], (int, float)) and isinstance(seq[1], (int, float)):
            return [(float(seq[0]), float(seq[1]))]
        # nested lists
        pts: list[tuple[float, float]] = []
        for item in seq:
            pts.extend(list(_iter(item)))
        return pts

    if gt in {"Point", "MultiPoint", "LineString", "MultiLineString", "Polygon", "MultiPolygon"}:
        return _iter(coordinates)
    return []


def bbox_of_features(features: Iterable[GeoFeature]) -> tuple[float, float, float, float] | None:
    minx = miny = maxx = maxy = None
    for f in features:
        for x, y in iter_points(f.geometry_type, f.coordinates):
            if minx is None:
                minx = maxx = x
                miny = maxy = y
                continue
            minx = min(minx, x)
            miny = min(miny, y)
            maxx = max(maxx, x)
            maxy = max(maxy, y)
    if minx is None or miny is None or maxx is None or maxy is None:
        return None
    return float(minx), float(miny), float(maxx), float(maxy)

