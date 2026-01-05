#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import yaml
from pyproj import Transformer
from shapely.geometry import shape
from shapely.ops import transform as shp_transform


@dataclass(frozen=True)
class Bbox:
    minx: float
    miny: float
    maxx: float
    maxy: float

    def as_list(self) -> list[float]:
        return [self.minx, self.miny, self.maxx, self.maxy]


def _as_str(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _as_int(v: Any) -> int | None:
    s = _as_str(v)
    if not s:
        return None
    try:
        return int(float(s))
    except Exception:
        return None


def _as_float(v: Any) -> float | None:
    s = _as_str(v)
    if not s:
        return None
    try:
        return float(s)
    except Exception:
        return None


def _sanitize_id(s: str) -> str:
    # Keep it readable but filesystem/anchor-friendly.
    s = (s or "").strip()
    if not s:
        return ""
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^A-Za-z0-9_\-\.]+", "", s)
    return s


def _dedupe_closed_ring(points: list[list[float]]) -> list[list[float]]:
    if len(points) >= 2 and points[0] == points[-1]:
        return points[:-1]
    return points


def _read_overlay_bbox_from_csv(csv_path: Path, *, overlay_id: str | None) -> tuple[Bbox, int, str, str]:
    """
    Reads bbox/epsg/image/image_size from an AUTO_GIS WMS overlay CSV evidence row.
    Expected columns:
      - overlay_id
      - epsg
      - bbox: "minx,miny,maxx,maxy"
      - image: absolute/relative image path
      - image_size: "512x512"
    """
    with csv_path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    if not rows:
        raise SystemExit(f"overlay csv has no rows: {csv_path}")

    row = None
    if overlay_id:
        for r in rows:
            if _as_str(r.get("overlay_id")) == overlay_id:
                row = r
                break
        if row is None:
            raise SystemExit(f"overlay_id not found in csv: {overlay_id} ({csv_path})")
    else:
        row = rows[0]

    epsg = _as_int(row.get("epsg"))
    if epsg is None:
        raise SystemExit(f"missing epsg in csv row: {csv_path}")

    bbox_raw = _as_str(row.get("bbox")).strip().strip('"').strip("'")
    parts = [p.strip() for p in bbox_raw.split(",") if p.strip()]
    if len(parts) != 4:
        raise SystemExit(f"invalid bbox in csv row: {bbox_raw} ({csv_path})")
    bb = Bbox(minx=float(parts[0]), miny=float(parts[1]), maxx=float(parts[2]), maxy=float(parts[3]))

    image = _as_str(row.get("image"))
    image_size = _as_str(row.get("image_size"))
    return bb, epsg, image, image_size


def _load_features(geojson_path: Path) -> list[dict[str, Any]]:
    obj = json.loads(geojson_path.read_text(encoding="utf-8"))
    t = _as_str(obj.get("type"))
    if t == "FeatureCollection":
        feats = obj.get("features") or []
        return [f for f in feats if isinstance(f, dict)]
    if t == "Feature":
        return [obj]
    # bare geometry
    if t in {"Polygon", "MultiPolygon"}:
        return [{"type": "Feature", "geometry": obj, "properties": {}}]
    raise SystemExit(f"Unsupported GeoJSON type: {t} ({geojson_path})")


def _transform_geom(geom: Any, *, input_epsg: int, output_epsg: int) -> Any:
    if input_epsg == output_epsg:
        return geom
    transformer = Transformer.from_crs(f"EPSG:{input_epsg}", f"EPSG:{output_epsg}", always_xy=True)

    def _f(x, y, z=None):  # noqa: ANN001
        return transformer.transform(x, y)

    return shp_transform(_f, geom)


def _iter_polygons(geom: Any) -> Iterable[Any]:
    gt = geom.geom_type
    if gt == "Polygon":
        yield geom
    elif gt == "MultiPolygon":
        yield from list(geom.geoms)


def _polygon_to_layer_points(poly: Any) -> list[list[float]]:
    ring = list(poly.exterior.coords)
    pts = [[float(x), float(y)] for x, y in ring]
    return _dedupe_closed_ring(pts)


def main() -> None:
    ap = argparse.ArgumentParser(
        description=(
            "Convert a GeoJSON polygon/multipolygon into an image-annotation YAML (WORLD_LINEAR_BBOX). "
            "This is a utility to draw 'site boundary' overlays on WMS/WMTS images via scripts/annotate_image.py."
        )
    )
    ap.add_argument("--geojson", type=Path, required=True, help="Input GeoJSON (Polygon/MultiPolygon)")
    ap.add_argument("--out-annotations", type=Path, required=True, help="Output YAML path")

    # bbox sources
    ap.add_argument(
        "--bbox",
        nargs=4,
        metavar=("MINX", "MINY", "MAXX", "MAXY"),
        help="WORLD bbox (same CRS as --bbox-epsg). Required unless --bbox-from-overlay-csv is used.",
    )
    ap.add_argument("--bbox-epsg", type=int, default=None, help="EPSG of bbox/world output (e.g. 3857)")
    ap.add_argument("--input-epsg", type=int, default=4326, help="EPSG of input GeoJSON (default: 4326)")
    ap.add_argument(
        "--bbox-from-overlay-csv",
        type=Path,
        default=None,
        help="AUTO_GIS WMS overlay CSV (has epsg+bbox+image+image_size).",
    )
    ap.add_argument(
        "--overlay-id",
        type=str,
        default=None,
        help="When --bbox-from-overlay-csv is used, select row by overlay_id (default: first row).",
    )

    # output layers
    ap.add_argument("--id-prop", type=str, default=None, help="Feature property key to use as layer id")
    ap.add_argument("--id-prefix", type=str, default="GEO", help="Fallback id prefix (default: GEO)")
    ap.add_argument("--label-prop", type=str, default=None, help="Feature property key to use as label text")
    ap.add_argument(
        "--label-text",
        type=str,
        default=None,
        help="Constant label text to emit for each polygon (used when --label-prop is absent).",
    )
    ap.add_argument("--emit-labels", action="store_true", help="Emit label layers at polygon centroid")
    ap.add_argument(
        "--label-offset",
        nargs=2,
        type=int,
        default=[40, -60],
        metavar=("DX", "DY"),
        help="Label box offset from anchor (default: 40 -60)",
    )

    # style (uniform)
    ap.add_argument(
        "--fill-rgba",
        nargs=4,
        type=int,
        default=[60, 90, 140, 80],
        metavar=("R", "G", "B", "A"),
        help="Polygon fill RGBA (default: 60 90 140 80)",
    )
    ap.add_argument("--stroke", type=str, default="#2F5AA5", help="Polygon stroke color (#RRGGBB)")
    ap.add_argument("--stroke-px", type=int, default=6, help="Polygon stroke width in px (default: 6)")

    # metadata convenience
    ap.add_argument(
        "--image-path",
        type=str,
        default=None,
        help="Optional: store image_path in YAML for convenience (ignored by annotate_image).",
    )
    args = ap.parse_args()

    geojson_path = args.geojson.expanduser().resolve()
    if not geojson_path.exists():
        raise SystemExit(f"geojson not found: {geojson_path}")

    bbox = None
    bbox_epsg = args.bbox_epsg
    image_path = args.image_path
    image_size: list[int] | None = None

    if args.bbox_from_overlay_csv:
        csv_path = args.bbox_from_overlay_csv.expanduser().resolve()
        if not csv_path.exists():
            raise SystemExit(f"overlay csv not found: {csv_path}")
        bbox, epsg, img, img_size = _read_overlay_bbox_from_csv(csv_path, overlay_id=args.overlay_id)
        bbox_epsg = bbox_epsg or epsg
        image_path = image_path or img
        m = re.fullmatch(r"\s*(\d+)\s*[xX]\s*(\d+)\s*", img_size or "")
        if m:
            image_size = [int(m.group(1)), int(m.group(2))]

    if bbox is None:
        if not args.bbox:
            raise SystemExit("bbox is required: provide --bbox or --bbox-from-overlay-csv")
        b = [float(x) for x in args.bbox]
        bbox = Bbox(minx=b[0], miny=b[1], maxx=b[2], maxy=b[3])

    if bbox_epsg is None:
        raise SystemExit("bbox EPSG is required: set --bbox-epsg or use --bbox-from-overlay-csv")

    feats = _load_features(geojson_path)

    layers: list[dict[str, Any]] = []
    label_layers: list[dict[str, Any]] = []
    poly_idx = 0

    for fi, feat in enumerate(feats, start=1):
        geom_raw = feat.get("geometry")
        if not isinstance(geom_raw, dict):
            continue
        props = feat.get("properties") or {}
        if not isinstance(props, dict):
            props = {}

        try:
            geom = shape(geom_raw)
        except Exception:
            continue

        geom2 = _transform_geom(geom, input_epsg=int(args.input_epsg), output_epsg=int(bbox_epsg))

        for pi, poly in enumerate(_iter_polygons(geom2), start=1):
            poly_idx += 1

            # id
            prop_id = _as_str(props.get(args.id_prop)) if args.id_prop else ""
            base_id = _sanitize_id(prop_id) or f"{_sanitize_id(args.id_prefix) or 'GEO'}-{poly_idx:02d}"

            # polygon layer
            pts = _polygon_to_layer_points(poly)
            layers.append(
                {
                    "type": "polygon",
                    "id": base_id,
                    "points": pts,
                    "style": {"fill_rgba": list(args.fill_rgba), "stroke": args.stroke, "stroke_px": args.stroke_px},
                }
            )

            # optional label
            if args.emit_labels:
                text = ""
                if args.label_prop:
                    text = _as_str(props.get(args.label_prop))
                if not text and args.label_text:
                    text = args.label_text
                if text:
                    c = poly.representative_point()
                    label_layers.append(
                        {
                            "type": "label",
                            "id": f"LBL-{poly_idx:02d}",
                            "text": text,
                            "anchor": [float(c.x), float(c.y)],
                            "offset": [int(args.label_offset[0]), int(args.label_offset[1])],
                        }
                    )

    ann: dict[str, Any] = {
        "schema_version": "1.0",
        "coordinate_mode": "WORLD_LINEAR_BBOX",
        "bbox": bbox.as_list(),
        "world_epsg": int(bbox_epsg),
        "layers": layers + label_layers,
        "sources": {
            "geojson": str(geojson_path),
            "input_epsg": int(args.input_epsg),
            "bbox_source": str(args.bbox_from_overlay_csv.expanduser().resolve()) if args.bbox_from_overlay_csv else "manual",
        },
    }
    if image_path:
        ann["image_path"] = image_path
    if image_size:
        ann["image_size"] = image_size

    out = args.out_annotations.expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(yaml.safe_dump(ann, sort_keys=False, allow_unicode=True), encoding="utf-8")
    print(f"WROTE: {out}")


if __name__ == "__main__":
    main()
