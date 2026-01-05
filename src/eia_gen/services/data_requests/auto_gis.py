from __future__ import annotations

import csv
import io
import json
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from PIL import Image, ImageDraw
from pyproj import CRS, Transformer
from shapely.geometry import GeometryCollection, MultiPoint, Point, shape
from shapely.ops import nearest_points, transform, unary_union


@dataclass(frozen=True)
class AutoGisOutput:
    rows: list[dict[str, Any]]
    evidence_bytes: bytes
    evidence_filename: str
    warnings: list[str]


@dataclass(frozen=True)
class WmsOverlayInput:
    overlay_id: str
    category: str
    designation_name: str
    image_path: Path
    image_bbox: tuple[float, float, float, float]
    image_epsg: int
    src_id: str
    basis: str


def _parse_epsg(v: Any, default: int) -> int:
    if v is None:
        return default
    if isinstance(v, (int, float)):
        return int(v)
    s = str(v).strip().upper()
    if not s:
        return default
    if s.startswith("EPSG:"):
        s = s.split("EPSG:", 1)[1].strip()
    try:
        return int(s)
    except Exception:
        return default


def _load_geojson_geom(path: Path):
    obj = json.loads(path.read_text(encoding="utf-8"))
    if "features" in obj:
        geoms = []
        for f in obj.get("features") or []:
            geom = (f or {}).get("geometry")
            if geom:
                geoms.append(shape(geom))
        if not geoms:
            return GeometryCollection()
        return unary_union(geoms)
    if "geometry" in obj:
        return shape(obj["geometry"])
    return shape(obj)


def _transform_geom(geom, *, from_epsg: int, to_epsg: int):
    if from_epsg == to_epsg:
        return geom
    t = Transformer.from_crs(CRS.from_epsg(from_epsg), CRS.from_epsg(to_epsg), always_xy=True)
    return transform(t.transform, geom)


def _direction_8(from_xy: tuple[float, float], to_xy: tuple[float, float]) -> str:
    fx, fy = from_xy
    tx, ty = to_xy
    dx = tx - fx
    dy = ty - fy
    if abs(dx) < 1e-9 and abs(dy) < 1e-9:
        return "-"
    # 0deg = North
    deg = math.degrees(math.atan2(dx, dy))
    if deg < 0:
        deg += 360.0
    dirs = ["북", "북동", "동", "남동", "남", "남서", "서", "북서"]
    idx = int((deg + 22.5) // 45) % 8
    return dirs[idx]


def _world_to_pixel(
    x: float,
    y: float,
    *,
    bbox: tuple[float, float, float, float],
    width: int,
    height: int,
) -> tuple[float, float]:
    minx, miny, maxx, maxy = bbox
    dx = maxx - minx
    dy = maxy - miny
    if abs(dx) < 1e-12 or abs(dy) < 1e-12:
        return (0.0, 0.0)
    px = (x - minx) / dx * (width - 1)
    py = (maxy - y) / dy * (height - 1)
    return (float(px), float(py))


def _pixel_to_world(
    px: int,
    py: int,
    *,
    bbox: tuple[float, float, float, float],
    width: int,
    height: int,
) -> tuple[float, float]:
    minx, miny, maxx, maxy = bbox
    dx = maxx - minx
    dy = maxy - miny
    if abs(dx) < 1e-12 or abs(dy) < 1e-12:
        return (float(minx), float(miny))
    x = minx + (float(px) + 0.5) / float(width) * dx
    y = maxy - (float(py) + 0.5) / float(height) * dy
    return (float(x), float(y))


def _ensure_rgba(img: Image.Image) -> Image.Image:
    if img.mode == "RGBA":
        return img
    if img.mode == "LA":
        return img.convert("RGBA")
    if img.mode in {"RGB", "P", "L"}:
        return img.convert("RGBA")
    return img.convert("RGBA")


def _downsample(img: Image.Image, *, max_size: int) -> Image.Image:
    w, h = img.size
    if max(w, h) <= max_size:
        return img
    scale = float(max_size) / float(max(w, h))
    nw = max(1, int(round(w * scale)))
    nh = max(1, int(round(h * scale)))
    return img.resize((nw, nh), resample=Image.Resampling.NEAREST)


def _polygon_mask(
    geom,
    *,
    bbox: tuple[float, float, float, float],
    width: int,
    height: int,
) -> Image.Image:
    mask = Image.new("L", (width, height), 0)
    draw = ImageDraw.Draw(mask)

    def _draw_poly(poly):
        coords = [_world_to_pixel(x, y, bbox=bbox, width=width, height=height) for x, y in poly.exterior.coords]
        if len(coords) >= 3:
            draw.polygon(coords, fill=255)
        for ring in poly.interiors:
            hole = [_world_to_pixel(x, y, bbox=bbox, width=width, height=height) for x, y in ring.coords]
            if len(hole) >= 3:
                draw.polygon(hole, fill=0)

    if geom.geom_type == "Polygon":
        _draw_poly(geom)
    elif geom.geom_type == "MultiPolygon":
        for p in geom.geoms:
            _draw_poly(p)
    else:
        # Best-effort: use envelope.
        env = geom.envelope
        if env.geom_type == "Polygon":
            _draw_poly(env)

    return mask


def _analyze_wms_raster(
    *,
    boundary_geom,
    boundary_centroid_xy: tuple[float, float],
    image_path: Path,
    bbox: tuple[float, float, float, float],
    epsg: int,
    alpha_threshold: int,
    analysis_max_size: int,
    distance_sample_stride: int,
    distance_max_points: int,
    metric_epsg: int,
) -> tuple[dict[str, Any], list[str]]:
    warnings: list[str] = []

    if not image_path.exists():
        return ({"intersects": False, "is_applicable": "UNKNOWN", "distance_m": "", "direction": ""}, [f"missing image: {image_path}"])

    try:
        img0 = Image.open(image_path)
    except Exception as e:
        return ({"intersects": False, "is_applicable": "UNKNOWN", "distance_m": "", "direction": ""}, [f"invalid image: {e}"])

    img = _ensure_rgba(img0)
    img = _downsample(img, max_size=analysis_max_size)
    w, h = img.size

    alpha = img.getchannel("A")

    # If the alpha channel is effectively opaque everywhere, the layer likely isn't transparent.
    # In that case, overlap checks become unreliable.
    try:
        alpha_extrema = alpha.getextrema()  # (min, max)
    except Exception:
        alpha_extrema = (0, 255)
    opaque_alpha = alpha_extrema == (255, 255)
    if opaque_alpha:
        warnings.append("alpha channel is fully opaque; cannot reliably detect feature pixels")

    mask = _polygon_mask(boundary_geom, bbox=bbox, width=w, height=h)

    a_bytes = alpha.tobytes()
    m_bytes = mask.tobytes()
    if len(a_bytes) != len(m_bytes):
        return (
            {"intersects": False, "is_applicable": "UNKNOWN", "distance_m": "", "direction": ""},
            ["internal raster size mismatch"],
        )

    mask_px = 0
    active_total_px = 0
    active_inside_px = 0

    # Count pixels
    for ab, mb in zip(a_bytes, m_bytes):
        if ab > alpha_threshold:
            active_total_px += 1
            if mb:
                active_inside_px += 1
        if mb:
            mask_px += 1

    intersects = active_inside_px > 0

    if opaque_alpha:
        # If the layer isn't transparent, treating alpha>threshold as "feature pixels" becomes meaningless
        # (e.g., basemaps or generated placeholder evidences). Mark as UNKNOWN to avoid false positives.
        out = {
            "intersects": False,
            "is_applicable": "UNKNOWN",
            "distance_m": "",
            "direction": "",
            "mask_px": mask_px,
            "active_total_px": active_total_px,
            "active_inside_px": active_inside_px,
            "overlap_area_m2": "",
            "epsg": epsg,
            "bbox": bbox,
            "image_size": [w, h],
        }
        return out, warnings

    # Approx overlap area (only meaningful for projected CRS)
    overlap_area_m2 = ""
    try:
        crs = CRS.from_epsg(epsg)
        if crs.is_projected:
            minx, miny, maxx, maxy = bbox
            px_area = abs((maxx - minx) / float(w) * (maxy - miny) / float(h))
            overlap_area_m2 = round(float(active_inside_px) * px_area, 1)
    except Exception:
        overlap_area_m2 = ""

    distance_m = 0.0 if intersects else ""
    direction = "-" if intersects else ""

    if not intersects:
        if active_total_px <= 0:
            distance_m = ""
            direction = ""
        else:
            # Build a sampled MultiPoint of feature pixels.
            pts: list[tuple[float, float]] = []
            stride = max(1, int(distance_sample_stride))
            max_pts = max(10, int(distance_max_points))
            for idx, ab in enumerate(a_bytes):
                if ab <= alpha_threshold:
                    continue
                px = idx % w
                py = idx // w
                if stride > 1 and ((px % stride) != 0 or (py % stride) != 0):
                    continue
                x, y = _pixel_to_world(px, py, bbox=bbox, width=w, height=h)
                pts.append((x, y))
                if len(pts) >= max_pts:
                    break

            if not pts:
                distance_m = ""
                direction = ""
            else:
                try:
                    crs_img = CRS.from_epsg(epsg)
                    crs_metric = CRS.from_epsg(metric_epsg)

                    if crs_img.is_projected:
                        boundary_for_dist = boundary_geom
                        pts_for_dist = pts
                        centroid_for_dir = boundary_centroid_xy
                    else:
                        t = Transformer.from_crs(crs_img, crs_metric, always_xy=True)
                        boundary_for_dist = transform(t.transform, boundary_geom)
                        pts_for_dist = [t.transform(x, y) for x, y in pts]
                        centroid_for_dir = tuple(t.transform(boundary_centroid_xy[0], boundary_centroid_xy[1]))

                    mp = MultiPoint([Point(xy) for xy in pts_for_dist])
                    d = float(boundary_for_dist.distance(mp))
                    distance_m = round(d, 1)

                    _, nearest_on_overlay = nearest_points(boundary_for_dist, mp)
                    direction = _direction_8(
                        (float(centroid_for_dir[0]), float(centroid_for_dir[1])),
                        (float(nearest_on_overlay.x), float(nearest_on_overlay.y)),
                    )
                except Exception as e:
                    warnings.append(f"distance calc failed: {e}")
                    distance_m = ""
                    direction = ""

    out = {
        "intersects": intersects,
        "is_applicable": "O" if intersects else "X",
        "distance_m": distance_m,
        "direction": direction,
        "mask_px": mask_px,
        "active_total_px": active_total_px,
        "active_inside_px": active_inside_px,
        "overlap_area_m2": overlap_area_m2,
        "epsg": epsg,
        "bbox": bbox,
        "image_size": [w, h],
    }
    return out, warnings


def zoning_breakdown_from_parcels(*, parcels_rows: list[dict[str, Any]], req_id: str) -> AutoGisOutput:
    """Aggregate PARCELS.zoning + area_m2 into ZONING_BREAKDOWN rows."""
    warnings: list[str] = []
    agg: dict[str, float] = {}
    src_ids: set[str] = set()

    for r in parcels_rows:
        zoning = str(r.get("zoning") or "").strip()
        if not zoning:
            continue
        try:
            area_m2 = float(r.get("area_m2") or 0)
        except Exception:
            area_m2 = 0
        if area_m2 <= 0:
            continue
        agg[zoning] = agg.get(zoning, 0.0) + area_m2
        src = str(r.get("src_id") or "").strip()
        if src:
            src_ids.add(src)

    if not agg:
        warnings.append("PARCELS.zoning/area_m2 is empty; cannot compute zoning breakdown.")

    # Build output rows
    out_rows = [
        {"zoning": z, "area_m2": round(a, 2), "src_id": ";".join(sorted(src_ids)) if src_ids else "S-TBD"}
        for z, a in sorted(agg.items(), key=lambda x: (-x[1], x[0]))
    ]

    # Evidence as CSV (reproducible calculation output)
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["zoning", "area_m2", "src_id"])
    for r in out_rows:
        w.writerow([r.get("zoning", ""), r.get("area_m2", ""), r.get("src_id", "")])
    evidence_bytes = buf.getvalue().encode("utf-8")

    return AutoGisOutput(
        rows=out_rows,
        evidence_bytes=evidence_bytes,
        evidence_filename=f"{req_id}_zoning_breakdown.csv",
        warnings=warnings,
    )


def overlay_from_geojson(
    *,
    case_dir: Path,
    boundary_file: str,
    boundary_epsg: int,
    overlays: list[dict[str, Any]],
    metric_epsg: int,
    req_id: str,
) -> AutoGisOutput:
    """Compute O/X + min distance + direction for each overlay geometry against boundary geometry."""
    warnings: list[str] = []

    if not boundary_file.strip():
        return AutoGisOutput(
            rows=[],
            evidence_bytes=b"",
            evidence_filename=f"{req_id}_overlay.csv",
            warnings=["boundary_file is empty; cannot compute overlays."],
        )

    bpath = Path(boundary_file)
    if not bpath.is_absolute():
        bpath = (case_dir / bpath).resolve()
    if not bpath.exists():
        return AutoGisOutput(
            rows=[],
            evidence_bytes=b"",
            evidence_filename=f"{req_id}_overlay.csv",
            warnings=[f"boundary_file not found: {bpath}"],
        )

    boundary = _load_geojson_geom(bpath)
    if boundary.is_empty:
        return AutoGisOutput(
            rows=[],
            evidence_bytes=b"",
            evidence_filename=f"{req_id}_overlay.csv",
            warnings=[f"boundary geometry is empty: {bpath}"],
        )

    boundary_m = _transform_geom(boundary, from_epsg=boundary_epsg, to_epsg=metric_epsg)
    boundary_centroid = boundary_m.centroid
    bc_xy = (float(boundary_centroid.x), float(boundary_centroid.y))

    rows: list[dict[str, Any]] = []
    evidence_rows: list[dict[str, Any]] = []

    for item in overlays:
        overlay_id = str(item.get("overlay_id") or item.get("id") or "").strip()
        category = str(item.get("category") or "").strip()
        designation_name = str(item.get("designation_name") or item.get("name") or "").strip()
        geom_file = str(item.get("geometry_file") or item.get("geom_file") or "").strip()
        geom_epsg = _parse_epsg(item.get("epsg"), boundary_epsg)
        data_origin = str(item.get("data_origin") or "OFFICIAL_DB").strip()
        src_id = str(item.get("src_id") or "").strip()

        if not overlay_id:
            warnings.append("overlay item missing overlay_id")
            continue
        if not geom_file:
            warnings.append(f"[{overlay_id}] geometry_file is empty")
            rows.append(
                {
                    "overlay_id": overlay_id,
                    "category": category,
                    "designation_name": designation_name,
                    "is_applicable": "UNKNOWN",
                    "distance_m": "",
                    "direction": "",
                    "basis": "AUTO_GIS (missing geometry_file)",
                    "data_origin": data_origin,
                    "src_id": src_id or "S-TBD",
                }
            )
            continue

        gpath = Path(geom_file)
        if not gpath.is_absolute():
            gpath = (case_dir / gpath).resolve()
        if not gpath.exists():
            warnings.append(f"[{overlay_id}] geometry_file not found: {gpath}")
            rows.append(
                {
                    "overlay_id": overlay_id,
                    "category": category,
                    "designation_name": designation_name,
                    "is_applicable": "UNKNOWN",
                    "distance_m": "",
                    "direction": "",
                    "basis": f"AUTO_GIS (missing file: {gpath.name})",
                    "data_origin": data_origin,
                    "src_id": src_id or "S-TBD",
                }
            )
            continue

        try:
            geom = _load_geojson_geom(gpath)
            geom_m = _transform_geom(geom, from_epsg=geom_epsg, to_epsg=metric_epsg)
        except Exception as e:
            warnings.append(f"[{overlay_id}] failed to parse/transform geometry: {e}")
            rows.append(
                {
                    "overlay_id": overlay_id,
                    "category": category,
                    "designation_name": designation_name,
                    "is_applicable": "UNKNOWN",
                    "distance_m": "",
                    "direction": "",
                    "basis": f"AUTO_GIS (parse error: {e})",
                    "data_origin": data_origin,
                    "src_id": src_id or "S-TBD",
                }
            )
            continue

        intersects = bool(boundary_m.intersects(geom_m))
        distance_m = float(boundary_m.distance(geom_m)) if not intersects else 0.0

        dir_text = "-"
        if not intersects and not geom_m.is_empty:
            _, nearest_on_overlay = nearest_points(boundary_m, geom_m)
            dir_text = _direction_8(bc_xy, (float(nearest_on_overlay.x), float(nearest_on_overlay.y)))

        is_app = "O" if intersects else "X"
        row = {
            "overlay_id": overlay_id,
            "category": category,
            "designation_name": designation_name,
            "is_applicable": is_app,
            "distance_m": round(distance_m, 1),
            "direction": dir_text,
            "basis": f"AUTO_GIS (epsg:{geom_epsg}→{metric_epsg})",
            "data_origin": data_origin,
            "src_id": src_id or "S-TBD",
        }
        rows.append(row)

        overlap_area_m2 = 0.0
        try:
            inter = boundary_m.intersection(geom_m)
            overlap_area_m2 = float(getattr(inter, "area", 0.0) or 0.0)
        except Exception:
            overlap_area_m2 = 0.0
        evidence_rows.append(
            {
                "overlay_id": overlay_id,
                "category": category,
                "designation_name": designation_name,
                "is_applicable": is_app,
                "distance_m": round(distance_m, 3),
                "direction": dir_text,
                "overlap_area_m2": round(overlap_area_m2, 3),
                "geometry_file": str(Path(geom_file).as_posix()),
                "geometry_epsg": geom_epsg,
            }
        )

    buf = io.StringIO()
    w = csv.DictWriter(
        buf,
        fieldnames=[
            "overlay_id",
            "category",
            "designation_name",
            "is_applicable",
            "distance_m",
            "direction",
            "overlap_area_m2",
            "geometry_file",
            "geometry_epsg",
        ],
    )
    w.writeheader()
    for r in evidence_rows:
        w.writerow(r)

    return AutoGisOutput(
        rows=rows,
        evidence_bytes=buf.getvalue().encode("utf-8"),
        evidence_filename=f"{req_id}_overlay.csv",
        warnings=warnings,
    )


def overlay_from_wms_evidence(
    *,
    case_dir: Path,
    boundary_file: str,
    boundary_epsg: int,
    center_lon: float | None,
    center_lat: float | None,
    center_epsg: int,
    radius_m: int,
    items: list[WmsOverlayInput],
    req_id: str,
    metric_epsg: int = 5186,
    alpha_threshold: int = 10,
    analysis_max_size: int = 512,
    distance_sample_stride: int = 4,
    distance_max_points: int = 5000,
) -> AutoGisOutput:
    """Compute O/X + min distance + direction against boundary using WMS evidence rasters.

    This is a best-effort heuristic that treats non-transparent pixels as "feature presence".
    It is most reliable for WMS layers that render only the target polygons/lines over a
    transparent background (e.g., hazard boundaries).
    """
    warnings: list[str] = []
    rows: list[dict[str, Any]] = []
    evidence_rows: list[dict[str, Any]] = []

    boundary_geom_src = None
    bpath = None
    if boundary_file.strip():
        bpath = Path(boundary_file)
        if not bpath.is_absolute():
            bpath = (case_dir / bpath).resolve()
        if not bpath.exists():
            warnings.append(f"boundary_file not found: {bpath}")
        else:
            try:
                boundary_geom_src = _load_geojson_geom(bpath)
            except Exception as e:
                warnings.append(f"failed to parse boundary geojson: {e}")

    for it in items:
        if not it.overlay_id:
            warnings.append("WMS overlay item missing overlay_id")
            continue

        epsg = int(it.image_epsg)
        bbox = it.image_bbox
        bbox_text = ",".join(str(x) for x in bbox)

        if "__PLACEHOLDER__" in it.image_path.name:
            warnings.append(f"[{it.overlay_id}] placeholder evidence image; returning UNKNOWN")
            rows.append(
                {
                    "overlay_id": it.overlay_id,
                    "category": it.category,
                    "designation_name": it.designation_name,
                    "is_applicable": "UNKNOWN",
                    "distance_m": "",
                    "direction": "",
                    "basis": f"{it.basis} (placeholder evidence)",
                    "data_origin": "OFFICIAL_DB",
                    "src_id": it.src_id or "S-TBD",
                }
            )
            evidence_rows.append(
                {
                    "overlay_id": it.overlay_id,
                    "category": it.category,
                    "designation_name": it.designation_name,
                    "is_applicable": "UNKNOWN",
                    "distance_m": "",
                    "direction": "",
                    "mask_px": "",
                    "active_total_px": "",
                    "active_inside_px": "",
                    "overlap_area_m2": "",
                    "image": str(it.image_path.as_posix()),
                    "epsg": epsg,
                    "bbox": bbox_text,
                    "image_size": "",
                }
            )
            continue

        # Resolve boundary geometry in raster CRS.
        boundary_geom = None
        boundary_centroid_xy = (0.0, 0.0)
        if boundary_geom_src is not None and not boundary_geom_src.is_empty:
            try:
                boundary_geom = _transform_geom(boundary_geom_src, from_epsg=boundary_epsg, to_epsg=epsg)
                c = boundary_geom.centroid
                boundary_centroid_xy = (float(c.x), float(c.y))
            except Exception as e:
                warnings.append(f"[{it.overlay_id}] boundary transform failed: {e}")

        if boundary_geom is None:
            # Fallback: circle buffer from center point.
            if center_lon is None or center_lat is None:
                rows.append(
                    {
                        "overlay_id": it.overlay_id,
                        "category": it.category,
                        "designation_name": it.designation_name,
                        "is_applicable": "UNKNOWN",
                        "distance_m": "",
                        "direction": "",
                        "basis": f"{it.basis} (missing boundary and center coords)",
                        "data_origin": "OFFICIAL_DB",
                        "src_id": it.src_id or "S-TBD",
                    }
                )
                evidence_rows.append(
                    {
                        "overlay_id": it.overlay_id,
                        "image": str(it.image_path.as_posix()),
                        "epsg": epsg,
                        "bbox": ",".join(str(x) for x in bbox),
                        "result": "UNKNOWN",
                        "reason": "missing boundary and center coords",
                    }
                )
                continue
            try:
                t = Transformer.from_crs(CRS.from_epsg(center_epsg), CRS.from_epsg(epsg), always_xy=True)
                x, y = t.transform(float(center_lon), float(center_lat))
                boundary_geom = Point(float(x), float(y)).buffer(float(radius_m))
                boundary_centroid_xy = (float(x), float(y))
            except Exception as e:
                warnings.append(f"[{it.overlay_id}] failed to build point buffer boundary: {e}")
                rows.append(
                    {
                        "overlay_id": it.overlay_id,
                        "category": it.category,
                        "designation_name": it.designation_name,
                        "is_applicable": "UNKNOWN",
                        "distance_m": "",
                        "direction": "",
                        "basis": f"{it.basis} (boundary fallback error)",
                        "data_origin": "OFFICIAL_DB",
                        "src_id": it.src_id or "S-TBD",
                    }
                )
                continue

        analysis, w = _analyze_wms_raster(
            boundary_geom=boundary_geom,
            boundary_centroid_xy=boundary_centroid_xy,
            image_path=it.image_path,
            bbox=bbox,
            epsg=epsg,
            alpha_threshold=alpha_threshold,
            analysis_max_size=analysis_max_size,
            distance_sample_stride=distance_sample_stride,
            distance_max_points=distance_max_points,
            metric_epsg=metric_epsg,
        )
        for ww in w:
            warnings.append(f"[{it.overlay_id}] {ww}")

        is_app = str(analysis.get("is_applicable") or "UNKNOWN")
        distance_m = analysis.get("distance_m")
        direction = str(analysis.get("direction") or "")

        rows.append(
            {
                "overlay_id": it.overlay_id,
                "category": it.category,
                "designation_name": it.designation_name,
                "is_applicable": is_app,
                "distance_m": distance_m,
                "direction": direction,
                "basis": it.basis,
                "data_origin": "OFFICIAL_DB",
                "src_id": it.src_id or "S-TBD",
            }
        )
        evidence_rows.append(
            {
                "overlay_id": it.overlay_id,
                "category": it.category,
                "designation_name": it.designation_name,
                "is_applicable": is_app,
                "distance_m": distance_m,
                "direction": direction,
                "mask_px": analysis.get("mask_px"),
                "active_total_px": analysis.get("active_total_px"),
                "active_inside_px": analysis.get("active_inside_px"),
                "overlap_area_m2": analysis.get("overlap_area_m2"),
                "image": str(it.image_path.as_posix()),
                "epsg": analysis.get("epsg"),
                "bbox": bbox_text,
                "image_size": "x".join(str(x) for x in (analysis.get("image_size") or [])),
            }
        )

    buf = io.StringIO()
    fieldnames = [
        "overlay_id",
        "category",
        "designation_name",
        "is_applicable",
        "distance_m",
        "direction",
        "mask_px",
        "active_total_px",
        "active_inside_px",
        "overlap_area_m2",
        "image",
        "epsg",
        "bbox",
        "image_size",
    ]
    w = csv.DictWriter(buf, fieldnames=fieldnames)
    w.writeheader()
    for r in evidence_rows:
        w.writerow({k: r.get(k, "") for k in fieldnames})

    return AutoGisOutput(
        rows=rows,
        evidence_bytes=buf.getvalue().encode("utf-8"),
        evidence_filename=f"{req_id}_wms_overlay.csv",
        warnings=warnings,
    )
