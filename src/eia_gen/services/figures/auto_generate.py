from __future__ import annotations

import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from PIL import Image, ImageDraw, ImageFont

from eia_gen.models.case import Case
from eia_gen.services.figures.geojson_utils import GeoFeature, bbox_of_features, load_geojson_features
from eia_gen.services.figures.sample_style_v1 import (
    BOUNDARY_FILL,
    BOUNDARY_OUTLINE,
    DRAINAGE_LINE,
    FACILITY_BUILDING,
    FACILITY_OTHER,
    FACILITY_POND,
    FACILITY_ROAD,
    FACILITY_WALKWAY,
    LABEL_BG,
    LABEL_FG,
    LEGEND_BG,
    LEGEND_BORDER,
)
from eia_gen.spec.models import FigureSpec


@dataclass(frozen=True)
class _GeoInputs:
    boundary: list[GeoFeature]
    facilities: list[GeoFeature]
    drainage: list[GeoFeature]


def _get_case_base_dir(case_xlsx: Path) -> Path:
    return case_xlsx.parent


def _pick_first_existing(base_dir: Path, candidates: list[str]) -> Path | None:
    for rel in candidates:
        p = (base_dir / rel).expanduser()
        if p.exists():
            return p
    return None


def _load_geo_inputs(base_dir: Path, boundary_file: str | None = None) -> _GeoInputs | None:
    boundary_path = None
    if boundary_file:
        p = (base_dir / boundary_file).expanduser()
        if p.exists():
            boundary_path = p
    if boundary_path is None:
        boundary_path = _pick_first_existing(
            base_dir,
            [
                "attachments/gis/site_boundary.geojson",
                "attachments/gis/boundary.geojson",
                "attachments/gis/site_boundary.json",
                "attachments/gis/boundary.json",
            ],
        )
    facilities_path = _pick_first_existing(
        base_dir,
        [
            "attachments/gis/facility_layout.geojson",
            "attachments/gis/facilities.geojson",
            "attachments/gis/facility_layout.json",
            "attachments/gis/facilities.json",
        ],
    )
    drainage_path = _pick_first_existing(
        base_dir,
        [
            "attachments/gis/drainage_lines.geojson",
            "attachments/gis/drainage.geojson",
            "attachments/gis/drainage_lines.json",
            "attachments/gis/drainage.json",
        ],
    )

    if boundary_path is None:
        return None

    boundary = load_geojson_features(boundary_path)
    facilities = load_geojson_features(facilities_path) if facilities_path else []
    drainage = load_geojson_features(drainage_path) if drainage_path else []
    return _GeoInputs(boundary=boundary, facilities=facilities, drainage=drainage)


def _centroid_of_ring(points: list[tuple[float, float]]) -> tuple[float, float] | None:
    if len(points) < 3:
        return None
    # Polygon centroid (approx; ring may not be closed)
    area = 0.0
    cx = 0.0
    cy = 0.0
    pts = points[:]
    if pts[0] != pts[-1]:
        pts.append(pts[0])
    for i in range(len(pts) - 1):
        x0, y0 = pts[i]
        x1, y1 = pts[i + 1]
        cross = x0 * y1 - x1 * y0
        area += cross
        cx += (x0 + x1) * cross
        cy += (y0 + y1) * cross
    if abs(area) < 1e-12:
        return None
    area *= 0.5
    cx /= 6.0 * area
    cy /= 6.0 * area
    return cx, cy


def _scale_meters_per_degree(lat_deg: float) -> tuple[float, float]:
    # very rough
    lat_m = 110_540.0
    lon_m = 111_320.0 * math.cos(math.radians(lat_deg))
    return lon_m, lat_m


@dataclass(frozen=True)
class _Viewport:
    minx: float
    miny: float
    maxx: float
    maxy: float
    epsg: int

    @property
    def width(self) -> float:
        return self.maxx - self.minx

    @property
    def height(self) -> float:
        return self.maxy - self.miny

    def expand(self, frac: float = 0.05) -> "_Viewport":
        dx = self.width * frac
        dy = self.height * frac
        if self.width == 0:
            dx = 1.0
        if self.height == 0:
            dy = 1.0
        return _Viewport(self.minx - dx, self.miny - dy, self.maxx + dx, self.maxy + dy, self.epsg)


def _project_radius_to_coord_units(epsg: int, lat0: float, radius_m: float) -> tuple[float, float]:
    # returns (dx, dy) in coordinate units for a circle buffer
    if epsg in {5179, 5186}:
        return radius_m, radius_m
    # assume WGS84 degrees
    lon_m, lat_m = _scale_meters_per_degree(lat0)
    dx = radius_m / max(lon_m, 1.0)
    dy = radius_m / max(lat_m, 1.0)
    return dx, dy


class _CoordToPixel:
    def __init__(self, vp: _Viewport, width_px: int, height_px: int, pad_px: int = 60) -> None:
        self.vp = vp
        self.w = width_px
        self.h = height_px
        self.pad = pad_px
        self.inner_w = max(1, self.w - 2 * self.pad)
        self.inner_h = max(1, self.h - 2 * self.pad)

    def pt(self, x: float, y: float) -> tuple[int, int]:
        if self.vp.width == 0:
            nx = 0.5
        else:
            nx = (x - self.vp.minx) / self.vp.width
        if self.vp.height == 0:
            ny = 0.5
        else:
            ny = (y - self.vp.miny) / self.vp.height
        px = int(self.pad + nx * self.inner_w)
        py = int(self.h - self.pad - ny * self.inner_h)
        return px, py


def _extract_outer_ring(feature: GeoFeature) -> list[tuple[float, float]] | None:
    gt = feature.geometry_type
    coords = feature.coordinates
    if gt == "Polygon":
        if isinstance(coords, list) and coords and isinstance(coords[0], list):
            ring = coords[0]
            if isinstance(ring, list):
                pts: list[tuple[float, float]] = []
                for p in ring:
                    if isinstance(p, list) and len(p) >= 2:
                        pts.append((float(p[0]), float(p[1])))
                return pts
    if gt == "MultiPolygon":
        if isinstance(coords, list) and coords:
            first = coords[0]
            if isinstance(first, list) and first and isinstance(first[0], list):
                ring = first[0]
                if isinstance(ring, list):
                    pts2: list[tuple[float, float]] = []
                    for p in ring:
                        if isinstance(p, list) and len(p) >= 2:
                            pts2.append((float(p[0]), float(p[1])))
                    return pts2
    return None


def _iter_lines(feature: GeoFeature) -> list[list[tuple[float, float]]]:
    gt = feature.geometry_type
    coords = feature.coordinates
    out: list[list[tuple[float, float]]] = []
    if gt == "LineString":
        if isinstance(coords, list):
            pts: list[tuple[float, float]] = []
            for p in coords:
                if isinstance(p, list) and len(p) >= 2:
                    pts.append((float(p[0]), float(p[1])))
            if pts:
                out.append(pts)
    elif gt == "MultiLineString":
        if isinstance(coords, list):
            for line in coords:
                if not isinstance(line, list):
                    continue
                pts = []
                for p in line:
                    if isinstance(p, list) and len(p) >= 2:
                        pts.append((float(p[0]), float(p[1])))
                if pts:
                    out.append(pts)
    return out


def _classify_facility_color(props: dict[str, Any]) -> tuple[tuple[int, int, int, int], str]:
    raw = " ".join(
        [
            str(props.get(k) or "")
            for k in ("category", "fac_type", "facility_type", "type", "name", "facility_name")
        ]
    ).lower()
    if any(k in raw for k in ["pond", "water", "연못", "수면", "저류", "침사지"]):
        return FACILITY_POND.rgba(), "수면/저류"
    if any(k in raw for k in ["walk", "trail", "보행", "산책", "데크"]):
        return FACILITY_WALKWAY.rgba(), "보행로"
    if any(k in raw for k in ["road", "parking", "주차", "도로", "진입"]):
        return FACILITY_ROAD.rgba(), "도로/주차"
    if any(k in raw for k in ["building", "숙박", "관리", "화장실", "매점", "동"]):
        return FACILITY_BUILDING.rgba(), "건축물"
    return FACILITY_OTHER.rgba(), "기타시설"


def _draw_label(draw: ImageDraw.ImageDraw, xy: tuple[int, int], text: str, font: ImageFont.ImageFont) -> None:
    if not text:
        return
    w, h = draw.textbbox((0, 0), text, font=font)[2:]
    pad = 6
    x, y = xy
    box = (x - w // 2 - pad, y - h // 2 - pad, x + w // 2 + pad, y + h // 2 + pad)
    draw.rounded_rectangle(box, radius=6, fill=LABEL_BG.rgba())
    draw.text((x - w // 2, y - h // 2), text, font=font, fill=LABEL_FG)


def _draw_north_arrow(draw: ImageDraw.ImageDraw, x: int, y: int) -> None:
    # simple arrow
    draw.polygon([(x, y), (x - 10, y + 30), (x + 10, y + 30)], fill=(0, 0, 0, 180))
    draw.text((x - 6, y + 32), "N", fill=(0, 0, 0, 200))


def _draw_scale_bar(
    draw: ImageDraw.ImageDraw,
    vp: _Viewport,
    transformer: _CoordToPixel,
    lat0: float,
    x: int,
    y: int,
    max_width_px: int = 260,
) -> None:
    # approximate bbox width in meters
    if vp.epsg in {5179, 5186}:
        width_m = vp.width
    else:
        lon_m, _lat_m = _scale_meters_per_degree(lat0)
        width_m = vp.width * lon_m
    if not width_m or width_m <= 0:
        return
    target = width_m / 4.0
    nice = [50, 100, 200, 500, 1_000, 2_000, 5_000, 10_000]
    length_m = None
    for v in nice:
        if v <= target:
            length_m = v
    if length_m is None:
        length_m = nice[0]

    # convert length_m to coord units
    if vp.epsg in {5179, 5186}:
        dx = length_m
    else:
        lon_m, _lat_m = _scale_meters_per_degree(lat0)
        dx = length_m / max(lon_m, 1.0)

    x0_coord = vp.minx + vp.width * 0.05
    y0_coord = vp.miny + vp.height * 0.05
    p0 = transformer.pt(x0_coord, y0_coord)
    p1 = transformer.pt(x0_coord + dx, y0_coord)
    bar_len_px = max(10, p1[0] - p0[0])
    if bar_len_px > max_width_px:
        # too long; clip visually
        bar_len_px = max_width_px
    x0, y0 = x, y
    x1 = x0 + bar_len_px
    draw.rectangle((x0, y0, x1, y0 + 10), fill=(0, 0, 0, 160))
    draw.text((x0, y0 + 12), f"{int(length_m)} m", fill=(0, 0, 0, 200))


def _draw_legend(draw: ImageDraw.ImageDraw, w: int, h: int, entries: list[tuple[str, tuple[int, int, int, int]]]) -> None:
    if not entries:
        return
    # de-dup by label
    seen: set[str] = set()
    uniq: list[tuple[str, tuple[int, int, int, int]]] = []
    for label, color in entries:
        if label in seen:
            continue
        seen.add(label)
        uniq.append((label, color))

    box_w = 320
    box_h = 28 + 22 * len(uniq)
    x0 = w - box_w - 30
    y0 = h - box_h - 30
    draw.rounded_rectangle((x0, y0, x0 + box_w, y0 + box_h), radius=8, fill=LEGEND_BG.rgba(), outline=LEGEND_BORDER.rgba())
    draw.text((x0 + 12, y0 + 8), "범례", fill=(0, 0, 0, 220))
    for i, (label, color) in enumerate(uniq):
        yy = y0 + 28 + i * 22
        draw.rectangle((x0 + 12, yy + 4, x0 + 32, yy + 18), fill=color, outline=(0, 0, 0, 80))
        draw.text((x0 + 40, yy + 3), label, fill=(0, 0, 0, 220))


def _render_map_base(
    geo: _GeoInputs,
    epsg: int,
    title: str,
    out_path: Path,
    *,
    include_facilities: bool = True,
    include_drainage: bool = True,
    include_buffer_m: float | None = None,
) -> None:
    all_feats = geo.boundary + (geo.facilities if include_facilities else []) + (geo.drainage if include_drainage else [])
    bbox = bbox_of_features(all_feats)
    if bbox is None:
        raise ValueError("failed to compute bbox from geojson")
    minx, miny, maxx, maxy = bbox
    vp = _Viewport(minx, miny, maxx, maxy, epsg).expand(0.06)

    W, H = 1920, 1080
    img = Image.new("RGBA", (W, H), (245, 245, 245, 255))
    draw = ImageDraw.Draw(img, "RGBA")
    font = ImageFont.load_default()

    transformer = _CoordToPixel(vp, W, H)

    # boundary
    legend: list[tuple[str, tuple[int, int, int, int]]] = [("사업지 경계", BOUNDARY_FILL.rgba())]
    for feat in geo.boundary:
        ring = _extract_outer_ring(feat)
        if not ring:
            continue
        pts = [transformer.pt(x, y) for x, y in ring]
        draw.polygon(pts, fill=BOUNDARY_FILL.rgba())
        draw.line(pts + [pts[0]], fill=BOUNDARY_OUTLINE.rgba(), width=6)

    # facilities
    if include_facilities:
        for feat in geo.facilities:
            ring = _extract_outer_ring(feat)
            if not ring:
                continue
            color, label = _classify_facility_color(feat.properties)
            legend.append((label, color))
            pts = [transformer.pt(x, y) for x, y in ring]
            draw.polygon(pts, fill=color)
            draw.line(pts + [pts[0]], fill=(255, 255, 255, 180), width=2)
            name = str(feat.properties.get("name") or feat.properties.get("facility_name") or "").strip()
            if name:
                c = _centroid_of_ring(ring)
                if c:
                    _draw_label(draw, transformer.pt(c[0], c[1]), name, font)

    # drainage lines
    if include_drainage:
        legend.append(("배수/수로", DRAINAGE_LINE.rgba()))
        for feat in geo.drainage:
            for line in _iter_lines(feat):
                pts = [transformer.pt(x, y) for x, y in line]
                if len(pts) >= 2:
                    draw.line(pts, fill=DRAINAGE_LINE.rgba(), width=3)

    # buffer ring (influence area)
    if include_buffer_m is not None:
        # center at boundary bbox center
        cx = (vp.minx + vp.maxx) / 2.0
        cy = (vp.miny + vp.maxy) / 2.0
        lat0 = cy if epsg != 4326 else cy
        dx, dy = _project_radius_to_coord_units(epsg, lat0=lat0, radius_m=include_buffer_m)
        p0 = transformer.pt(cx - dx, cy - dy)
        p1 = transformer.pt(cx + dx, cy + dy)
        # Pillow expects bbox as (left, top, right, bottom) with right>=left and bottom>=top.
        left = min(p0[0], p1[0])
        top = min(p0[1], p1[1])
        right = max(p0[0], p1[0])
        bottom = max(p0[1], p1[1])
        draw.ellipse((left, top, right, bottom), outline=(255, 80, 80, 220), width=4)
        legend.append((f"영향권({int(include_buffer_m)}m)", (255, 80, 80, 120)))

    # title
    draw.rounded_rectangle((30, 30, 30 + 520, 30 + 44), radius=10, fill=(255, 255, 255, 200))
    draw.text((46, 42), title, fill=(0, 0, 0, 220))

    # north arrow + scale bar
    _draw_north_arrow(draw, W - 40, 40)
    lat0 = (vp.miny + vp.maxy) / 2.0
    _draw_scale_bar(draw, vp, transformer, lat0=lat0, x=30, y=H - 60)

    _draw_legend(draw, W, H, legend)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    img.convert("RGB").save(out_path, format="PNG")


def _existing_asset_paths(case: Case, asset_type: str, base_dir: Path) -> list[Path]:
    paths: list[Path] = []
    for a in case.assets:
        if a.type != asset_type:
            continue
        fp = (a.file_path or "").strip()
        if not fp:
            continue
        p = (base_dir / fp).expanduser() if not Path(fp).is_absolute() else Path(fp)
        if p.exists():
            paths.append(p)
    return paths


def _asset_file_exists(file_path: str, base_dir: Path) -> bool:
    fp = (file_path or "").strip()
    if not fp:
        return False
    p = Path(fp)
    if p.is_absolute():
        return p.exists()
    # try both cwd-relative and case-base-relative
    return p.exists() or (base_dir / fp).exists()


def ensure_figures_from_geojson(
    *,
    case: Case,
    figure_specs: list[FigureSpec],
    case_xlsx: Path,
    out_dir: Path,
    derived_dir: Path | None = None,
) -> None:
    """Best-effort auto-generation for map-like figures.

    This does NOT attempt to synthesize official design drawings.
    """
    base_dir = _get_case_base_dir(case_xlsx).resolve()
    # v2: LOCATION.boundary_file stored as extra if present (optional)
    boundary_file = None
    try:
        boundary_file = getattr(case.project_overview.location, "boundary_file", None)
    except Exception:
        boundary_file = None

    geo = _load_geo_inputs(base_dir, boundary_file=str(boundary_file) if boundary_file else None)
    if geo is None:
        return

    epsg = int(getattr(case.project_overview.location.center_coord, "epsg", 4326) or 4326)
    out_fig_dir: Path | None = None
    try:
        derived_dir2 = (derived_dir or (base_dir / "attachments" / "derived")).resolve()
        out_fig_dir = derived_dir2 / "figures" / "auto"
        out_fig_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        out_fig_dir = out_dir.resolve() / "figures"
        out_fig_dir.mkdir(parents=True, exist_ok=True)

    for spec in figure_specs:
        if spec.asset_type not in {
            "location_map",
            "landuse_plan",
            "layout_plan",
            "drainage_map",
            "influence_area_map",
            "dia_target_area_map",
            "stormwater_plan_map",
        }:
            continue
        # If user already provided an existing figure for this type, skip.
        if _existing_asset_paths(case, spec.asset_type, base_dir):
            continue

        out_path = out_fig_dir / f"{spec.id}.png"
        # Store file_path relative to the case folder when possible (portable),
        # otherwise fall back to an absolute path.
        try:
            stored_path = str(out_path.relative_to(base_dir))
        except Exception:
            stored_path = str(out_path)
        title = spec.caption

        try:
            if spec.asset_type == "location_map":
                _render_map_base(geo, epsg, title=title, out_path=out_path, include_facilities=False, include_drainage=False)
            elif spec.asset_type == "landuse_plan":
                _render_map_base(geo, epsg, title=title, out_path=out_path, include_facilities=False, include_drainage=False)
            elif spec.asset_type == "layout_plan":
                _render_map_base(geo, epsg, title=title, out_path=out_path, include_facilities=True, include_drainage=False)
            elif spec.asset_type == "drainage_map":
                _render_map_base(geo, epsg, title=title, out_path=out_path, include_facilities=True, include_drainage=True)
            elif spec.asset_type == "stormwater_plan_map":
                _render_map_base(geo, epsg, title=title, out_path=out_path, include_facilities=True, include_drainage=True)
            elif spec.asset_type == "dia_target_area_map":
                _render_map_base(geo, epsg, title=title, out_path=out_path, include_facilities=False, include_drainage=True)
            elif spec.asset_type == "influence_area_map":
                radius = None
                if case.survey_plan and case.survey_plan.influence_area and case.survey_plan.influence_area.radius_m:
                    radius = case.survey_plan.influence_area.radius_m.v
                _render_map_base(
                    geo,
                    epsg,
                    title=title,
                    out_path=out_path,
                    include_facilities=False,
                    include_drainage=False,
                    include_buffer_m=float(radius) if radius else 500.0,
                )
        except Exception:
            continue

        # Attach generated file as an asset (so resolve_figure picks it up)
        # Prefer updating existing blank asset entry if present.
        updated = False
        for a in case.assets:
            if a.type != spec.asset_type:
                continue
            # If the current asset points to a missing/non-existent file, overwrite it.
            if _asset_file_exists(a.file_path, base_dir):
                continue
            a.file_path = stored_path
            a.caption.t = a.caption.t or spec.caption
            updated = True
            break
        if not updated:
            from eia_gen.models.case import Asset
            from eia_gen.models.fields import TextField

            case.assets.append(
                Asset(
                    asset_id=spec.id,
                    type=spec.asset_type,
                    file_path=stored_path,
                    caption=TextField(t=spec.caption),
                    source_ids=["S-TBD"],
                )
            )
