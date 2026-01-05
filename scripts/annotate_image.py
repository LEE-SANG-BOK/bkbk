#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import re
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

import yaml
from PIL import Image, ImageDraw, ImageFont


CoordinateMode = Literal["PIXEL", "WORLD_LINEAR_BBOX"]

STYLE = dict[str, Any]


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


def _as_rgba(v: Any, *, default: tuple[int, int, int, int]) -> tuple[int, int, int, int]:
    if v is None:
        return default
    if isinstance(v, (list, tuple)) and len(v) == 4:
        return (int(v[0]), int(v[1]), int(v[2]), int(v[3]))
    return default


def _as_rgb(v: Any, *, default: tuple[int, int, int]) -> tuple[int, int, int]:
    if v is None:
        return default
    if isinstance(v, (list, tuple)) and len(v) == 3:
        return (int(v[0]), int(v[1]), int(v[2]))
    return default


def _hex_to_rgb(hex_color: str, *, default: tuple[int, int, int]) -> tuple[int, int, int]:
    s = (hex_color or "").strip()
    if not s:
        return default
    if s.startswith("#"):
        s = s[1:]
    if len(s) != 6:
        return default
    try:
        return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
    except Exception:
        return default


def _style_get(style: STYLE, path: list[str], default: Any) -> Any:
    cur: Any = style
    for k in path:
        if not isinstance(cur, dict) or k not in cur:
            return default
        cur = cur[k]
    return cur


def _load_yaml_or_json(path: Path) -> dict[str, Any]:
    raw = path.read_text(encoding="utf-8")
    if path.suffix.lower() in {".yaml", ".yml"}:
        return yaml.safe_load(raw) or {}
    return json.loads(raw)


def _load_font(*, size: int) -> ImageFont.ImageFont:
    # Prefer explicit font path for determinism (optional).
    font_path = os.getenv("EIA_GEN_FONT_PATH", "").strip() or os.getenv("EIA_GEN_WATERMARK_FONT_PATH", "").strip()
    if font_path:
        p = Path(font_path).expanduser()
        if p.exists():
            try:
                return ImageFont.truetype(str(p), size)
            except Exception:
                pass
    return ImageFont.load_default()


def _world_to_px(
    x: float,
    y: float,
    *,
    bbox: tuple[float, float, float, float],
    ref_size: tuple[int, int],
) -> tuple[float, float]:
    minx, miny, maxx, maxy = bbox
    w, h = ref_size
    if maxx == minx or maxy == miny:
        return (0.0, 0.0)
    x_px = (x - minx) / (maxx - minx) * w
    y_px = (maxy - y) / (maxy - miny) * h
    return (x_px, y_px)


def _points_to_px(
    points: list[list[float]],
    *,
    mode: CoordinateMode,
    bbox: tuple[float, float, float, float] | None,
    ref_size: tuple[int, int],
    render_size: tuple[int, int],
) -> list[tuple[float, float]]:
    out: list[tuple[float, float]] = []
    rw, rh = render_size
    fw, fh = ref_size
    sx = (rw / fw) if fw else 1.0
    sy = (rh / fh) if fh else 1.0
    for xy in points:
        x = float(xy[0])
        y = float(xy[1])
        if mode == "PIXEL":
            out.append((x, y))
        else:
            assert bbox is not None
            x0, y0 = _world_to_px(x, y, bbox=bbox, ref_size=ref_size)
            out.append((x0 * sx, y0 * sy))
    return out


def _ellipse_perimeter(rx: float, ry: float) -> float:
    # Ramanujan approximation (good enough for dashed rendering).
    a = abs(rx)
    b = abs(ry)
    if a == 0 or b == 0:
        return 0.0
    return math.pi * (3 * (a + b) - math.sqrt((3 * a + b) * (a + 3 * b)))


def _draw_dashed_ellipse(
    draw: ImageDraw.ImageDraw,
    *,
    bbox: tuple[int, int, int, int],
    stroke_rgb: tuple[int, int, int],
    stroke_px: int,
    dash_on_px: int,
    dash_off_px: int,
) -> None:
    x0, y0, x1, y1 = bbox
    rx = (x1 - x0) / 2.0
    ry = (y1 - y0) / 2.0
    per = _ellipse_perimeter(rx, ry)
    if per <= 0:
        draw.ellipse([x0, y0, x1, y1], outline=stroke_rgb, width=stroke_px)
        return
    on = max(1, int(dash_on_px))
    off = max(0, int(dash_off_px))
    deg_per_px = 360.0 / per
    on_deg = on * deg_per_px
    off_deg = off * deg_per_px
    angle = 0.0
    # Draw short arcs around ellipse.
    while angle < 360.0:
        end = min(360.0, angle + on_deg)
        draw.arc([x0, y0, x1, y1], start=angle, end=end, fill=stroke_rgb, width=stroke_px)
        angle = end + off_deg


def _draw_polygon(
    img: Image.Image,
    draw: ImageDraw.ImageDraw,
    layer: dict[str, Any],
    *,
    mode: CoordinateMode,
    bbox,
    ref_size,
    render_size,
    style_tokens: STYLE,
):
    pts_raw = layer.get("points") or []
    if not pts_raw:
        return
    pts = _points_to_px(pts_raw, mode=mode, bbox=bbox, ref_size=ref_size, render_size=render_size)

    style = layer.get("style") or {}
    fill_rgba = _as_rgba(
        style.get("fill_rgba"),
        default=tuple(_style_get(style_tokens, ["colors", "overlay_fill_rgba"], [60, 90, 140, 80])),
    )
    stroke_px = int(style.get("stroke_px") or _style_get(style_tokens, ["strokes", "boundary_px"], 6))
    stroke_rgb = _hex_to_rgb(
        _as_str(style.get("stroke")) or _as_str(_style_get(style_tokens, ["colors", "overlay_stroke"], "")),
        default=(47, 90, 165),
    )

    # fill
    overlay = Image.new("RGBA", img.size, (255, 255, 255, 0))
    od = ImageDraw.Draw(overlay)
    od.polygon(pts, fill=fill_rgba)
    img.alpha_composite(overlay)

    # stroke
    draw.line(pts + [pts[0]], fill=stroke_rgb, width=stroke_px, joint="curve")


def _draw_label(
    draw: ImageDraw.ImageDraw,
    layer: dict[str, Any],
    *,
    mode: CoordinateMode,
    bbox,
    ref_size,
    render_size,
    style_tokens: STYLE,
):
    text = _as_str(layer.get("text"))
    if not text:
        return

    anchor_raw = layer.get("anchor")
    if not anchor_raw:
        return
    anchor_pts = _points_to_px([anchor_raw], mode=mode, bbox=bbox, ref_size=ref_size, render_size=render_size)
    ax, ay = anchor_pts[0]

    box_xy = layer.get("box_xy")
    if box_xy:
        bx, by = _points_to_px([box_xy], mode=mode, bbox=bbox, ref_size=ref_size, render_size=render_size)[0]
    else:
        off = layer.get("offset") or [40, -60]
        bx = ax + float(off[0])
        by = ay + float(off[1])

    style = layer.get("style") or {}
    pad_px = int(style.get("pad_px") or _style_get(style_tokens, ["callout_labels", "box", "pad_px"], 14))
    radius_px = int(style.get("radius_px") or _style_get(style_tokens, ["callout_labels", "box", "radius_px"], 18))
    bg_rgba = _as_rgba(
        style.get("bg_rgba"),
        default=tuple(_style_get(style_tokens, ["callout_labels", "box", "bg_rgba"], [255, 255, 255, 235])),
    )

    font = _load_font(size=28)
    # bbox for text
    try:
        tw = int(draw.textlength(text, font=font))
        th = 28
    except Exception:
        tw, th = draw.textsize(text, font=font)

    x0 = int(bx)
    y0 = int(by)
    x1 = x0 + tw + pad_px * 2
    y1 = y0 + th + pad_px * 2

    # leader line
    leader_rgb = _hex_to_rgb(_as_str(style.get("leader_stroke")), default=(255, 255, 255))
    leader_px = int(style.get("leader_px") or _style_get(style_tokens, ["strokes", "leader_px"], 3))
    draw.line([(ax, ay), (x0, y0 + (y1 - y0) / 2)], fill=leader_rgb, width=leader_px)

    # box (RGBA overlay)
    # ImageDraw on RGBA image supports RGBA fill; caller should pass draw bound to RGBA image.
    draw.rounded_rectangle([x0, y0, x1, y1], radius=radius_px, fill=bg_rgba, outline=None)
    draw.text((x0 + pad_px, y0 + pad_px), text, fill=(0, 0, 0), font=font)


def _draw_number_badge(
    draw: ImageDraw.ImageDraw,
    layer: dict[str, Any],
    *,
    mode: CoordinateMode,
    bbox,
    ref_size,
    render_size,
    style_tokens: STYLE,
):
    text = _as_str(layer.get("text"))
    if not text:
        return
    center_raw = layer.get("center")
    if not center_raw:
        return
    cx, cy = _points_to_px([center_raw], mode=mode, bbox=bbox, ref_size=ref_size, render_size=render_size)[0]

    rx, ry = (44, 30)
    radius = layer.get("radius")
    if isinstance(radius, (list, tuple)) and len(radius) == 2:
        rx, ry = int(radius[0]), int(radius[1])

    style = layer.get("style") or {}
    fill_rgba = _as_rgba(
        style.get("fill_rgba"),
        default=tuple(_style_get(style_tokens, ["number_badges", "ellipse", "fill_rgba"], [255, 213, 79, 160])),
    )
    stroke_rgb = _hex_to_rgb(
        _as_str(style.get("stroke")) or _as_str(_style_get(style_tokens, ["number_badges", "ellipse", "stroke"], "")),
        default=(47, 90, 165),
    )
    stroke_px = int(style.get("stroke_px") or _style_get(style_tokens, ["number_badges", "ellipse", "stroke_px"], 4))
    dashed_pattern = style.get("dashed_pattern") or _style_get(style_tokens, ["number_badges", "ellipse", "dashed_pattern"], None)

    x0, y0, x1, y1 = int(cx - rx), int(cy - ry), int(cx + rx), int(cy + ry)
    draw.ellipse([x0, y0, x1, y1], fill=fill_rgba, outline=None)
    if isinstance(dashed_pattern, (list, tuple)) and len(dashed_pattern) == 2:
        _draw_dashed_ellipse(
            draw,
            bbox=(x0, y0, x1, y1),
            stroke_rgb=stroke_rgb,
            stroke_px=stroke_px,
            dash_on_px=int(dashed_pattern[0]),
            dash_off_px=int(dashed_pattern[1]),
        )
    else:
        draw.ellipse([x0, y0, x1, y1], outline=stroke_rgb, width=stroke_px)

    font = _load_font(size=28)
    try:
        tw = int(draw.textlength(text, font=font))
        th = 28
    except Exception:
        tw, th = draw.textsize(text, font=font)
    draw.text((int(cx - tw / 2), int(cy - th / 2)), text, fill=(0, 0, 0), font=font)


def _draw_legend(
    draw: ImageDraw.ImageDraw,
    layer: dict[str, Any],
    *,
    mode: CoordinateMode,
    bbox,
    ref_size,
    render_size,
    style_tokens: STYLE,
):
    box_xy = layer.get("box_xy")
    if not box_xy:
        return
    x0, y0 = _points_to_px([box_xy], mode=mode, bbox=bbox, ref_size=ref_size, render_size=render_size)[0]
    x0, y0 = int(x0), int(y0)

    title = _as_str(layer.get("title") or "범례")
    items = layer.get("items") or []
    if not items:
        return

    pad = 16
    radius = 18
    sw = 26
    gap = 10
    line_h = 34

    font_t = _load_font(size=28)
    font_i = _load_font(size=24)

    # box size heuristic
    max_text = 0
    for it in items:
        lab = _as_str((it or {}).get("label"))
        max_text = max(max_text, len(lab))
    w = min(700, pad * 2 + sw + gap + max(200, max_text * 18))
    h = pad * 2 + line_h * (len(items) + 1)

    bg_rgba = _as_rgba(
        (layer.get("style") or {}).get("bg_rgba"),
        default=tuple(_style_get(style_tokens, ["legend", "box", "bg_rgba"], [255, 255, 255, 220])),
    )
    draw.rounded_rectangle([x0, y0, x0 + w, y0 + h], radius=radius, fill=bg_rgba)
    draw.text((x0 + pad, y0 + pad), title, fill=(0, 0, 0), font=font_t)

    y = y0 + pad + line_h
    for it in items:
        it = it or {}
        lab = _as_str(it.get("label"))
        swatch = it.get("swatch") or {}
        fill = _as_rgba(swatch.get("fill_rgba"), default=(60, 90, 140, 80))
        stroke = _hex_to_rgb(_as_str(swatch.get("stroke")), default=(47, 90, 165))
        draw.rectangle([x0 + pad, y + 4, x0 + pad + sw, y + 4 + sw], fill=fill, outline=stroke, width=2)
        draw.text((x0 + pad + sw + gap, y), lab, fill=(0, 0, 0), font=font_i)
        y += line_h


def main() -> None:
    ap = argparse.ArgumentParser(description="Annotate an image with polygon/label/badge/legend layers (prototype).")
    ap.add_argument("--image", type=Path, required=True, help="Input image (PNG/JPG)")
    ap.add_argument("--annotations", type=Path, required=True, help="Annotations spec (YAML/JSON)")
    ap.add_argument(
        "--style",
        type=Path,
        default=Path("config/figure_style.yaml"),
        help="Style tokens yaml (repo-relative by default)",
    )
    ap.add_argument("--out", type=Path, required=True, help="Output PNG path")
    args = ap.parse_args()

    repo_root = Path(__file__).resolve().parents[1]  # eia-gen/
    img_path = args.image.expanduser().resolve()
    ann_path = args.annotations.expanduser().resolve()
    style_path = args.style
    if not style_path.is_absolute():
        style_path = (repo_root / style_path).resolve()

    if not img_path.exists():
        raise SystemExit(f"image not found: {img_path}")
    if not ann_path.exists():
        raise SystemExit(f"annotations not found: {ann_path}")
    if not style_path.exists():
        raise SystemExit(f"style yaml not found: {style_path}")

    ann = _load_yaml_or_json(ann_path)
    style = _load_yaml_or_json(style_path)
    style_tokens: STYLE = style or {}

    mode = _as_str(ann.get("coordinate_mode") or "PIXEL").upper()
    if mode not in ("PIXEL", "WORLD_LINEAR_BBOX"):
        raise SystemExit(f"Unsupported coordinate_mode: {mode}")
    coord_mode: CoordinateMode = mode  # type: ignore[assignment]

    img = Image.open(img_path).convert("RGBA")
    w, h = img.size
    render_size = (w, h)
    ref_size = render_size

    bbox = None
    if coord_mode == "WORLD_LINEAR_BBOX":
        raw_bbox = ann.get("bbox")
        if not (isinstance(raw_bbox, (list, tuple)) and len(raw_bbox) == 4):
            raise SystemExit("WORLD_LINEAR_BBOX requires bbox: [minx, miny, maxx, maxy]")
        bbox = (float(raw_bbox[0]), float(raw_bbox[1]), float(raw_bbox[2]), float(raw_bbox[3]))

        raw_size = ann.get("image_size")
        if isinstance(raw_size, str):
            m = re.fullmatch(r"\s*(\d+)\s*[xX]\s*(\d+)\s*", raw_size)
            if m:
                raw_size = [int(m.group(1)), int(m.group(2))]
        if isinstance(raw_size, (list, tuple)) and len(raw_size) == 2:
            iw = _as_int(raw_size[0]) or w
            ih = _as_int(raw_size[1]) or h
            ref_size = (iw, ih)

    draw = ImageDraw.Draw(img, "RGBA")
    layers = ann.get("layers") or []
    for layer in layers:
        if not isinstance(layer, dict):
            continue
        t = _as_str(layer.get("type")).lower()
        if t == "polygon":
            _draw_polygon(
                img,
                draw,
                layer,
                mode=coord_mode,
                bbox=bbox,
                ref_size=ref_size,
                render_size=render_size,
                style_tokens=style_tokens,
            )
        elif t == "label":
            _draw_label(
                draw, layer, mode=coord_mode, bbox=bbox, ref_size=ref_size, render_size=render_size, style_tokens=style_tokens
            )
        elif t == "number_badge":
            _draw_number_badge(
                draw,
                layer,
                mode=coord_mode,
                bbox=bbox,
                ref_size=ref_size,
                render_size=render_size,
                style_tokens=style_tokens,
            )
        elif t == "legend":
            _draw_legend(
                draw, layer, mode=coord_mode, bbox=bbox, ref_size=ref_size, render_size=render_size, style_tokens=style_tokens
            )

    out = args.out.expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    img.convert("RGB").save(out, format="PNG", optimize=True)
    print(f"WROTE: {out}")


if __name__ == "__main__":
    main()
