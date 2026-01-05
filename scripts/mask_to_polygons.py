#!/usr/bin/env python3
from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import numpy as np
import yaml
from PIL import Image, ImageDraw


@dataclass(frozen=True)
class Params:
    closing_radius: int
    opening_radius: int
    min_area_px: int
    hole_area_px: int
    approx_tol: float
    max_polygons: int


def _as_list(s: str | None) -> list[str]:
    if not s:
        return []
    tokens: list[str] = []
    for chunk in s.replace(";", ",").split(","):
        t = chunk.strip()
        if t:
            tokens.append(t)
    return tokens


def _postprocess_mask(
    mask: np.ndarray,
    *,
    closing_radius: int,
    opening_radius: int,
    min_area_px: int,
    hole_area_px: int,
) -> np.ndarray:
    import inspect

    from skimage.morphology import closing, disk, opening, remove_small_holes, remove_small_objects

    m = mask.astype(bool)
    if closing_radius > 0:
        m = closing(m, disk(closing_radius))
    if hole_area_px > 0:
        # skimage>=0.26: area_threshold is deprecated; max_size is the new kw.
        if "max_size" in inspect.signature(remove_small_holes).parameters:
            m = remove_small_holes(m, max_size=hole_area_px)
        else:
            m = remove_small_holes(m, area_threshold=hole_area_px)
    if min_area_px > 0:
        # skimage>=0.26: min_size is deprecated; max_size is the new kw.
        if "max_size" in inspect.signature(remove_small_objects).parameters:
            m = remove_small_objects(m, max_size=min_area_px)
        else:
            m = remove_small_objects(m, min_size=min_area_px)
    if opening_radius > 0:
        m = opening(m, disk(opening_radius))
    return m.astype(bool)


def _contour_to_points(contour: np.ndarray, *, approx_tol: float) -> list[list[int]]:
    # contour is (row, col) floats; convert to (x, y)
    pts = [(float(p[1]), float(p[0])) for p in contour]
    if not pts:
        return []

    # simple polygon approximation by skipping points within tol distance
    if approx_tol > 0:
        out: list[tuple[float, float]] = [pts[0]]
        for x, y in pts[1:]:
            lx, ly = out[-1]
            if (x - lx) ** 2 + (y - ly) ** 2 >= approx_tol**2:
                out.append((x, y))
        pts2 = out
    else:
        pts2 = pts

    # close loop
    if len(pts2) >= 3:
        x0, y0 = pts2[0]
        x1, y1 = pts2[-1]
        if (x0 - x1) ** 2 + (y0 - y1) ** 2 >= 1.0:
            pts2.append((x0, y0))

    return [[int(round(x)), int(round(y))] for x, y in pts2]


def _render_debug(
    base_rgb: np.ndarray,
    *,
    mask: np.ndarray,
    polys: list[list[list[int]]],
    out_dir: Path,
) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)

    # 1) mask image
    m = (mask.astype(np.uint8) * 255)
    Image.fromarray(m, mode="L").save(out_dir / "mask_pp.png")

    # 2) overlay preview with contours
    img = Image.fromarray(base_rgb, mode="RGB").convert("RGBA")
    draw = ImageDraw.Draw(img, "RGBA")
    for pts in polys:
        if len(pts) < 3:
            continue
        draw.line([tuple(p) for p in pts], fill=(255, 255, 255, 220), width=4)
    img.convert("RGB").save(out_dir / "contours_overlay.png", format="PNG", optimize=True)


def main() -> None:
    ap = argparse.ArgumentParser(
        description=(
            "Convert a binary mask (white=foreground) into polygon layers and emit an annotation spec "
            "usable by scripts/annotate_image.py. (No inference: segmentation-only.)"
        )
    )
    ap.add_argument("--mask", type=Path, required=True, help="Binary mask image (PNG/JPG). White=foreground.")
    ap.add_argument("--base-image", type=Path, default=None, help="Optional base image path for metadata/debug only.")
    ap.add_argument("--invert", action="store_true", help="Invert mask (useful if foreground is black).")
    ap.add_argument("--out-annotations", type=Path, required=True, help="Output annotations YAML path")
    ap.add_argument("--labels", type=str, default="", help="Optional labels for polygons (comma/; separated)")
    ap.add_argument("--max-polygons", type=int, default=5, help="Keep up to N largest polygons")
    ap.add_argument(
        "--roi",
        nargs=4,
        type=int,
        default=None,
        metavar=("X0", "Y0", "X1", "Y1"),
        help="Optional ROI rectangle in pixel coords. Anything outside is ignored. (x1/y1 are exclusive)",
    )
    ap.add_argument(
        "--exclude-rect",
        nargs=4,
        type=int,
        action="append",
        default=[],
        metavar=("X0", "Y0", "X1", "Y1"),
        help="Optional exclusion rectangle(s) in pixel coords. Can be repeated.",
    )
    ap.add_argument(
        "--min-area-ratio",
        type=float,
        default=0.001,
        help="Drop components smaller than ratio*(W*H)",
    )
    ap.add_argument("--closing-radius", type=int, default=6, help="Binary closing radius (px)")
    ap.add_argument("--opening-radius", type=int, default=2, help="Binary opening radius (px)")
    ap.add_argument(
        "--hole-area-ratio",
        type=float,
        default=0.001,
        help="Fill holes smaller than ratio*(W*H)",
    )
    ap.add_argument(
        "--approx-tol",
        type=float,
        default=6.0,
        help="Point decimation tolerance (px). Higher => fewer points.",
    )
    ap.add_argument("--emit-number-badges", action="store_true", help="Emit number_badge layers at polygon centroids")
    ap.add_argument("--badge-prefix", type=str, default="#", help="Prefix for number badge text (default: '#')")
    ap.add_argument("--emit-legend", action="store_true", help="Emit a legend box layer (bottom-right by default)")
    ap.add_argument(
        "--legend-box",
        nargs=2,
        type=int,
        default=None,
        metavar=("X", "Y"),
        help="Legend box top-left position (PIXEL). Default: bottom-right with margin.",
    )
    ap.add_argument("--debug-dir", type=Path, default=None, help="Optional debug outputs (mask/overlay)")
    args = ap.parse_args()

    try:
        # Trigger optional dependency check early for a clearer error message.
        import skimage  # noqa: F401
    except Exception as e:
        raise SystemExit(
            "Missing optional dependency: scikit-image.\n"
            "Install one of:\n"
            "  - cd eia-gen && ./.venv/bin/python -m pip install scikit-image\n"
            "  - cd eia-gen && ./.venv/bin/python -m pip install -e '.[image-tools]'\n"
            f"Original error: {e}"
        ) from e

    mask_path = args.mask.expanduser().resolve()
    if not mask_path.exists():
        raise SystemExit(f"mask not found: {mask_path}")

    base_path = None
    if args.base_image:
        base_path = args.base_image.expanduser().resolve()
        if not base_path.exists():
            raise SystemExit(f"base image not found: {base_path}")

    m_img = Image.open(mask_path).convert("L")
    m_arr = np.asarray(m_img)
    h, w = m_arr.shape[0], m_arr.shape[1]

    # binary threshold (anything bright is foreground)
    mask = m_arr >= 128
    if args.invert:
        mask = ~mask

    def _clamp_rect(r: list[int] | tuple[int, int, int, int] | None) -> tuple[int, int, int, int] | None:
        if not r:
            return None
        x0, y0, x1, y1 = int(r[0]), int(r[1]), int(r[2]), int(r[3])
        if x1 < x0:
            x0, x1 = x1, x0
        if y1 < y0:
            y0, y1 = y1, y0
        x0 = max(0, min(w, x0))
        x1 = max(0, min(w, x1))
        y0 = max(0, min(h, y0))
        y1 = max(0, min(h, y1))
        if x1 <= x0 or y1 <= y0:
            return None
        return (x0, y0, x1, y1)

    roi = _clamp_rect(args.roi)
    exclude_rects = [r for r in (_clamp_rect(x) for x in (args.exclude_rect or [])) if r is not None]

    if roi is not None:
        x0, y0, x1, y1 = roi
        m2 = np.zeros_like(mask, dtype=bool)
        m2[y0:y1, x0:x1] = mask[y0:y1, x0:x1]
        mask = m2

    for r in exclude_rects:
        x0, y0, x1, y1 = r
        mask[y0:y1, x0:x1] = False

    min_area_px = int(max(1, float(args.min_area_ratio) * w * h))
    hole_area_px = int(max(0, float(args.hole_area_ratio) * w * h))
    pp = _postprocess_mask(
        mask,
        closing_radius=int(args.closing_radius),
        opening_radius=int(args.opening_radius),
        min_area_px=min_area_px,
        hole_area_px=hole_area_px,
    )

    # connected components
    from skimage.measure import find_contours, label, regionprops

    lbl = label(pp)
    props = sorted(regionprops(lbl), key=lambda p: p.area, reverse=True)
    props = [p for p in props if p.area >= min_area_px][: int(args.max_polygons)]
    if not props:
        raise SystemExit(
            "No polygons detected. Try lowering --min-area-ratio, or adjusting postprocess "
            "(--closing-radius/--opening-radius/--hole-area-ratio), or set --roi/--exclude-rect."
        )

    labels = _as_list(args.labels)
    layers: list[dict[str, Any]] = []
    polys: list[list[list[int]]] = []
    legend_labels: list[str] = []

    for i, p in enumerate(props, start=1):
        region_mask = (lbl == p.label)
        contours = find_contours(region_mask.astype(float), level=0.5)
        if not contours:
            continue
        contour = max(contours, key=lambda c: c.shape[0])
        pts = _contour_to_points(contour, approx_tol=float(args.approx_tol))
        if len(pts) < 4:
            continue
        polys.append(pts)

        layers.append({"type": "polygon", "id": f"POLY-{i:02d}", "points": pts})

        label_text = labels[i - 1] if (i - 1) < len(labels) else f"영역 {i}"
        cy, cx = p.centroid  # (row, col)
        layers.append(
            {
                "type": "label",
                "id": f"LBL-{i:02d}",
                "text": str(label_text),
                "anchor": [int(round(cx)), int(round(cy))],
                "offset": [40, -60],
            }
        )
        legend_labels.append(str(label_text))

        if args.emit_number_badges:
            layers.append(
                {
                    "type": "number_badge",
                    "id": f"N-{i:02d}",
                    "text": f"{args.badge_prefix}{i}",
                    "center": [int(round(cx)), int(round(cy))],
                }
            )

    if args.emit_legend and legend_labels:
        # Heuristic legend size similar to scripts/annotate_image.py
        pad = 16
        sw = 26
        gap = 10
        line_h = 34
        max_text = max(len(x) for x in legend_labels) if legend_labels else 0
        leg_w = min(700, pad * 2 + sw + gap + max(200, max_text * 18))
        leg_h = pad * 2 + line_h * (len(legend_labels) + 1)

        if args.legend_box and len(args.legend_box) == 2:
            lx, ly = int(args.legend_box[0]), int(args.legend_box[1])
        else:
            margin = 40
            lx = max(20, int(w - leg_w - margin))
            ly = max(20, int(h - leg_h - margin))

        layers.append(
            {
                "type": "legend",
                "id": "LEG-01",
                "title": "범례",
                "box_xy": [lx, ly],
                "items": [{"label": lab} for lab in legend_labels],
            }
        )

    spec: dict[str, Any] = {
        "schema_version": "1.0",
        "coordinate_mode": "PIXEL",
        "image_path": str(base_path or mask_path),
        "layers": layers,
        "extraction": {
            "source": "MASK_BINARY",
            "mask_path": str(mask_path),
            "invert": bool(args.invert),
            "spatial_filter": {
                "roi": list(roi) if roi is not None else None,
                "exclude_rects": [list(r) for r in exclude_rects],
            },
            "postprocess": {
                "closing_radius": int(args.closing_radius),
                "opening_radius": int(args.opening_radius),
                "min_area_ratio": float(args.min_area_ratio),
                "hole_area_ratio": float(args.hole_area_ratio),
                "approx_tol": float(args.approx_tol),
                "max_polygons": int(args.max_polygons),
            },
            "emit": {
                "number_badges": bool(args.emit_number_badges),
                "legend": bool(args.emit_legend),
            },
        },
    }

    out_ann = args.out_annotations.expanduser().resolve()
    out_ann.parent.mkdir(parents=True, exist_ok=True)
    out_ann.write_text(yaml.safe_dump(spec, sort_keys=False, allow_unicode=True), encoding="utf-8")
    print(f"WROTE: {out_ann}")

    if args.debug_dir:
        if base_path and base_path.exists():
            base_rgb = np.asarray(Image.open(base_path).convert("RGB"))
        else:
            base_rgb = np.dstack([m_arr, m_arr, m_arr]).astype(np.uint8)
        _render_debug(base_rgb, mask=pp, polys=polys, out_dir=args.debug_dir.expanduser().resolve())
        print(f"WROTE: {args.debug_dir}/mask_pp.png")
        print(f"WROTE: {args.debug_dir}/contours_overlay.png")


if __name__ == "__main__":
    main()
