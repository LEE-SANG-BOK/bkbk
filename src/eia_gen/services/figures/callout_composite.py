from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml
from PIL import Image, ImageDraw, ImageFont, ImageOps


@dataclass(frozen=True)
class CalloutItem:
    path: Path | None
    caption: str


def _as_str(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _style_get(style: dict[str, Any], path: list[str], default: Any) -> Any:
    cur: Any = style
    for key in path:
        if not isinstance(cur, dict) or key not in cur:
            return default
        cur = cur[key]
    return cur


def _load_font(*, size: int, font_path: str | None) -> ImageFont.ImageFont:
    # Prefer explicit font path for determinism; otherwise allow env injection.
    p = _as_str(font_path)
    if not p:
        p = os.getenv("EIA_GEN_FONT_PATH", "").strip()
    if p:
        fp = Path(p).expanduser()
        if fp.exists():
            try:
                return ImageFont.truetype(str(fp), size)
            except Exception:
                pass
    return ImageFont.load_default()


def _wrap_text(
    draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_width_px: int
) -> list[str]:
    # Character-wise wrapping for mixed Korean/English without spaces.
    lines: list[str] = []
    for paragraph in (text or "").split("\n"):
        paragraph = paragraph.strip()
        if not paragraph:
            lines.append("")
            continue
        cur = ""
        for ch in paragraph:
            candidate = cur + ch
            try:
                w = draw.textlength(candidate, font=font)
            except Exception:
                w, _ = draw.textsize(candidate, font=font)
            if w <= max_width_px:
                cur = candidate
            else:
                if cur:
                    lines.append(cur)
                cur = ch
        if cur:
            lines.append(cur)
    return lines


def _cover_fit(img: Image.Image, *, target_w: int, target_h: int) -> Image.Image:
    img = ImageOps.exif_transpose(img)
    if img.mode not in ("RGB", "RGBA"):
        img = img.convert("RGB")
    if img.mode == "RGBA":
        bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
        img = Image.alpha_composite(bg, img).convert("RGB")

    w, h = img.size
    if w <= 0 or h <= 0:
        raise ValueError("Invalid image size")

    scale = max(target_w / w, target_h / h)
    new_w = max(1, int(round(w * scale)))
    new_h = max(1, int(round(h * scale)))
    try:
        resample = Image.Resampling.LANCZOS  # type: ignore[attr-defined]
    except Exception:
        resample = Image.LANCZOS
    resized = img.resize((new_w, new_h), resample=resample)

    left = (new_w - target_w) // 2
    top = (new_h - target_h) // 2
    return resized.crop((left, top, left + target_w, top + target_h))


def _grid_from_count(n: int) -> int:
    if n <= 2:
        return 2
    if n <= 4:
        return 4
    if n <= 6:
        return 6
    raise ValueError(f"CALLOUT_COMPOSITE supports up to 6 images (got {n})")


def _rows_cols(grid: int) -> tuple[int, int]:
    if grid == 2:
        return (1, 2)
    if grid == 4:
        return (2, 2)
    if grid == 6:
        return (2, 3)  # 3x2
    raise ValueError(f"Unsupported grid: {grid} (expected 2/4/6)")


def _default_style_path() -> Path:
    # Best-effort repo-relative default.
    try:
        repo_root = Path(__file__).resolve().parents[4]
        cand = repo_root / "config" / "figure_style.yaml"
        if cand.exists():
            return cand
    except Exception:
        pass
    return Path("config/figure_style.yaml")


def load_figure_style(style_path: Path | None = None) -> dict[str, Any]:
    style_path = (style_path or _default_style_path()).expanduser()
    if not style_path.exists():
        raise FileNotFoundError(f"style yaml not found: {style_path}")
    return yaml.safe_load(style_path.read_text(encoding="utf-8")) or {}


def compose_callout_composite(
    *,
    items: list[CalloutItem],
    out_path: Path,
    title: str = "",
    grid: int | None = None,
    style: dict[str, Any] | None = None,
    style_path: Path | None = None,
    font_path: str | None = None,
) -> dict[str, Any]:
    """Compose a deterministic 'photo board' panel (CALLOUT_COMPOSITE) with fixed grids (2/4/6)."""

    if not items:
        raise ValueError("No items provided")

    # Filter missing files but keep placeholders to preserve layout.
    cleaned: list[CalloutItem] = []
    for it in items:
        p = it.path
        if p is not None:
            p = Path(p).expanduser()
        cleaned.append(CalloutItem(path=p, caption=_as_str(it.caption) or (p.name if p else "")))

    n = len(cleaned)
    if n > 6:
        raise ValueError(f"Too many images (max 6). got={n}")

    style_obj = style or load_figure_style(style_path)
    grid2 = int(grid) if grid is not None else _grid_from_count(n)
    rows, cols = _rows_cols(grid2)

    # style tokens
    panel_w = int(_style_get(style_obj, ["callout_composite", "panel", "width_px"], 2400))
    panel_h = int(_style_get(style_obj, ["callout_composite", "panel", "height_px"], 1600))
    outer_pad = int(_style_get(style_obj, ["callout_composite", "panel", "outer_pad_px"], 36))
    gap = int(_style_get(style_obj, ["callout_composite", "panel", "gap_px"], 24))
    panel_border = int(_style_get(style_obj, ["callout_composite", "panel", "border_px"], 2))
    bg_rgb = tuple(_style_get(style_obj, ["callout_composite", "panel", "bg_rgb"], [255, 255, 255]))

    cell_border = int(_style_get(style_obj, ["callout_composite", "slot", "cell_border_px"], 1))
    caption_bar_h = int(_style_get(style_obj, ["callout_composite", "slot", "caption_bar_h_px"], 120))
    caption_pad = int(_style_get(style_obj, ["callout_composite", "slot", "caption_pad_px"], 12))
    caption_bg_rgb = tuple(_style_get(style_obj, ["callout_composite", "slot", "caption_bg_rgb"], [248, 248, 248]))

    badge_w = int(_style_get(style_obj, ["callout_composite", "badge", "w_px"], 64))
    badge_h = int(_style_get(style_obj, ["callout_composite", "badge", "h_px"], 44))
    badge_margin = int(_style_get(style_obj, ["callout_composite", "badge", "margin_px"], 10))
    badge_bg_rgb = tuple(_style_get(style_obj, ["callout_composite", "badge", "bg_rgb"], [0, 0, 0]))
    badge_text_rgb = tuple(_style_get(style_obj, ["callout_composite", "badge", "text_rgb"], [255, 255, 255]))

    # layout sizes
    usable_w = panel_w - 2 * outer_pad - (cols - 1) * gap
    usable_h = panel_h - 2 * outer_pad - (rows - 1) * gap
    if usable_w <= 100 or usable_h <= 100:
        raise ValueError("Panel too small; increase panel size or reduce padding/gap.")
    cell_w = usable_w // cols
    cell_h = usable_h // rows
    img_h = cell_h - caption_bar_h
    if img_h <= 80:
        raise ValueError("caption_bar_h_px too large; image area too small.")

    # create panel
    panel = Image.new("RGB", (panel_w, panel_h), bg_rgb)
    draw = ImageDraw.Draw(panel)

    # panel border
    if panel_border > 0:
        border_rgb = (0, 0, 0)
        for i in range(panel_border):
            draw.rectangle([i, i, panel_w - 1 - i, panel_h - 1 - i], outline=border_rgb)

    font_caption = _load_font(size=24, font_path=font_path)
    font_badge = _load_font(size=24, font_path=font_path)
    font_title = _load_font(size=32, font_path=font_path)

    title2 = _as_str(title)
    if title2:
        draw.text((outer_pad, 10), title2, font=font_title, fill=(0, 0, 0))

    # pad with empty slots
    total_slots = rows * cols
    extended = cleaned + [CalloutItem(path=None, caption="사진 미제출")] * (total_slots - len(cleaned))

    idx = 0
    for row_idx in range(rows):
        for col_idx in range(cols):
            x0 = outer_pad + col_idx * (cell_w + gap)
            y0 = outer_pad + row_idx * (cell_h + gap)
            x1 = x0 + cell_w
            y1 = y0 + cell_h

            # cell border
            if cell_border > 0:
                for i in range(cell_border):
                    draw.rectangle([x0 + i, y0 + i, x1 - 1 - i, y1 - 1 - i], outline=(0, 0, 0))

            img_y1 = y0 + img_h
            item = extended[idx]
            if item.path and item.path.exists():
                img = Image.open(item.path)
                fitted = _cover_fit(img, target_w=cell_w, target_h=img_h)
                panel.paste(fitted, (x0, y0))
            else:
                panel.paste(Image.new("RGB", (cell_w, img_h), (245, 245, 245)), (x0, y0))

            # caption bar
            draw.rectangle([x0, img_y1, x1, y1], fill=caption_bg_rgb)
            max_w = cell_w - 2 * caption_pad
            lines = _wrap_text(draw, item.caption, font_caption, max_w)[:2]
            tx = x0 + caption_pad
            ty = img_y1 + caption_pad
            for ln in lines:
                draw.text((tx, ty), ln, font=font_caption, fill=(0, 0, 0))
                ty += int(24 * 1.25)

            # badge
            bx = x0 + badge_margin
            by = y0 + badge_margin
            draw.rounded_rectangle([bx, by, bx + badge_w, by + badge_h], radius=10, fill=badge_bg_rgb)
            label = f"({idx + 1})"
            try:
                tw = int(draw.textlength(label, font=font_badge))
                th = 24
            except Exception:
                tw, th = draw.textsize(label, font=font_badge)
            draw.text(
                (bx + (badge_w - tw) // 2, by + (badge_h - th) // 2 - 2),
                label,
                font=font_badge,
                fill=badge_text_rgb,
            )

            idx += 1

    out_path = Path(out_path).expanduser()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    panel.save(out_path, format="PNG", optimize=True)

    return {
        "grid": grid2,
        "rows": rows,
        "cols": cols,
        "panel_px": [panel_w, panel_h],
        "inputs": [{"path": str(it.path) if it.path else "", "caption": it.caption} for it in cleaned],
        "title": title2,
    }

