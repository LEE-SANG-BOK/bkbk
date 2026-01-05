from __future__ import annotations

import hashlib
import os
import re
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Any

import yaml
from PIL import Image, ImageDraw, ImageFont, ImageOps


_PAGE_RE = re.compile(r"(?i)(?:^|\b)(?:PDF_PAGE|FROM_PDF_PAGE|PAGE)\s*[:=]\s*(\d+)\b")
_FILE_PAGE_RE = re.compile(r"(?i)(?:[#?@]page=)(\d+)\b")
_AUTO_CROP_RE = re.compile(r"(?i)^auto(?:\s*[:=]\s*(.*))?$")
_AUTH_REFERENCE_RE = re.compile(r"(?i)(?:^|\b)AUTHENTICITY\s*[:=]\s*REFERENCE\b")
_AUTH_SSOT_SAMPLE_RE = re.compile(r"(?i)(?:^|\b)AUTHENTICITY\s*[:=]\s*SSOT_SAMPLE\b")
_GEN_TARGET_DPI_RE = re.compile(r"(?i)(?:^|\b)(?:TARGET_DPI|DPI)\s*[:=]\s*(\d+)\b")
_GEN_MAX_WIDTH_RE = re.compile(r"(?i)(?:^|\b)(?:MAX_WIDTH_PX|MAX_PX|MAX_WIDTH)\s*[:=]\s*(\d+)\b")
_GEN_OUTPUT_FORMAT_RE = re.compile(r"(?i)(?:^|\b)(?:OUTPUT_FORMAT|FORMAT)\s*[:=]\s*(PNG|JPG|JPEG)\b")
_GEN_JPEG_QUALITY_RE = re.compile(r"(?i)(?:^|\b)(?:JPEG_QUALITY|JPG_QUALITY|QUALITY)\s*[:=]\s*(\d+)\b")


@dataclass(frozen=True)
class MaterializeOptions:
    out_dir: Path
    fig_id: str
    gen_method: str | None = None
    crop: str | None = None
    width_mm: float | None = None
    asset_type: str | None = None
    # Quality controls (trade-off: size vs sharpness)
    target_dpi: int | None = None
    max_width_px: int | None = None
    output_format: str | None = None  # "PNG" | "JPEG" (None=auto)
    jpeg_quality: int | None = None


@dataclass(frozen=True)
class MaterializeResult:
    path: Path
    pdf_page: int | None = None
    pdf_page_source: str | None = None  # "explicit" | "heuristic"
    meta: dict[str, Any] | None = None


def _sha1_text(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()


_SHA256_FILE_CACHE: dict[tuple[str, int, int], str] = {}


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _sha256_file_cached(path: Path) -> str:
    """Compute sha256 once per (path, mtime_ns, size) for the current process."""
    st = path.stat()
    key = (str(path.resolve()), int(st.st_mtime_ns), int(st.st_size))
    cached = _SHA256_FILE_CACHE.get(key)
    if cached:
        return cached
    digest = _sha256_file(path)
    _SHA256_FILE_CACHE[key] = digest
    # Keep memory bounded (best-effort).
    if len(_SHA256_FILE_CACHE) > 256:
        _SHA256_FILE_CACHE.clear()
    return digest


def _cfg_get(cfg: dict[str, Any], path: list[str], default: Any) -> Any:
    cur: Any = cfg
    for k in path:
        if not isinstance(cur, dict) or k not in cur:
            return default
        cur = cur[k]
    return cur


def _default_asset_normalization_path() -> Path:
    # Best-effort repo-relative default.
    try:
        repo_root = Path(__file__).resolve().parents[4]
        cand = repo_root / "config" / "asset_normalization.yaml"
        if cand.exists():
            return cand
    except Exception:
        pass
    return Path("config/asset_normalization.yaml")


@lru_cache(maxsize=4)
def _load_asset_normalization_cfg(path_str: str) -> dict[str, Any]:
    path = Path(path_str).expanduser()
    if not path.exists():
        return {}
    obj = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    return obj if isinstance(obj, dict) else {}


def _load_font(*, size: int, font_path: str | None) -> ImageFont.ImageFont:
    p = (font_path or "").strip()
    if p:
        fp = Path(p).expanduser()
        if fp.exists():
            try:
                return ImageFont.truetype(str(fp), size)
            except Exception:
                pass
    return ImageFont.load_default()


def _reference_watermark_params(gen_method: str | None) -> tuple[dict[str, Any] | None, str]:
    """Return (params, signature) for AUTHENTICITY:REFERENCE watermark."""
    gm = (gen_method or "").strip()
    if not gm or not _AUTH_REFERENCE_RE.search(gm):
        return None, ""

    cfg_path = _default_asset_normalization_path()
    cfg = _load_asset_normalization_cfg(str(cfg_path.resolve()))

    enabled = bool(_cfg_get(cfg, ["watermark", "reference", "enabled"], True))
    if not enabled:
        return None, "wm:reference:disabled"

    text = str(_cfg_get(cfg, ["watermark", "reference", "text"], "참고도(개략) - 공식 도면 아님") or "").strip()
    if not text:
        return None, "wm:reference:empty-text"

    try:
        angle = float(_cfg_get(cfg, ["watermark", "reference", "angle_deg"], 32.0))
    except Exception:
        angle = 32.0
    try:
        opacity = float(_cfg_get(cfg, ["watermark", "reference", "opacity"], 0.18))
    except Exception:
        opacity = 0.18
    try:
        font_size = int(float(_cfg_get(cfg, ["watermark", "reference", "font_size"], 42)))
    except Exception:
        font_size = 42

    font_env = str(_cfg_get(cfg, ["watermark", "reference", "font_env_var"], "EIA_GEN_WATERMARK_FONT_PATH") or "").strip()
    font_path = (os.getenv(font_env, "") or "").strip() or (os.getenv("EIA_GEN_FONT_PATH", "") or "").strip() or None

    sig = f"wm:reference:text={text}|angle={angle:.3f}|opacity={opacity:.3f}|font={font_size}|font_path={font_path or 'default'}"
    return (
        {"text": text, "angle": angle, "opacity": opacity, "font_size": font_size, "font_path": font_path},
        sig,
    )


def _ssot_sample_watermark_params(gen_method: str | None) -> tuple[dict[str, Any] | None, str]:
    """Return (params, signature) for AUTHENTICITY:SSOT_SAMPLE watermark."""
    gm = (gen_method or "").strip()
    if not gm or not _AUTH_SSOT_SAMPLE_RE.search(gm):
        return None, ""

    cfg_path = _default_asset_normalization_path()
    cfg = _load_asset_normalization_cfg(str(cfg_path.resolve()))

    enabled = bool(_cfg_get(cfg, ["watermark", "ssot_sample", "enabled"], True))
    if not enabled:
        return None, "wm:ssot_sample:disabled"

    text = str(_cfg_get(cfg, ["watermark", "ssot_sample", "text"], "SSOT 샘플(기허가) 참고용") or "").strip()
    if not text:
        return None, "wm:ssot_sample:empty-text"

    try:
        angle = float(_cfg_get(cfg, ["watermark", "ssot_sample", "angle_deg"], 32.0))
    except Exception:
        angle = 32.0
    try:
        opacity = float(_cfg_get(cfg, ["watermark", "ssot_sample", "opacity"], 0.16))
    except Exception:
        opacity = 0.16
    try:
        font_size = int(float(_cfg_get(cfg, ["watermark", "ssot_sample", "font_size"], 42)))
    except Exception:
        font_size = 42

    font_env = str(_cfg_get(cfg, ["watermark", "ssot_sample", "font_env_var"], "EIA_GEN_WATERMARK_FONT_PATH") or "").strip()
    font_path = (os.getenv(font_env, "") or "").strip() or (os.getenv("EIA_GEN_FONT_PATH", "") or "").strip() or None

    sig = f"wm:ssot_sample:text={text}|angle={angle:.3f}|opacity={opacity:.3f}|font={font_size}|font_path={font_path or 'default'}"
    return (
        {"text": text, "angle": angle, "opacity": opacity, "font_size": font_size, "font_path": font_path},
        sig,
    )


def _parse_int_from_gen_method(gen_method: str | None, *, pat: re.Pattern[str]) -> int | None:
    gm = (gen_method or "").strip()
    if not gm:
        return None
    m = pat.search(gm)
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


def _parse_output_format_from_gen_method(gen_method: str | None) -> str | None:
    gm = (gen_method or "").strip()
    if not gm:
        return None
    m = _GEN_OUTPUT_FORMAT_RE.search(gm)
    if not m:
        return None
    raw = (m.group(1) or "").strip().upper()
    if raw == "PNG":
        return "PNG"
    if raw in {"JPG", "JPEG"}:
        return "JPEG"
    return None


def _normalize_output_format(value: str | None) -> str | None:
    if value is None:
        return None
    v = str(value).strip().upper()
    if not v:
        return None
    if v == "PNG":
        return "PNG"
    if v in {"JPG", "JPEG"}:
        return "JPEG"
    return None


def _materialize_defaults(cfg: dict[str, Any], *, asset_type: str | None, is_pdf: bool) -> dict[str, Any]:
    at = (asset_type or "").strip()
    defaults: dict[str, Any] = {
        "target_dpi": _cfg_get(cfg, ["materialize", "defaults", "target_dpi"], 250),
        "max_width_px": _cfg_get(cfg, ["materialize", "defaults", "max_width_px"], 2600),
        "output_format": _cfg_get(cfg, ["materialize", "defaults", "output_format"], "PNG"),
        "jpeg_quality": _cfg_get(cfg, ["materialize", "defaults", "jpeg_quality"], 85),
    }
    if at:
        defaults["target_dpi"] = _cfg_get(cfg, ["materialize", "by_asset_type", at, "target_dpi"], defaults["target_dpi"])
        defaults["max_width_px"] = _cfg_get(cfg, ["materialize", "by_asset_type", at, "max_width_px"], defaults["max_width_px"])
        defaults["output_format"] = _cfg_get(cfg, ["materialize", "by_asset_type", at, "output_format"], defaults["output_format"])
        defaults["jpeg_quality"] = _cfg_get(cfg, ["materialize", "by_asset_type", at, "jpeg_quality"], defaults["jpeg_quality"])

    # Safe default: PDFs are rendered to PNG unless explicitly overridden (avoid line-art degradation).
    if is_pdf and _normalize_output_format(defaults.get("output_format")) == "JPEG":
        defaults["output_format"] = "PNG"

    return defaults


def _apply_text_watermark(img: Image.Image, *, text: str, angle: float, opacity: float, font_size: int, font_path: str | None) -> Image.Image:
    if not text:
        return img

    base = img.convert("RGBA")
    overlay = Image.new("RGBA", base.size, (255, 255, 255, 0))

    alpha = int(round(max(0.0, min(float(opacity), 1.0)) * 255.0))
    if alpha <= 0:
        return img

    font = _load_font(size=max(10, int(font_size)), font_path=font_path)
    d = ImageDraw.Draw(overlay)

    try:
        box = d.textbbox((0, 0), text, font=font)
        tw = max(1, int(box[2] - box[0]))
        th = max(1, int(box[3] - box[1]))
    except Exception:
        tw, th = d.textsize(text, font=font)
        tw = max(1, int(tw))
        th = max(1, int(th))

    pad = 12
    text_img = Image.new("RGBA", (tw + 2 * pad, th + 2 * pad), (255, 255, 255, 0))
    d2 = ImageDraw.Draw(text_img)
    d2.text((pad, pad), text, font=font, fill=(0, 0, 0, alpha))

    rotated = text_img.rotate(float(angle), expand=True, resample=Image.Resampling.BICUBIC)
    step_x = max(120, rotated.width + 240)
    step_y = max(120, rotated.height + 240)

    for y in range(-rotated.height, base.height + rotated.height, step_y):
        for x in range(-rotated.width, base.width + rotated.width, step_x):
            overlay.alpha_composite(rotated, (x, y))

    out = Image.alpha_composite(base, overlay).convert("RGB")
    return out


def _materialize_cache_key_mode() -> str:
    # Default to sha256 for reproducible outputs across copies/machines.
    return (os.getenv("EIA_GEN_MATERIALIZE_CACHE_KEY_MODE", "sha256") or "").strip().lower()


def _parse_page_number(file_path: str, gen_method: str | None) -> int | None:
    """Return 1-based page number when specified, otherwise None."""
    if gen_method:
        m = _PAGE_RE.search(gen_method.strip())
        if m:
            try:
                return max(1, int(m.group(1)))
            except Exception:
                pass
    m = _FILE_PAGE_RE.search(file_path)
    if m:
        try:
            return max(1, int(m.group(1)))
        except Exception:
            return None
    return None


def _strip_page_fragment(file_path: str) -> str:
    # Remove "#page=.." or "?page=.." fragments for filesystem resolution
    return _FILE_PAGE_RE.sub("", file_path)


def select_pdf_page(
    pdf_path: Path,
    *,
    gen_method: str | None,
    target_dpi: int,
    max_pages: int = 15,
) -> tuple[int, str]:
    """Return (page_1based, source) for a PDF when rendering figures.

    - source="explicit": when page was specified via gen_method (PDF_PAGE/PAGE tokens) or file_path fragment.
    - source="heuristic": when page was not specified and we chose a likely drawing page.
    """
    page = _parse_page_number(pdf_path.name, gen_method)
    if page is not None:
        return page, "explicit"
    return _find_best_pdf_page(pdf_path, target_dpi=target_dpi, max_pages=max_pages), "heuristic"


def _parse_crop_box(crop: str | None, *, w: int, h: int) -> tuple[int, int, int, int] | None:
    """Parse crop box.

    Supported:
    - "l,t,r,b" (either pixels or 0~1 ratios)
    - "left=...,top=...,right=...,bottom=..." (same value rules)
    """
    if not crop:
        return None
    s = str(crop).strip()
    if not s:
        return None
    # Reserved keyword: AUTO crop is handled separately.
    if _AUTO_CROP_RE.match(s):
        return None

    # key=value form
    if "=" in s:
        parts = {}
        for token in re.split(r"[;\s]+", s):
            if not token or "=" not in token:
                continue
            k, v = token.split("=", 1)
            parts[k.strip().lower()] = v.strip()
        vals = [parts.get("left"), parts.get("top"), parts.get("right"), parts.get("bottom")]
        if any(v is None for v in vals):
            return None
        raw = vals
    else:
        raw = [p.strip() for p in s.split(",") if p.strip()]
        if len(raw) != 4:
            return None

    nums: list[float] = []
    for v in raw:
        try:
            nums.append(float(v))
        except Exception:
            return None

    # Heuristic: if all are between 0~1 -> ratio crop
    if all(0.0 <= x <= 1.0 for x in nums):
        l = int(round(nums[0] * w))
        t = int(round(nums[1] * h))
        r = int(round(nums[2] * w))
        b = int(round(nums[3] * h))
    else:
        l = int(round(nums[0]))
        t = int(round(nums[1]))
        r = int(round(nums[2]))
        b = int(round(nums[3]))

    # clamp
    l = max(0, min(l, w - 1))
    t = max(0, min(t, h - 1))
    r = max(l + 1, min(r, w))
    b = max(t + 1, min(b, h))
    return (l, t, r, b)


def _parse_auto_crop_params(crop: str) -> dict[str, float]:
    """Parse AUTO crop parameters.

    Supported:
    - "AUTO"
    - "AUTO:delta=10;pad=0.02;min_area=0.005"

    Params:
    - delta: 0~80. How far from white is considered "content". Higher means tighter crop.
    - pad: 0~0.2 (ratio of min(w,h)) extra padding around detected bbox.
    - min_area: 0~1. Minimum bbox area fraction to accept (avoid cropping on noise).
    """
    m = _AUTO_CROP_RE.match(str(crop).strip())
    if not m:
        return {}
    tail = (m.group(1) or "").strip()
    parts: dict[str, str] = {}
    if tail:
        for token in re.split(r"[;\s]+", tail):
            if not token or "=" not in token:
                continue
            k, v = token.split("=", 1)
            parts[k.strip().lower()] = v.strip()

    def _f(key: str, default: float) -> float:
        try:
            return float(parts.get(key, default))
        except Exception:
            return default

    delta = max(0.0, min(_f("delta", 10.0), 80.0))
    pad = max(0.0, min(_f("pad", 0.02), 0.2))
    min_area = max(0.0, min(_f("min_area", 0.005), 1.0))
    return {"delta": delta, "pad": pad, "min_area": min_area}


def _auto_crop_box(
    img: Image.Image, *, delta: float = 10.0, pad: float = 0.02, min_area: float = 0.005
) -> tuple[int, int, int, int] | None:
    """Auto-detect content bbox and return a crop box.

    This is mainly to handle CAD/plan PDFs where the drawing is centered and the rest is whitespace.
    The algorithm is intentionally simple and deterministic:
    - Convert to grayscale
    - Treat pixels darker than (255 - delta) as content
    - Compute bbox and pad it by `pad` ratio
    """
    if img.width <= 2 or img.height <= 2:
        return None

    # Heuristic: content is any pixel not close to white.
    try:
        d = int(round(float(delta)))
    except Exception:
        d = 10
    d = max(0, min(d, 80))
    threshold = 255 - d

    gray = ImageOps.grayscale(img)
    mask = gray.point(lambda x: 255 if x < threshold else 0, mode="L")
    bbox = mask.getbbox()
    if not bbox:
        return None

    l, t, r, b = bbox
    if r <= l + 1 or b <= t + 1:
        return None

    area_frac = ((r - l) * (b - t)) / float(img.width * img.height)
    if area_frac < float(min_area or 0.0):
        return None

    try:
        pad_ratio = float(pad)
    except Exception:
        pad_ratio = 0.02
    pad_ratio = max(0.0, min(pad_ratio, 0.2))
    pad_px = int(round(min(img.width, img.height) * pad_ratio))

    l = max(0, l - pad_px)
    t = max(0, t - pad_px)
    r = min(img.width, r + pad_px)
    b = min(img.height, b + pad_px)

    # Clamp
    l = max(0, min(l, img.width - 2))
    t = max(0, min(t, img.height - 2))
    r = max(l + 1, min(r, img.width))
    b = max(t + 1, min(b, img.height))
    return (l, t, r, b)


def _normalize_rgb_triplet(value: Any, *, default: tuple[int, int, int]) -> tuple[int, int, int]:
    if isinstance(value, (list, tuple)) and len(value) >= 3:
        try:
            r = int(value[0])
            g = int(value[1])
            b = int(value[2])
            return (max(0, min(r, 255)), max(0, min(g, 255)), max(0, min(b, 255)))
        except Exception:
            return default
    return default


def _frame_params(cfg: dict[str, Any], *, asset_type: str | None) -> tuple[bool, int, int, tuple[int, int, int], str]:
    """Return (apply_frame, pad_px, border_px, border_rgb, signature)."""
    at = (asset_type or "").strip()

    # Default: apply a submission-style frame to most assets, but avoid double-framing photo sheets
    # (they already have deterministic panel borders/caption bars).
    apply_frame = at != "photo_sheet"

    try:
        pad_px = int(float(_cfg_get(cfg, ["defaults", "common", "pad_px"], 0) or 0))
    except Exception:
        pad_px = 0
    try:
        border_px = int(float(_cfg_get(cfg, ["defaults", "common", "border_px"], 0) or 0))
    except Exception:
        border_px = 0

    pad_px = max(0, min(pad_px, 200))
    border_px = max(0, min(border_px, 50))

    border_rgb = _normalize_rgb_triplet(
        _cfg_get(cfg, ["defaults", "common", "border_rgb"], None), default=(0, 0, 0)
    )

    if pad_px == 0 and border_px == 0:
        apply_frame = False

    sig = (
        f"frame:apply={1 if apply_frame else 0}|pad={pad_px}|border={border_px}"
        f"|rgb={border_rgb[0]},{border_rgb[1]},{border_rgb[2]}"
    )
    return apply_frame, pad_px, border_px, border_rgb, sig


def _ensure_rgb(img: Image.Image) -> Image.Image:
    if img.mode == "RGB":
        return img
    if img.mode == "RGBA":
        # Flatten onto white background for DOCX safety.
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[-1])
        return bg
    return img.convert("RGB")


def _apply_frame(
    img: Image.Image, *, pad_px: int, border_px: int, border_rgb: tuple[int, int, int]
) -> Image.Image:
    img = _ensure_rgb(img)
    if border_px > 0:
        img = ImageOps.expand(img, border=int(border_px), fill=border_rgb)
    if pad_px > 0:
        img = ImageOps.expand(img, border=int(pad_px), fill=(255, 255, 255))
    return img


def _resize_to_target(
    img: Image.Image, *, width_mm: float | None, target_dpi: int, max_width_px: int, reserved_px: int = 0
) -> Image.Image:
    if img.width <= 0:
        return img

    reserved_px = max(0, int(reserved_px or 0))

    if width_mm and width_mm > 0:
        # Target pixel width from intended printed width (mm) at target_dpi.
        # Keep the overall output bounded; reserved_px accounts for frame pixels (pad+border).
        target_w_final = int(round((width_mm / 25.4) * float(target_dpi)))
        target_w_final = max(800, min(target_w_final, max_width_px))
    else:
        # No explicit width: keep the overall output <= max_width_px, including any frame.
        target_w_final = min(img.width + reserved_px, max_width_px)

    # Content area should fit inside the reserved frame.
    target_w = max(1, int(target_w_final) - reserved_px)

    if img.width <= target_w:
        return img
    scale = target_w / float(img.width)
    target_h = int(round(img.height * scale))
    return img.resize((target_w, max(1, target_h)), resample=Image.Resampling.LANCZOS)


def _open_image_from_pdf(pdf_path: Path, *, page_1based: int, dpi: int) -> Image.Image:
    import fitz  # PyMuPDF

    doc = fitz.open(str(pdf_path))
    try:
        idx = max(0, min(page_1based - 1, doc.page_count - 1))
        page = doc.load_page(idx)
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        return img
    finally:
        doc.close()


def _iter_pdf_pages(pdf_path: Path, *, max_pages: int | None = None) -> list[int]:
    """Return 1-based page numbers for a PDF (best-effort)."""
    import fitz  # PyMuPDF

    doc = fitz.open(str(pdf_path))
    try:
        total = int(doc.page_count)
    finally:
        doc.close()
    if total <= 0:
        return []
    pages = list(range(1, total + 1))
    if max_pages and max_pages > 0:
        pages = pages[: int(max_pages)]
    return pages


def _find_best_pdf_page(pdf_path: Path, *, target_dpi: int, max_pages: int = 15) -> int:
    """Pick a likely drawing page when no page was specified.

    Many plan PDFs start with a cover page (page 1). When FIGURES points to the PDF
    without an explicit page, choosing page 1 often yields the cover instead of the drawing.

    Heuristic (deterministic):
    - Rasterize first N pages at low DPI (fast)
    - Compute content bbox via `_auto_crop_box` (no padding)
    - Select the page with the largest bbox area fraction
    """
    pages = _iter_pdf_pages(pdf_path, max_pages=max_pages)
    if not pages:
        return 1

    best_page = 1
    best_score = -1.0
    dpi = int(target_dpi) if target_dpi and int(target_dpi) > 0 else 120
    dpi = min(120, dpi)

    for p in pages:
        try:
            img = _open_image_from_pdf(pdf_path, page_1based=p, dpi=dpi)
        except Exception:
            continue
        box = _auto_crop_box(img, delta=10.0, pad=0.0, min_area=0.0)
        if not box:
            score = 0.0
        else:
            l, t, r, b = box
            score = ((r - l) * (b - t)) / float(max(1, img.width * img.height))

        if score > best_score:
            best_score = score
            best_page = p

    return best_page


def materialize_figure_image(src_path: Path, opts: MaterializeOptions) -> Path:
    return materialize_figure_image_result(src_path, opts, include_meta=False).path


def materialize_figure_image_result(
    src_path: Path, opts: MaterializeOptions, *, include_meta: bool = True
) -> MaterializeResult:
    """Return a raster path (PNG/JPEG) that is safe to insert into DOCX.

    - If input is PDF: rasterize (page selectable).
    - If crop is present: crop (ratio or px).
    - If width_mm is present: resize toward target_dpi for size control.
    """
    opts.out_dir.mkdir(parents=True, exist_ok=True)

    # watermark signature (impacts output; must be part of the cache key)
    ref_params, ref_sig = _reference_watermark_params(opts.gen_method)
    ssot_params, ssot_sig = _ssot_sample_watermark_params(opts.gen_method)
    wm_sig = "|".join([s for s in (ref_sig, ssot_sig) if s])

    cfg_path = _default_asset_normalization_path()
    cfg = _load_asset_normalization_cfg(str(cfg_path.resolve()))

    is_pdf = src_path.suffix.lower() == ".pdf"
    defaults = _materialize_defaults(cfg, asset_type=opts.asset_type, is_pdf=is_pdf)

    # Final params (SSOT defaults → opts → gen_method override)
    target_dpi = int(opts.target_dpi) if (opts.target_dpi and int(opts.target_dpi) > 0) else int(defaults["target_dpi"])
    max_width_px = (
        int(opts.max_width_px) if (opts.max_width_px and int(opts.max_width_px) > 0) else int(defaults["max_width_px"])
    )

    gm_dpi = _parse_int_from_gen_method(opts.gen_method, pat=_GEN_TARGET_DPI_RE)
    if gm_dpi is not None:
        target_dpi = max(72, min(int(gm_dpi), 600))
    gm_maxw = _parse_int_from_gen_method(opts.gen_method, pat=_GEN_MAX_WIDTH_RE)
    if gm_maxw is not None:
        max_width_px = max(800, min(int(gm_maxw), 12000))

    fmt_override = _parse_output_format_from_gen_method(opts.gen_method)
    output_format = (
        fmt_override
        or _normalize_output_format(opts.output_format)
        or _normalize_output_format(defaults.get("output_format"))
        or "PNG"
    )
    # PDFs default to PNG unless user explicitly overrides.
    if is_pdf and output_format == "JPEG" and fmt_override is None and _normalize_output_format(opts.output_format) is None:
        output_format = "PNG"

    jpeg_quality = int(opts.jpeg_quality) if opts.jpeg_quality is not None else int(defaults.get("jpeg_quality") or 85)
    gm_q = _parse_int_from_gen_method(opts.gen_method, pat=_GEN_JPEG_QUALITY_RE)
    if gm_q is not None:
        jpeg_quality = int(gm_q)
    jpeg_quality = max(30, min(int(jpeg_quality), 95))

    ext = "png" if output_format == "PNG" else "jpg"

    apply_frame, pad_px, border_px, border_rgb, frame_sig = _frame_params(cfg, asset_type=opts.asset_type)
    reserved_px = (2 * (pad_px + border_px)) if apply_frame else 0

    # cache key
    stat = src_path.stat()
    mode = _materialize_cache_key_mode()
    if mode.startswith("mtime"):
        source_key = "|".join([str(src_path.resolve()), str(stat.st_mtime_ns)])
    else:
        # sha256 avoids cache misses when files are copied/moved and only mtime differs.
        source_key = _sha256_file_cached(src_path)

    key = _sha1_text(
        "|".join(
            [
                str(source_key),
                str(opts.gen_method or ""),
                str(wm_sig),
                str(frame_sig),
                str(opts.crop or ""),
                str(opts.width_mm or ""),
                str(target_dpi),
                str(max_width_px),
                str(output_format),
                str(jpeg_quality if output_format == "JPEG" else ""),
            ]
        )
    )[:10]
    out_path = opts.out_dir / f"{opts.fig_id}_{key}.{ext}"

    def _build_meta(*, pdf_page: int | None, pdf_page_source: str | None) -> dict[str, Any]:
        # Portable traceability: keep basenames + sha256; avoid absolute paths.
        out_sha = _sha256_file_cached(out_path) if out_path.exists() else ""
        recipe: dict[str, Any] = {
            "gen_method": (opts.gen_method or "").strip(),
            "crop": (opts.crop or "").strip(),
            "width_mm": float(opts.width_mm) if opts.width_mm is not None else None,
            "target_dpi": int(target_dpi),
            "max_width_px": int(max_width_px),
            "output_format": str(output_format),
            "jpeg_quality": int(jpeg_quality) if output_format == "JPEG" else None,
            "apply_frame": bool(apply_frame),
            "pad_px": int(pad_px) if apply_frame else 0,
            "border_px": int(border_px) if apply_frame else 0,
            "border_rgb": [int(x) for x in (border_rgb or [0, 0, 0])] if apply_frame else [0, 0, 0],
            "frame_sig": str(frame_sig or ""),
            "watermark_sig": wm_sig,
            "reserved_px": int(reserved_px),
        }
        input_manifest: list[dict[str, Any]] = [
            {
                "name": src_path.name,
                "sha256": _sha256_file_cached(src_path),
                "mtime_ns": int(stat.st_mtime_ns),
                "size": int(stat.st_size),
            }
        ]
        meta: dict[str, Any] = {
            "kind": "MATERIALIZE",
            "recipe": recipe,
            "input_manifest": input_manifest,
            "cache_key_mode": str(mode),
            "cache_key": str(key),
            "src_name": src_path.name,
            "src_sha256": _sha256_file_cached(src_path),
            "src_mtime_ns": int(stat.st_mtime_ns),
            "src_size": int(stat.st_size),
            "out_name": out_path.name,
            "out_sha256": out_sha,
            "asset_type": (opts.asset_type or "").strip(),
            "is_pdf": bool(is_pdf),
            "pdf_page": int(pdf_page) if pdf_page is not None else None,
            "pdf_page_source": str(pdf_page_source or ""),
            # Flat keys kept for backwards compatibility with existing exports/tests.
            "gen_method": str(recipe.get("gen_method") or ""),
            "crop": str(recipe.get("crop") or ""),
            "width_mm": recipe.get("width_mm"),
            "target_dpi": recipe.get("target_dpi"),
            "max_width_px": recipe.get("max_width_px"),
            "output_format": str(recipe.get("output_format") or ""),
            "jpeg_quality": recipe.get("jpeg_quality"),
            "apply_frame": recipe.get("apply_frame"),
            "pad_px": recipe.get("pad_px"),
            "border_px": recipe.get("border_px"),
            "border_rgb": recipe.get("border_rgb"),
            "frame_sig": str(recipe.get("frame_sig") or ""),
            "reserved_px": recipe.get("reserved_px"),
            "watermark_sig": str(recipe.get("watermark_sig") or ""),
        }
        return meta

    if out_path.exists():
        if not include_meta:
            return MaterializeResult(path=out_path)
        if is_pdf:
            page, page_source = select_pdf_page(
                src_path, gen_method=opts.gen_method, target_dpi=target_dpi, max_pages=15
            )
            return MaterializeResult(
                path=out_path,
                pdf_page=page,
                pdf_page_source=page_source,
                meta=_build_meta(pdf_page=page, pdf_page_source=page_source),
            )
        return MaterializeResult(path=out_path, meta=_build_meta(pdf_page=None, pdf_page_source=None))

    # load
    if is_pdf:
        page, page_source = select_pdf_page(
            src_path, gen_method=opts.gen_method, target_dpi=target_dpi, max_pages=15
        )
        img = _open_image_from_pdf(src_path, page_1based=page, dpi=target_dpi)
    else:
        page = None
        page_source = None
        img = Image.open(src_path)
        # normalize orientation (EXIF)
        img = ImageOps.exif_transpose(img)
        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGB")

    # crop
    box = None
    if opts.crop and _AUTO_CROP_RE.match(str(opts.crop).strip()):
        params = _parse_auto_crop_params(str(opts.crop))
        box = _auto_crop_box(img, **params)
    else:
        box = _parse_crop_box(opts.crop, w=img.width, h=img.height)
    if box:
        img = img.crop(box)

    # resize to control docx size (reserve pixels for the submission-style frame)
    img = _resize_to_target(
        img,
        width_mm=opts.width_mm,
        target_dpi=target_dpi,
        max_width_px=max_width_px,
        reserved_px=reserved_px,
    )

    # Apply watermarks deterministically (order matters when multiple are present).
    for wm in (ssot_params, ref_params):
        if wm is not None:
            img = _apply_text_watermark(img, **wm)

    if apply_frame:
        img = _apply_frame(img, pad_px=pad_px, border_px=border_px, border_rgb=border_rgb)

    # save
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if output_format == "JPEG":
        img_rgb = _ensure_rgb(img)
        img_rgb.save(out_path, format="JPEG", quality=jpeg_quality, optimize=True, progressive=False)
    else:
        # PNG is safe for diagrams; for photos this can be large but predictable.
        img_rgb = _ensure_rgb(img)
        img_rgb.save(out_path, format="PNG", optimize=True)
    if not include_meta:
        return MaterializeResult(path=out_path, pdf_page=page, pdf_page_source=page_source)
    return MaterializeResult(
        path=out_path,
        pdf_page=page,
        pdf_page_source=page_source,
        meta=_build_meta(pdf_page=page, pdf_page_source=page_source),
    )


def resolve_source_path(raw_file_path: str, *, asset_search_dirs: list[Path] | None) -> Path | None:
    """Resolve file paths that may include '#page=N' fragments."""
    fp = (raw_file_path or "").strip()
    if not fp:
        return None
    fp2 = _strip_page_fragment(fp)
    p = Path(fp2).expanduser()
    if p.is_absolute():
        return p if p.exists() else None
    if asset_search_dirs:
        for base in asset_search_dirs:
            cand = (base / p).expanduser()
            if cand.exists():
                return cand
    return p if p.exists() else None
