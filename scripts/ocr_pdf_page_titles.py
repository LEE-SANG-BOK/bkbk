#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import fitz  # PyMuPDF
import pytesseract
from PIL import Image, ImageEnhance, ImageOps


_WS_RE = re.compile(r"\s+")
_KEEP_RE = re.compile(r"[^0-9A-Za-z가-힣]+")


@dataclass(frozen=True)
class PageOcr:
    page: int  # 1-based
    text: str

    @property
    def text_one_line(self) -> str:
        return _WS_RE.sub(" ", self.text).strip()

    @property
    def compact(self) -> str:
        """Whitespace/punct removed for robust keyword matching."""
        s = _KEEP_RE.sub("", self.text_one_line)
        return s.strip().lower()


def _render_page(doc: fitz.Document, page_1based: int, *, dpi: int) -> Image.Image:
    page = doc.load_page(page_1based - 1)
    pix = page.get_pixmap(dpi=int(dpi), alpha=False)
    return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)


def _crop_top(img: Image.Image, ratio: float) -> Image.Image:
    try:
        r = float(ratio)
    except Exception:
        return img
    if r <= 0.0:
        return img
    if r >= 1.0:
        return img
    w, h = img.size
    y1 = int(round(h * r))
    y1 = max(1, min(y1, h))
    return img.crop((0, 0, w, y1))


def _preprocess_for_ocr(img: Image.Image, *, contrast: float, threshold: int | None) -> Image.Image:
    out = ImageOps.grayscale(img)
    if contrast and abs(float(contrast) - 1.0) > 1e-6:
        out = ImageEnhance.Contrast(out).enhance(float(contrast))
    if threshold is not None:
        t = int(max(0, min(int(threshold), 255)))
        out = out.point(lambda x: 0 if x < t else 255, mode="L")
    return out


def _ocr_image(img: Image.Image, *, lang: str, psm: int) -> str:
    cfg = f"--psm {int(psm)}" if psm else ""
    return pytesseract.image_to_string(img, lang=lang, config=cfg)


def _normalize_keywords(keywords: Iterable[str]) -> list[str]:
    out: list[str] = []
    for kw in keywords:
        kw = str(kw).strip()
        if not kw:
            continue
        out.append(_KEEP_RE.sub("", _WS_RE.sub(" ", kw)).strip().lower())
    return [x for x in out if x]


def _select_pages(doc: fitz.Document, *, page_start: int | None, page_end: int | None) -> list[int]:
    start = max(1, int(page_start or 1))
    end = int(page_end or doc.page_count)
    end = min(end, doc.page_count)
    if end < start:
        return []
    return list(range(start, end + 1))


def main() -> None:
    ap = argparse.ArgumentParser(
        description=(
            "OCR the TOP area of a scanned PDF to quickly find page titles "
            "(e.g., 위치도/계획평면도/배수계획평면도). Useful for SSOT_PAGE_OVERRIDES."
        )
    )
    ap.add_argument("--pdf", required=True, type=Path)
    ap.add_argument("--out", type=Path, default=None, help="Optional JSON output path.")
    ap.add_argument("--lang", default="kor+eng")
    ap.add_argument("--dpi", type=int, default=160, help="Rasterize DPI (higher=slower, better OCR).")

    ap.add_argument("--page-start", type=int, default=None, help="1-based start page (inclusive)")
    ap.add_argument("--page-end", type=int, default=None, help="1-based end page (inclusive)")

    ap.add_argument(
        "--crop-top-ratio",
        type=float,
        default=0.35,
        help="OCR only top area ratio (0~1). Drawings often have titles at top.",
    )
    ap.add_argument("--contrast", type=float, default=2.0, help="OCR contrast boost (default 2.0).")
    ap.add_argument("--threshold", type=int, default=180, help="OCR binarization threshold (0~255).")
    ap.add_argument("--psm", type=int, default=6, help="Tesseract page segmentation mode (default 6).")

    ap.add_argument(
        "--keywords",
        default="",
        help="Optional keywords (comma/semicolon separated). Matching uses 'compact' text (spaces/punct removed).",
    )
    ap.add_argument("--max-print", type=int, default=200, help="Max printed rows.")

    args = ap.parse_args()

    if not args.pdf.exists():
        raise SystemExit(f"PDF not found: {args.pdf}")

    doc = fitz.open(str(args.pdf))
    pages = _select_pages(doc, page_start=args.page_start, page_end=args.page_end)

    raw_keywords = re.split(r"[;,]\s*", str(args.keywords or ""))
    keywords = _normalize_keywords(raw_keywords)

    rows: list[dict[str, Any]] = []
    printed = 0

    for p in pages:
        img = _render_page(doc, p, dpi=int(args.dpi))
        img = _crop_top(img, float(args.crop_top_ratio))
        img = _preprocess_for_ocr(img, contrast=float(args.contrast or 1.0), threshold=args.threshold)
        txt = _ocr_image(img, lang=args.lang, psm=int(args.psm or 6))
        po = PageOcr(page=p, text=txt)

        match = True
        if keywords:
            match = any((kw in po.compact) for kw in keywords)

        if match:
            row = {"page": p, "text": po.text_one_line, "compact": po.compact}
            rows.append(row)
            if printed < int(args.max_print):
                snip = row["text"][:120]
                print(f"{p:>4}: {snip}")
                printed += 1

    out_obj = {
        "pdf_path": str(args.pdf),
        "page_count": doc.page_count,
        "scanned_pages": {"start": pages[0] if pages else None, "end": pages[-1] if pages else None, "count": len(pages)},
        "settings": {
            "lang": args.lang,
            "dpi": int(args.dpi),
            "crop_top_ratio": float(args.crop_top_ratio),
            "contrast": float(args.contrast),
            "threshold": args.threshold,
            "psm": int(args.psm),
            "keywords": keywords,
        },
        "matches": rows,
    }

    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(json.dumps(out_obj, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"OK wrote {args.out}")


if __name__ == "__main__":
    main()

