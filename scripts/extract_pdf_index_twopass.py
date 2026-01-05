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


CHAPTER_RE = re.compile(r"제\s*(\d+)\s*장")
# 다양한 표기(띄어쓰기/괄호/하이픈) OCR 흔들림 대응
# - "7.6-3", "7.3.1-3" 같은 번호 패턴을 허용
_NUM_RE = r"[0-9]+(?:\.[0-9]+)*(?:-[0-9]+)?"
FIG_RE = re.compile(r"<\s*그림\s*(%s)\s*[^>]*>\s*(.+)" % _NUM_RE)
TBL_RE = re.compile(r"<\s*표\s*(%s)\s*[^>]*>\s*(.+)" % _NUM_RE)


@dataclass(frozen=True)
class Hit:
    kind: str  # chapter|figure|table
    page: int
    label: str
    text: str


def _render_page(doc: fitz.Document, page_1based: int, zoom: float) -> Image.Image:
    page = doc.load_page(page_1based - 1)
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)


def _crop_bottom(img: Image.Image, ratio: float) -> Image.Image:
    """Crop the bottom area of a page by ratio (0~1).

    Many scanned EIA/DIA PDFs place figure/table captions at the bottom.
    Cropping bottom improves OCR recall and reduces noise.
    """
    try:
        r = float(ratio)
    except Exception:
        return img
    if r <= 0.0:
        return img
    if r >= 1.0:
        return img
    w, h = img.size
    y0 = int(round(h * (1.0 - r)))
    y0 = max(0, min(y0, h - 1))
    return img.crop((0, y0, w, h))


def _preprocess_for_ocr(img: Image.Image, *, contrast: float = 1.0, threshold: int | None = None) -> Image.Image:
    out = ImageOps.grayscale(img)
    if contrast and abs(float(contrast) - 1.0) > 1e-6:
        out = ImageEnhance.Contrast(out).enhance(float(contrast))
    if threshold is not None:
        t = int(max(0, min(int(threshold), 255)))
        out = out.point(lambda x: 0 if x < t else 255, mode="L")
    return out


def _ocr_image(img: Image.Image, lang: str, *, psm: int = 6) -> str:
    cfg = f"--psm {int(psm)}" if psm else ""
    return pytesseract.image_to_string(img, lang=lang, config=cfg)


def _normalize_lines(txt: str) -> list[str]:
    return [ln.strip() for ln in txt.splitlines() if ln.strip()]


def _extract_hits_from_lines(page: int, lines: Iterable[str]) -> list[Hit]:
    hits: list[Hit] = []
    for ln in lines:
        m = CHAPTER_RE.search(ln)
        if m:
            hits.append(Hit(kind="chapter", page=page, label=f"CH{m.group(1)}", text=ln))

        m = FIG_RE.search(ln)
        if m:
            hits.append(Hit(kind="figure", page=page, label=m.group(1), text=m.group(2).strip()))

        m = TBL_RE.search(ln)
        if m:
            hits.append(Hit(kind="table", page=page, label=m.group(1), text=m.group(2).strip()))

    return hits


def _select_pages(doc: fitz.Document, *, page_start: int | None, page_end: int | None, max_pages: int | None) -> list[int]:
    start = max(1, int(page_start or 1))
    end = int(page_end or doc.page_count)
    end = min(end, doc.page_count)
    if end < start:
        return []

    pages = list(range(start, end + 1))
    if max_pages:
        pages = pages[: max(0, int(max_pages))]
    return pages


def pass1_candidates(
    doc: fitz.Document,
    *,
    pages_1based: list[int],
    zoom: float,
    lang: str,
    crop_bottom_ratio: float,
    contrast: float,
    threshold: int | None,
    psm: int,
) -> dict[str, Any]:
    """빠른 OCR로 후보 페이지를 고른다.

    - '그림', '표', '제..장' 같은 키워드가 한 줄이라도 잡히면 후보로 선정.
    - 스캔본 PDF는 캡션이 하단에 몰려있는 경우가 많아 `--crop-bottom-ratio`를 추천.
    """
    candidates: set[int] = set()
    page_snips: dict[int, list[str]] = {}

    for p in pages_1based:
        img = _render_page(doc, p, zoom=zoom)
        img = _crop_bottom(img, crop_bottom_ratio)
        img = _preprocess_for_ocr(img, contrast=contrast, threshold=threshold)
        txt = _ocr_image(img, lang=lang, psm=psm)
        lines = _normalize_lines(txt)

        # 후보 선정용 단순 힌트
        hint = [ln for ln in lines if ("그림" in ln or "표" in ln or ("제" in ln and "장" in ln))]
        if hint:
            candidates.add(p)
            page_snips[p] = hint[:20]

    return {
        "zoom": zoom,
        "lang": lang,
        "crop_bottom_ratio": crop_bottom_ratio,
        "contrast": contrast,
        "threshold": threshold,
        "psm": psm,
        "candidate_pages": sorted(candidates),
        "snips": page_snips,
    }


def pass2_extract(
    doc: fitz.Document,
    *,
    pages_1based: list[int],
    zoom: float,
    lang: str,
    crop_bottom_ratio: float,
    contrast: float,
    threshold: int | None,
    psm: int,
) -> dict[str, Any]:
    """후보 페이지만 고해상도로 재-OCR하고, 정규식 기반으로 chapter/figure/table을 추출한다."""
    hits: list[Hit] = []
    for p in pages_1based:
        img = _render_page(doc, p, zoom=zoom)
        img = _crop_bottom(img, crop_bottom_ratio)
        img = _preprocess_for_ocr(img, contrast=contrast, threshold=threshold)
        txt = _ocr_image(img, lang=lang, psm=psm)
        lines = _normalize_lines(txt)
        hits.extend(_extract_hits_from_lines(page=p, lines=lines))

    # de-dup
    seen: set[tuple[str, int, str, str]] = set()
    out_hits: list[dict[str, Any]] = []
    for h in hits:
        key = (h.kind, h.page, h.label, h.text)
        if key in seen:
            continue
        seen.add(key)
        out_hits.append({"kind": h.kind, "page": h.page, "label": h.label, "text": h.text})

    return {
        "zoom": zoom,
        "lang": lang,
        "crop_bottom_ratio": crop_bottom_ratio,
        "contrast": contrast,
        "threshold": threshold,
        "psm": psm,
        "hits": out_hits,
    }


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True, type=Path)
    ap.add_argument("--out-dir", required=True, type=Path)
    ap.add_argument("--lang", default="kor+eng")
    ap.add_argument("--pass1-zoom", type=float, default=0.9)
    ap.add_argument("--pass2-zoom", type=float, default=1.8)

    ap.add_argument("--page-start", type=int, default=None, help="1-based start page (inclusive)")
    ap.add_argument("--page-end", type=int, default=None, help="1-based end page (inclusive)")
    ap.add_argument("--max-pages", type=int, default=None, help="개발/디버그용: 선택된 범위에서 앞 n페이지")

    ap.add_argument(
        "--crop-bottom-ratio",
        type=float,
        default=0.0,
        help="OCR을 하단 영역(비율)로 제한 (예: 0.30=하단 30%). 스캔본 캡션 인식률에 도움.",
    )
    ap.add_argument("--contrast", type=float, default=1.0, help="OCR 전 대비(contrast) 증폭. (예: 2.0)")
    ap.add_argument("--threshold", type=int, default=None, help="OCR 전 이진화 임계값(0~255). (예: 180)")
    ap.add_argument("--psm", type=int, default=6, help="Tesseract page segmentation mode (기본 6)")

    args = ap.parse_args()

    args.out_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(args.pdf))

    pages = _select_pages(doc, page_start=args.page_start, page_end=args.page_end, max_pages=args.max_pages)

    p1 = pass1_candidates(
        doc,
        pages_1based=pages,
        zoom=args.pass1_zoom,
        lang=args.lang,
        crop_bottom_ratio=float(args.crop_bottom_ratio or 0.0),
        contrast=float(args.contrast or 1.0),
        threshold=args.threshold,
        psm=int(args.psm or 6),
    )
    p1_path = args.out_dir / "pass1_candidates.json"
    p1_path.write_text(json.dumps(p1, ensure_ascii=False, indent=2), encoding="utf-8")

    candidate_pages = p1["candidate_pages"]
    p2 = pass2_extract(
        doc,
        pages_1based=candidate_pages,
        zoom=args.pass2_zoom,
        lang=args.lang,
        crop_bottom_ratio=float(args.crop_bottom_ratio or 0.0),
        contrast=float(args.contrast or 1.0),
        threshold=args.threshold,
        psm=int(args.psm or 6),
    )
    p2_path = args.out_dir / "pass2_hits.json"
    p2_path.write_text(json.dumps(p2, ensure_ascii=False, indent=2), encoding="utf-8")

    combined = {
        "pdf_path": str(args.pdf),
        "page_count": doc.page_count,
        "scanned_pages": {
            "start": pages[0] if pages else None,
            "end": pages[-1] if pages else None,
            "count": len(pages),
        },
        "pass1": {
            "zoom": args.pass1_zoom,
            "lang": args.lang,
            "crop_bottom_ratio": float(args.crop_bottom_ratio or 0.0),
            "contrast": float(args.contrast or 1.0),
            "threshold": args.threshold,
            "psm": int(args.psm or 6),
            "candidate_pages": candidate_pages,
        },
        "pass2": {
            "zoom": args.pass2_zoom,
            "lang": args.lang,
            "crop_bottom_ratio": float(args.crop_bottom_ratio or 0.0),
            "contrast": float(args.contrast or 1.0),
            "threshold": args.threshold,
            "psm": int(args.psm or 6),
            "hits": p2["hits"],
        },
    }
    combined_path = args.out_dir / "combined_index.json"
    combined_path.write_text(json.dumps(combined, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"OK wrote:\n- {p1_path}\n- {p2_path}\n- {combined_path}")


if __name__ == "__main__":
    main()
