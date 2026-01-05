#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF
import pytesseract
from PIL import Image


CHAPTER_RE = re.compile(r"제\s*(\d+)\s*장")
# Allow labels like 7.6-3 or 7.3.1-3 (common in Korean EIA/DIA reports)
_NUM_RE = r"[0-9]+(?:\.[0-9]+)*(?:-[0-9]+)?"
FIG_RE = re.compile(r"<\s*그림\s*(%s)\s*[^>]*>\s*(.+)" % _NUM_RE)
TBL_RE = re.compile(r"<\s*표\s*(%s)\s*[^>]*>\s*(.+)" % _NUM_RE)


@dataclass
class Hit:
    kind: str  # chapter|figure|table
    page: int
    label: str
    text: str


def ocr_page(doc: fitz.Document, page_index: int, zoom: float, lang: str) -> str:
    page = doc.load_page(page_index)
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return pytesseract.image_to_string(img, lang=lang)


def extract_index(pdf_path: Path, max_pages: int | None, zoom: float, lang: str) -> dict[str, Any]:
    doc = fitz.open(str(pdf_path))
    page_count = doc.page_count
    limit = min(page_count, max_pages) if max_pages else page_count

    hits: list[Hit] = []

    for i in range(limit):
        txt = ocr_page(doc, i, zoom=zoom, lang=lang)
        # Normalize lines a bit for matching
        lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
        for ln in lines:
            m = CHAPTER_RE.search(ln)
            if m:
                hits.append(Hit(kind="chapter", page=i + 1, label=f"CH{m.group(1)}", text=ln))
            m = FIG_RE.search(ln)
            if m:
                hits.append(Hit(kind="figure", page=i + 1, label=m.group(1), text=m.group(2).strip()))
            m = TBL_RE.search(ln)
            if m:
                hits.append(Hit(kind="table", page=i + 1, label=m.group(1), text=m.group(2).strip()))

    # De-dup (same OCR line often repeats across pages in this PDF)
    seen: set[tuple[str, int, str, str]] = set()
    out_hits: list[dict[str, Any]] = []
    for h in hits:
        key = (h.kind, h.page, h.label, h.text)
        if key in seen:
            continue
        seen.add(key)
        out_hits.append({"kind": h.kind, "page": h.page, "label": h.label, "text": h.text})

    return {
        "pdf_path": str(pdf_path),
        "page_count": page_count,
        "limit_pages": limit,
        "ocr": {"zoom": zoom, "lang": lang},
        "hits": out_hits,
    }


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True, type=Path)
    ap.add_argument("--out", required=True, type=Path)
    ap.add_argument("--max-pages", type=int, default=None)
    ap.add_argument("--zoom", type=float, default=1.0)
    ap.add_argument("--lang", default="kor+eng")
    args = ap.parse_args()

    data = extract_index(args.pdf, args.max_pages, args.zoom, args.lang)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK wrote {args.out} ({len(data['hits'])} hits)")


if __name__ == "__main__":
    main()

