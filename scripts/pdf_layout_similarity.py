#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from PIL import Image, ImageChops


def _require_cmd(cmd: str) -> str:
    p = shutil.which(cmd)
    if not p:
        raise SystemExit(f"Missing command in PATH: {cmd}")
    return p


def _pdf_page_count(pdf: Path) -> int:
    pdfinfo = _require_cmd("pdfinfo")
    out = subprocess.check_output([pdfinfo, str(pdf)], text=True, stderr=subprocess.STDOUT)
    for line in out.splitlines():
        if line.lower().startswith("pages:"):
            try:
                return int(line.split(":", 1)[1].strip())
            except Exception as e:
                raise SystemExit(f"Failed to parse page count from pdfinfo output: {line!r}") from e
    raise SystemExit(f"pdfinfo output missing 'Pages:' line: {pdf}")


def _render_pdf_page(pdf: Path, *, page: int, dpi: int, out_png: Path) -> None:
    pdftoppm = _require_cmd("pdftoppm")
    out_png.parent.mkdir(parents=True, exist_ok=True)
    prefix = out_png.with_suffix("")  # pdftoppm writes <prefix>.png
    cmd = [
        pdftoppm,
        "-f",
        str(int(page)),
        "-l",
        str(int(page)),
        "-r",
        str(int(dpi)),
        "-png",
        "-singlefile",
        str(pdf),
        str(prefix),
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    if not out_png.exists():
        raise SystemExit(f"pdftoppm did not produce expected PNG: {out_png}")


def _pad_to_size(im: Image.Image, *, size: tuple[int, int]) -> Image.Image:
    if im.size == size:
        return im
    bg = Image.new("L", size, color=255)
    bg.paste(im, (0, 0))
    return bg


def _mean_abs_diff_0_255(diff_l: Image.Image) -> float:
    hist = diff_l.histogram()
    total = diff_l.size[0] * diff_l.size[1]
    if total <= 0:
        return 255.0
    s = 0
    for value, count in enumerate(hist):
        s += int(value) * int(count)
    return float(s) / float(total)


def _similarity_pct(a_png: Path, b_png: Path) -> float:
    with Image.open(a_png) as a:
        a_l = a.convert("L")
        with Image.open(b_png) as b:
            b_l = b.convert("L")

    w = max(a_l.size[0], b_l.size[0])
    h = max(a_l.size[1], b_l.size[1])
    size = (w, h)
    a_l = _pad_to_size(a_l, size=size)
    b_l = _pad_to_size(b_l, size=size)

    diff = ImageChops.difference(a_l, b_l)
    mad = _mean_abs_diff_0_255(diff)
    sim = max(0.0, min(1.0, 1.0 - (mad / 255.0)))
    return round(sim * 100.0, 2)


def _parse_pages(spec: str, *, max_page: int) -> list[int]:
    s = (spec or "").strip().lower()
    if not s or s in {"auto", "all"}:
        return list(range(1, max_page + 1))

    pages: set[int] = set()
    for part in (p.strip() for p in s.split(",") if p.strip()):
        if "-" in part:
            a, b = (x.strip() for x in part.split("-", 1))
            if not a or not b:
                raise SystemExit(f"Invalid pages range: {part!r}")
            start = int(a)
            end = int(b)
            if start > end:
                start, end = end, start
            for p in range(start, end + 1):
                if 1 <= p <= max_page:
                    pages.add(p)
            continue
        p = int(part)
        if 1 <= p <= max_page:
            pages.add(p)
    return sorted(pages)


def _convert_docx_to_pdf(docx: Path, *, out_pdf: Path, mode: str, update_fields: bool) -> None:
    here = Path(__file__).resolve()
    post = here.parent / "postprocess_docx_to_pdf.py"
    if not post.exists():
        raise SystemExit(f"Missing converter script: {post}")
    cmd = [
        sys.executable,
        str(post),
        "--docx",
        str(docx),
        "--out-pdf",
        str(out_pdf),
        "--mode",
        str(mode),
    ]
    if not update_fields:
        cmd.append("--no-update-fields")
    subprocess.run(cmd, check=True)
    if not out_pdf.exists():
        raise SystemExit(f"DOCX->PDF conversion failed (no output): {out_pdf}")


@dataclass(frozen=True)
class PageScore:
    page: int
    similarity_pct: float


def main() -> None:
    ap = argparse.ArgumentParser(
        description=(
            "Compute a raster-based page/layout similarity score (%) between two PDFs/DOCXs.\n"
            "- Renders both inputs via pdftoppm at a fixed DPI.\n"
            "- Computes mean-absolute pixel difference (0..255) and reports 100*(1 - MAD/255).\n"
            "Notes: This is a best-effort metric and is sensitive to fonts/renderers/DPI."
        )
    )
    ap.add_argument("--a", type=Path, required=True, help="Input A (.pdf or .docx)")
    ap.add_argument("--b", type=Path, required=True, help="Input B (.pdf or .docx)")
    ap.add_argument("--pages", default="auto", help="Pages to compare (e.g., '1-10,15', default 'auto'=all)")
    ap.add_argument("--dpi", type=int, default=120, help="Render DPI (default 120)")
    ap.add_argument(
        "--convert-mode",
        choices=["auto", "word", "pages", "soffice"],
        default="auto",
        help="When converting DOCX, use this backend (default auto)",
    )
    ap.add_argument(
        "--no-update-fields",
        action="store_true",
        help="When converting DOCX via Word backend, skip field/TOC updates",
    )
    ap.add_argument("--out-json", type=Path, default=None, help="Optional JSON report output path")
    ap.add_argument(
        "--keep-dir",
        type=Path,
        default=None,
        help="If set, keep rendered PNGs under this dir (otherwise uses a temp dir).",
    )
    args = ap.parse_args()

    a = args.a.expanduser().resolve()
    b = args.b.expanduser().resolve()
    if not a.exists():
        raise SystemExit(f"Input not found: {a}")
    if not b.exists():
        raise SystemExit(f"Input not found: {b}")

    if args.keep_dir:
        work_dir = args.keep_dir.expanduser().resolve()
        work_dir.mkdir(parents=True, exist_ok=True)
        tmp: Optional[tempfile.TemporaryDirectory[str]] = None
    else:
        tmp = tempfile.TemporaryDirectory(prefix="pdf_layout_similarity_")
        work_dir = Path(tmp.name)

    try:
        pdf_a = a
        pdf_b = b

        if a.suffix.lower() == ".docx":
            pdf_a = work_dir / f"{a.stem}.pdf"
            _convert_docx_to_pdf(a, out_pdf=pdf_a, mode=args.convert_mode, update_fields=not args.no_update_fields)
        if b.suffix.lower() == ".docx":
            pdf_b = work_dir / f"{b.stem}.pdf"
            _convert_docx_to_pdf(b, out_pdf=pdf_b, mode=args.convert_mode, update_fields=not args.no_update_fields)

        if pdf_a.suffix.lower() != ".pdf":
            raise SystemExit(f"Unsupported input A: {a} (expected .pdf or .docx)")
        if pdf_b.suffix.lower() != ".pdf":
            raise SystemExit(f"Unsupported input B: {b} (expected .pdf or .docx)")

        pages_a = _pdf_page_count(pdf_a)
        pages_b = _pdf_page_count(pdf_b)
        max_common = min(pages_a, pages_b)
        if max_common <= 0:
            raise SystemExit("No pages to compare (page_count <= 0).")

        pages = _parse_pages(args.pages, max_page=max_common)
        if not pages:
            raise SystemExit(f"No valid pages to compare (max_common={max_common}, pages={args.pages!r}).")

        out_a = work_dir / "a"
        out_b = work_dir / "b"
        scores: list[PageScore] = []
        for p in pages:
            a_png = out_a / f"p-{p:04d}.png"
            b_png = out_b / f"p-{p:04d}.png"
            _render_pdf_page(pdf_a, page=p, dpi=args.dpi, out_png=a_png)
            _render_pdf_page(pdf_b, page=p, dpi=args.dpi, out_png=b_png)
            scores.append(PageScore(page=p, similarity_pct=_similarity_pct(a_png, b_png)))

        avg = round(sum(s.similarity_pct for s in scores) / float(len(scores)), 2)
        mn = min(s.similarity_pct for s in scores)
        mx = max(s.similarity_pct for s in scores)

        print(f"Compared pages: {len(scores)} (common_pages={max_common}, dpi={args.dpi})")
        print(f"Similarity: avg={avg:.2f}% min={mn:.2f}% max={mx:.2f}%")

        if args.out_json:
            out = args.out_json.expanduser().resolve()
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text(
                json.dumps(
                    {
                        "input_a": str(a),
                        "input_b": str(b),
                        "pdf_a": str(pdf_a),
                        "pdf_b": str(pdf_b),
                        "dpi": int(args.dpi),
                        "pages_compared": [s.page for s in scores],
                        "page_scores": [{"page": s.page, "similarity_pct": s.similarity_pct} for s in scores],
                        "summary": {"avg_pct": avg, "min_pct": mn, "max_pct": mx},
                        "note": "Raster-based similarity (100*(1 - mean_abs_diff/255)). Sensitive to fonts/renderers/DPI.",
                    },
                    ensure_ascii=False,
                    indent=2,
                ),
                encoding="utf-8",
            )
            print(f"OK wrote: {out}")
    finally:
        if tmp is not None:
            tmp.cleanup()


if __name__ == "__main__":
    main()
