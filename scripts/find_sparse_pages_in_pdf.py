#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from PIL import Image


def _require_cmd(cmd: str) -> str:
    p = shutil.which(cmd)
    if not p:
        raise SystemExit(f"Missing command in PATH: {cmd}")
    return p


def _pdf_page_count(pdf: Path) -> Optional[int]:
    pdfinfo = shutil.which("pdfinfo")
    if not pdfinfo:
        return None
    try:
        out = subprocess.check_output([pdfinfo, str(pdf)], text=True, stderr=subprocess.STDOUT)
    except Exception:
        return None
    for line in out.splitlines():
        if line.lower().startswith("pages:"):
            try:
                return int(line.split(":", 1)[1].strip())
            except Exception:
                return None
    return None


def _render_range(pdf: Path, *, page_start: int, page_end: int, dpi: int, out_dir: Path) -> list[Path]:
    pdftoppm = _require_cmd("pdftoppm")
    out_dir.mkdir(parents=True, exist_ok=True)
    prefix = out_dir / "p"
    cmd = [
        pdftoppm,
        "-f",
        str(int(page_start)),
        "-l",
        str(int(page_end)),
        "-r",
        str(int(dpi)),
        "-png",
        str(pdf),
        str(prefix),
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return sorted(out_dir.glob("p-*.png"))


def _parse_page_from_filename(path: Path) -> int:
    # "p-001.png" or "p-320.png" -> 1 / 320
    stem = path.stem
    try:
        return int(stem.split("-")[-1])
    except Exception as e:
        raise RuntimeError(f"Unexpected pdftoppm output filename: {path.name}") from e


def _nonwhite_ratio(path: Path, *, nonwhite_threshold: int) -> float:
    with Image.open(path) as im:
        im = im.convert("L")
        hist = im.histogram()
        total = im.size[0] * im.size[1]
    thr = int(max(0, min(nonwhite_threshold, 255)))
    nonwhite = sum(hist[:thr])
    return float(nonwhite) / float(total)


@dataclass(frozen=True)
class PageStat:
    page: int
    nonwhite_ratio: float
    preview_png: str


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Find near-blank / low-ink pages in a PDF (render at low DPI via pdftoppm)."
    )
    ap.add_argument("--pdf", type=Path, required=True, help="PDF path")
    ap.add_argument("--page-start", type=int, default=1, help="1-based start page (inclusive)")
    ap.add_argument("--page-end", type=int, default=0, help="1-based end page (inclusive). 0=auto via pdfinfo")
    ap.add_argument("--dpi", type=int, default=30, help="render dpi (low dpi is enough; default 30)")
    ap.add_argument(
        "--nonwhite-threshold",
        type=int,
        default=245,
        help="pixel < threshold is treated as non-white (0..255), default 245",
    )
    ap.add_argument(
        "--ratio-lt",
        type=float,
        default=0.015,
        help="print pages with nonwhite_ratio < ratio_lt (default 0.015 ~= chapter separators/blank pages)",
    )
    ap.add_argument("--top-n", type=int, default=30, help="also print lowest top-n pages (default 30)")
    ap.add_argument("--out-json", type=Path, default=None, help="optional output JSON path")
    ap.add_argument(
        "--keep-dir",
        type=Path,
        default=None,
        help="if set, keep rendered PNGs in this dir (otherwise tempdir is removed)",
    )
    args = ap.parse_args()

    pdf = args.pdf
    if not pdf.exists():
        raise SystemExit(f"PDF not found: {pdf}")

    page_end = int(args.page_end)
    if page_end <= 0:
        pc = _pdf_page_count(pdf)
        if not pc:
            raise SystemExit("page_end is 0 but page_count could not be detected (missing pdfinfo).")
        page_end = pc

    page_start = max(1, int(args.page_start))
    page_end = max(page_start, int(page_end))

    if args.keep_dir:
        out_dir = args.keep_dir
        tmpdir: Optional[tempfile.TemporaryDirectory[str]] = None
    else:
        tmpdir = tempfile.TemporaryDirectory(prefix="pdf_sparse_pages_")
        out_dir = Path(tmpdir.name)

    try:
        pngs = _render_range(pdf, page_start=page_start, page_end=page_end, dpi=args.dpi, out_dir=out_dir)
        stats: list[PageStat] = []
        for p in pngs:
            page = _parse_page_from_filename(p)
            r = _nonwhite_ratio(p, nonwhite_threshold=args.nonwhite_threshold)
            stats.append(PageStat(page=page, nonwhite_ratio=r, preview_png=str(p)))

        stats_sorted = sorted(stats, key=lambda x: x.nonwhite_ratio)

        print(f"PDF: {pdf}")
        print(f"Pages scanned: {page_start}..{page_end} (dpi={args.dpi}, nonwhite_threshold={args.nonwhite_threshold})")
        print("")

        # top-n
        print(f"Lowest {min(args.top_n, len(stats_sorted))} pages:")
        for s in stats_sorted[: args.top_n]:
            print(f"- p{s.page:>4} ratio={s.nonwhite_ratio:.6f}  {s.preview_png}")
        print("")

        # below threshold
        picked = [s for s in stats_sorted if s.nonwhite_ratio < float(args.ratio_lt)]
        print(f"Pages with ratio < {args.ratio_lt}: {len(picked)}")
        if picked:
            print(" ", [s.page for s in picked])
        print("")

        if args.out_json:
            payload = {
                "pdf_path": str(pdf),
                "page_start": page_start,
                "page_end": page_end,
                "dpi": int(args.dpi),
                "nonwhite_threshold": int(args.nonwhite_threshold),
                "ratio_lt": float(args.ratio_lt),
                "stats": [
                    {"page": s.page, "nonwhite_ratio": s.nonwhite_ratio, "preview_png": s.preview_png}
                    for s in stats_sorted
                ],
            }
            args.out_json.parent.mkdir(parents=True, exist_ok=True)
            args.out_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"OK wrote JSON: {args.out_json}")
    finally:
        if tmpdir is not None:
            tmpdir.cleanup()


if __name__ == "__main__":
    main()

