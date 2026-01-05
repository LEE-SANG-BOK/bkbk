#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path


def _to_float(v: str, *, name: str) -> float:
    try:
        return float(str(v).strip())
    except Exception as e:  # pragma: no cover
        raise SystemExit(f"Invalid {name}: {v!r} ({e})")


def main() -> None:
    ap = argparse.ArgumentParser(description="Set DOCX page size/margins (best-effort, in-place by default).")
    ap.add_argument("--docx", required=True, type=Path, help="Input .docx path")
    ap.add_argument("--out", type=Path, default=None, help="Output .docx path (default: in-place)")
    ap.add_argument("--page", choices=["A4", "LETTER"], default="A4", help="Target page size")
    ap.add_argument("--margin-cm", default="2.0", help="Uniform margins in cm (default: 2.0)")

    args = ap.parse_args()

    docx_path = args.docx.expanduser().resolve()
    if not docx_path.exists():
        raise SystemExit(f"docx not found: {docx_path}")

    out_path = args.out.expanduser().resolve() if args.out else docx_path
    margin_cm = _to_float(str(args.margin_cm), name="margin-cm")

    try:
        from docx import Document  # type: ignore
        from docx.shared import Cm, Mm  # type: ignore
    except Exception as e:  # pragma: no cover
        raise SystemExit(f"Missing python-docx dependency: {e}")

    doc = Document(str(docx_path))

    if str(args.page).upper() == "A4":
        page_w = Mm(210)
        page_h = Mm(297)
    else:
        page_w = Mm(215.9)
        page_h = Mm(279.4)

    for sec in doc.sections:
        sec.page_width = page_w
        sec.page_height = page_h
        sec.top_margin = Cm(margin_cm)
        sec.bottom_margin = Cm(margin_cm)
        sec.left_margin = Cm(margin_cm)
        sec.right_margin = Cm(margin_cm)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    print(f"OK wrote: {out_path}")


if __name__ == "__main__":
    main()

