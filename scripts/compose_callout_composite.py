#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
from pathlib import Path
from typing import Any

from eia_gen.services.figures.callout_composite import CalloutItem, compose_callout_composite


def _as_str(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _parse_list(s: str) -> list[str]:
    out: list[str] = []
    for token in s.replace(";", ",").split(","):
        t = token.strip()
        if t:
            out.append(t)
    return out


def main() -> None:
    ap = argparse.ArgumentParser(
        description=(
            "Compose a deterministic 'photo board' panel (CALLOUT_COMPOSITE) with fixed grids (2/4/6).\n"
            "This script is a thin CLI wrapper around core library implementation."
        )
    )
    ap.add_argument("--images", nargs="*", default=None, help="Image paths (2~6).")
    ap.add_argument("--images-csv", type=Path, default=None, help="CSV with columns: path,caption (optional).")
    ap.add_argument("--captions", type=str, default="", help="Optional captions (comma/; separated).")
    ap.add_argument("--grid", type=int, default=None, help="Grid size: 2/4/6 (auto if omitted).")
    ap.add_argument("--title", type=str, default="", help="Optional title text (drawn top-left).")
    ap.add_argument("--out", type=Path, required=True, help="Output PNG path.")
    ap.add_argument("--style", type=Path, default=Path("config/figure_style.yaml"), help="Style tokens YAML path.")
    ap.add_argument("--font-path", type=str, default="", help="Optional font path (overrides EIA_GEN_FONT_PATH).")
    args = ap.parse_args()

    repo_root = Path(__file__).resolve().parents[1]  # eia-gen/
    style_path = args.style
    if not style_path.is_absolute():
        style_path = (repo_root / style_path).resolve()
    items: list[CalloutItem] = []

    # 1) CSV items (highest priority if provided)
    if args.images_csv:
        csv_path = args.images_csv.expanduser().resolve()
        if not csv_path.exists():
            raise SystemExit(f"images-csv not found: {csv_path}")
        with csv_path.open("r", encoding="utf-8", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                p = Path(_as_str(row.get("path"))).expanduser()
                if not p.is_absolute():
                    p = (repo_root / p).resolve()
                cap = _as_str(row.get("caption")) or p.name
                if p.exists():
                    items.append(CalloutItem(path=p, caption=cap))

    # 2) Direct images list
    if not items and args.images:
        caps = _parse_list(args.captions)
        for i, raw in enumerate(args.images):
            p = Path(raw).expanduser()
            if not p.is_absolute():
                p = (repo_root / p).resolve()
            if not p.exists():
                raise SystemExit(f"image not found: {p}")
            cap = caps[i] if i < len(caps) else p.name
            items.append(CalloutItem(path=p, caption=cap))

    if not items:
        raise SystemExit("No images provided. Use --images ... or --images-csv ...")

    out = args.out.expanduser()
    if not out.is_absolute():
        out = (repo_root / out).resolve()
    compose_callout_composite(
        items=items,
        out_path=out,
        title=_as_str(args.title),
        grid=args.grid,
        style_path=style_path,
        font_path=_as_str(args.font_path) or None,
    )
    print(f"WROTE: {out}")


if __name__ == "__main__":
    main()
