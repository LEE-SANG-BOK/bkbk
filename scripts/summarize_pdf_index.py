#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from collections import defaultdict
from pathlib import Path
from typing import Any


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--index", required=True, type=Path, help="combined_index.json (extract_pdf_index_twopass.py)")
    ap.add_argument("--out", required=True, type=Path, help="markdown 요약 파일")
    args = ap.parse_args()

    data: dict[str, Any] = json.loads(args.index.read_text(encoding="utf-8"))
    hits = data.get("pass2", {}).get("hits", [])

    chapters: dict[str, list[dict[str, Any]]] = defaultdict(list)
    figures: list[dict[str, Any]] = []
    tables: list[dict[str, Any]] = []
    for h in hits:
        if h["kind"] == "chapter":
            chapters[h["label"]].append(h)
        elif h["kind"] == "figure":
            figures.append(h)
        elif h["kind"] == "table":
            tables.append(h)

    def _fmt_hit(h: dict[str, Any]) -> str:
        return f'- p{h["page"]}: `{h["label"]}` {h["text"]}'

    lines: list[str] = []
    lines.append(f"# PDF Index Summary\n")
    lines.append(f"- pdf: `{data.get('pdf_path','')}`")
    lines.append(f"- pages: `{data.get('page_count','')}`")
    lines.append(f"- pass1 candidate pages: `{len(data.get('pass1', {}).get('candidate_pages', []))}`")
    lines.append(f"- extracted: chapters `{sum(len(v) for v in chapters.values())}`, figures `{len(figures)}`, tables `{len(tables)}`\n")

    lines.append("## Chapters")
    for ch in sorted(chapters.keys(), key=lambda s: int(s.replace("CH", "")) if s.startswith("CH") else 999):
        # 여러 번 잡히면 최초 페이지(작은 값)만 대표로
        rep = sorted(chapters[ch], key=lambda x: x["page"])[0]
        lines.append(_fmt_hit(rep))
    lines.append("")

    lines.append("## Figures (captions)")
    for h in sorted(figures, key=lambda x: (x["label"], x["page"])):
        lines.append(_fmt_hit(h))
    lines.append("")

    lines.append("## Tables (captions)")
    for h in sorted(tables, key=lambda x: (x["label"], x["page"])):
        lines.append(_fmt_hit(h))
    lines.append("")

    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text("\n".join(lines), encoding="utf-8")
    print(f"OK wrote {args.out}")


if __name__ == "__main__":
    main()

