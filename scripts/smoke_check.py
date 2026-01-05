from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def _bin_eia_gen(repo_root: Path) -> list[str]:
    # Prefer project venv when present.
    cand = repo_root / ".venv" / "bin" / "eia-gen"
    if cand.exists():
        return [str(cand)]
    # Fallback: module execution (requires PYTHONPATH/resolved install).
    return [sys.executable, "-m", "eia_gen.cli"]


def _run(cmd: list[str]) -> None:
    p = subprocess.run(cmd, stdout=sys.stdout, stderr=sys.stderr)
    if p.returncode != 0:
        raise SystemExit(p.returncode)


def main() -> None:
    ap = argparse.ArgumentParser(description="eia-gen smoke check (QA + optional full generation).")
    ap.add_argument(
        "--case-dir",
        default="output/case_changwon_2025",
        help="Case folder containing case.xlsx and sources.yaml",
    )
    ap.add_argument(
        "--mode",
        default="check",
        choices=["check", "generate", "both"],
        help="check=QA only, generate=generate docx, both=check then generate",
    )
    ap.add_argument(
        "--out-dir",
        default="",
        help="Output dir for generated reports (default: <case-dir>)",
    )
    args = ap.parse_args()

    repo_root = Path(__file__).resolve().parents[1]
    case_dir = (repo_root / args.case_dir).resolve() if not Path(args.case_dir).is_absolute() else Path(args.case_dir)
    xlsx = case_dir / "case.xlsx"
    sources = case_dir / "sources.yaml"

    if not xlsx.exists():
        raise SystemExit(f"missing: {xlsx}")
    if not sources.exists():
        raise SystemExit(f"missing: {sources}")

    out_dir = Path(args.out_dir).resolve() if args.out_dir else case_dir

    eia_gen = _bin_eia_gen(repo_root)

    if args.mode in ("check", "both"):
        _run(
            [
                *eia_gen,
                "check-xlsx-both",
                "--xlsx",
                str(xlsx),
                "--sources",
                str(sources),
                "--out-dir",
                str(out_dir / "_smoke_check"),
            ]
        )

    if args.mode in ("generate", "both"):
        _run(
            [
                *eia_gen,
                "generate-xlsx-both",
                "--xlsx",
                str(xlsx),
                "--sources",
                str(sources),
                "--out-dir",
                str(out_dir),
                "--no-use-llm",
            ]
        )

    print("OK smoke check complete")


if __name__ == "__main__":
    main()

