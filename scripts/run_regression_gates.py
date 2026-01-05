#!/usr/bin/env python3
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def _bin_python(repo_root: Path) -> Path:
    cand = repo_root / ".venv" / "bin" / "python"
    if cand.exists():
        return cand
    return Path(sys.executable)


def _run(cmd: list[str], *, cwd: Path) -> int:
    p = subprocess.Popen(cmd, cwd=str(cwd))
    return p.wait()


def main() -> None:
    ap = argparse.ArgumentParser(description="Run run_quality_gates.py for baseline regression cases.")
    ap.add_argument(
        "--cases",
        nargs="*",
        default=["output/case_new", "output/case_new_max_reuse"],
        help="Case dirs to run (repo-relative or absolute).",
    )
    ap.add_argument(
        "--mode",
        default="check",
        choices=["check", "generate", "both"],
        help="Pass-through for run_quality_gates.py.",
    )
    ap.add_argument("--skip-unit-tests", action="store_true", help="Skip unit tests (otherwise run once first).")
    ap.add_argument("--doctor-relaxed", action="store_true", help="Pass-through for run_quality_gates.py.")
    ap.add_argument("--verify-keys", action="store_true", help="Pass-through for run_quality_gates.py.")
    ap.add_argument("--verify-keys-strict", action="store_true", help="Pass-through for run_quality_gates.py.")
    ap.add_argument(
        "--verify-keys-mode",
        default=None,
        choices=["presence", "network"],
        help="Pass-through for run_quality_gates.py (default: gate default).",
    )
    ap.add_argument("--verify-keys-wms-use-cache", action="store_true", help="Pass-through for run_quality_gates.py.")
    ap.add_argument(
        "--verify-keys-strict-ignore-category",
        action="append",
        default=[],
        help="Pass-through for run_quality_gates.py (repeatable).",
    )
    ap.add_argument("--enrich", action="store_true", help="Pass-through for run_quality_gates.py.")
    ap.add_argument("--enrich-overwrite-plan", action="store_true", help="Pass-through for run_quality_gates.py.")
    ap.add_argument("--enrich-fail-on-warn", action="store_true", help="Pass-through for run_quality_gates.py.")
    args = ap.parse_args()

    repo_root = _repo_root()
    py = _bin_python(repo_root)

    if not args.skip_unit_tests:
        rc = _run([str(py), "-m", "unittest", "discover", "-s", "tests", "-p", "test_*.py"], cwd=repo_root)
        if rc != 0:
            raise SystemExit(rc)

    failures: list[tuple[str, int]] = []
    for case_dir in args.cases:
        cmd = [
            str(py),
            "scripts/run_quality_gates.py",
            "--case-dir",
            case_dir,
            "--mode",
            args.mode,
            "--skip-unit-tests",
        ]
        if args.doctor_relaxed:
            cmd.append("--doctor-relaxed")
        if args.verify_keys:
            cmd.append("--verify-keys")
        if args.verify_keys_strict:
            cmd.append("--verify-keys-strict")
        if args.verify_keys_mode:
            cmd.extend(["--verify-keys-mode", str(args.verify_keys_mode)])
        if args.verify_keys_wms_use_cache:
            cmd.append("--verify-keys-wms-use-cache")
        for cat in (args.verify_keys_strict_ignore_category or []):
            s = str(cat or "").strip()
            if not s:
                continue
            cmd.extend(["--verify-keys-strict-ignore-category", s])
        if args.enrich:
            cmd.append("--enrich")
        if args.enrich_overwrite_plan:
            cmd.append("--enrich-overwrite-plan")
        if args.enrich_fail_on_warn:
            cmd.append("--enrich-fail-on-warn")

        print("", flush=True)
        print(f"== Regression: {case_dir} ==", flush=True)
        rc = _run(cmd, cwd=repo_root)
        if rc != 0:
            failures.append((case_dir, rc))

    if failures:
        print("", flush=True)
        print("FAILED cases:", flush=True)
        for case_dir, rc in failures:
            print(f"- {case_dir}: exit={rc}", flush=True)
        raise SystemExit(1)

    print("", flush=True)
    print("OK all regression cases passed", flush=True)


if __name__ == "__main__":
    main()
