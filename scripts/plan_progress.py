#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path


CHECKBOX_RE = re.compile(r"- \[( |x|X)\] ")
OPEN_CHECKBOX_LINE_RE = re.compile(r"^\s*- \[ \] ")


@dataclass(frozen=True)
class Progress:
    done: int
    total: int
    pct: float


def compute_progress(text: str) -> Progress:
    hits = CHECKBOX_RE.findall(text)
    total = len(hits)
    done = sum(1 for h in hits if h.lower() == "x")
    pct = round((done / total * 100.0), 1) if total else 0.0
    return Progress(done=done, total=total, pct=pct)

def list_open_items(text: str) -> list[tuple[int, str]]:
    """
    Return (line_no, line_text) for unchecked checkboxes in the plan markdown.
    """
    out: list[tuple[int, str]] = []
    for i, line in enumerate(text.splitlines(), start=1):
        if OPEN_CHECKBOX_LINE_RE.match(line):
            out.append((i, line.rstrip("\n")))
    return out


def _update_exec_plan(text: str, p: Progress, today: str) -> tuple[str, int]:
    """
    Update the “체크리스트 진행율: ...” snippet if present.
    We keep the surrounding wording intact and only replace numbers/date.
    """
    pattern = re.compile(
        r"(체크리스트 진행율:\s*)([0-9.]+)%\s*=\s*(\d+)/(\d+),\s*(\d{4}-\d{2}-\d{2})\s*기준"
    )
    # Use \g<1> to avoid ambiguity when the next character is a digit (e.g., \154.8 -> octal escape).
    repl = rf"\g<1>{p.pct}% = {p.done}/{p.total}, {today} 기준"
    return pattern.subn(repl, text)


def _update_handoff(text: str, p: Progress, today: str) -> tuple[str, int]:
    pattern = re.compile(
        r"(체크리스트 진척률\(checkbox\):\s*)([0-9.]+)%\s*=\s*(\d+)/(\d+)\s*\((\d{4}-\d{2}-\d{2})\s*기준\)"
    )
    # Use \g<1> to avoid ambiguity when the next character is a digit (e.g., \154.8 -> octal escape).
    repl = rf"\g<1>{p.pct}% = {p.done}/{p.total} ({today} 기준)"
    return pattern.subn(repl, text)


def main() -> None:
    ap = argparse.ArgumentParser(
        description=(
            "Compute checkbox progress for docs/11_execution_plan.md.\n"
            "Optionally updates known progress snippets in plan/handoff docs to avoid manual drift."
        )
    )
    ap.add_argument(
        "--plan-md",
        type=Path,
        default=None,
        help="Execution plan markdown path (default: docs/11_execution_plan.md under repo root).",
    )
    ap.add_argument(
        "--update",
        action="store_true",
        help="Update progress snippets in the plan/handoff docs (in-place).",
    )
    ap.add_argument(
        "--list-open",
        action="store_true",
        help="List unchecked checkbox items (with line numbers).",
    )
    ap.add_argument(
        "--handoff-md",
        type=Path,
        default=Path("docs/30_handoff_agent_b_to_a.md"),
        help="Handoff markdown path to update when --update is set (repo-relative by default).",
    )
    ap.add_argument(
        "--date",
        default=None,
        help="Override date (YYYY-MM-DD). Default: today.",
    )
    args = ap.parse_args()

    repo_root = Path(__file__).resolve().parents[1]  # eia-gen/
    plan_md = args.plan_md or (repo_root / "docs" / "11_execution_plan.md")
    plan_md = plan_md.expanduser().resolve()
    if not plan_md.exists():
        raise SystemExit(f"plan.md not found: {plan_md}")

    today = args.date or date.today().isoformat()

    plan_text = plan_md.read_text(encoding="utf-8")
    p = compute_progress(plan_text)
    print(f"PLAN: {p.pct}% = {p.done}/{p.total} ({today}) :: {plan_md}")

    if args.list_open:
        open_items = list_open_items(plan_text)
        print(f"OPEN: {len(open_items)}")
        for line_no, line in open_items:
            print(f"- L{line_no}: {line}")

    if not args.update:
        return

    # Update plan snippet
    plan_updated, plan_subs = _update_exec_plan(plan_text, p, today)
    if plan_updated != plan_text:
        plan_md.write_text(plan_updated, encoding="utf-8")
        print(f"UPDATED: {plan_md}")
    else:
        reason = "already up-to-date" if plan_subs > 0 else "no progress snippet match"
        print(f"UNCHANGED: {plan_md} ({reason})")

    # Update handoff snippet (best-effort)
    handoff_md = args.handoff_md
    if not handoff_md.is_absolute():
        handoff_md = (repo_root / handoff_md).resolve()
    if handoff_md.exists():
        handoff_text = handoff_md.read_text(encoding="utf-8")
        handoff_updated, handoff_subs = _update_handoff(handoff_text, p, today)
        if handoff_updated != handoff_text:
            handoff_md.write_text(handoff_updated, encoding="utf-8")
            print(f"UPDATED: {handoff_md}")
        else:
            reason = "already up-to-date" if handoff_subs > 0 else "no progress snippet match"
            print(f"UNCHANGED: {handoff_md} ({reason})")
    else:
        print(f"SKIP: handoff.md not found: {handoff_md}")


if __name__ == "__main__":
    main()
