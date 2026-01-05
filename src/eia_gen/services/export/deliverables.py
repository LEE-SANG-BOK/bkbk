from __future__ import annotations

import os
import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from eia_gen.models.case import Case

_TS_RE = re.compile(r"^[0-9]{8}_[0-9]{6}$")


@dataclass(frozen=True)
class DeliverablesResult:
    copied: list[Path]
    skipped: list[str]


def _expand_path(value: str) -> Path:
    return Path(os.path.expandvars(os.path.expanduser(value))).resolve()


def _sanitize_component(value: str) -> str:
    s = (value or "").strip()
    if not s:
        return "case"
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^\w.\-]+", "_", s, flags=re.UNICODE)
    s = re.sub(r"_+", "_", s)
    s = s.strip("._-")
    return s or "case"


def _detect_run_ts(out_dir: Path) -> str:
    # quality gates layout: .../_quality_gates/<ts>/generate
    parent = out_dir.parent
    if parent.parent.name == "_quality_gates" and _TS_RE.match(parent.name):
        return parent.name
    # fallback: use current time
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _pick_tag(*, case: Case, xlsx_path: Path, tag_override: str | None) -> str:
    if tag_override and tag_override.strip():
        return tag_override.strip()

    meta_extra = getattr(case.meta, "model_extra", None)
    if isinstance(meta_extra, dict):
        case_id = str(meta_extra.get("case_id") or "").strip()
        if case_id:
            return case_id

    return xlsx_path.parent.name


def copy_docx_deliverable(
    *,
    report_path: Path,
    kind: str,
    case: Case,
    case_xlsx_path: Path,
    report_out_dir: Path,
    deliverables_dir: str | None,
    deliverables_tag: str | None,
) -> DeliverablesResult:
    """Best-effort: copy a generated DOCX to a deliverables directory.

    Naming:
      - report_eia.<tag>_<ts>.docx (kind=EIA)
      - report_dia.<tag>_<ts>.docx (kind=DIA)
      - report.<tag>_<ts>.docx (otherwise)
    """
    if not deliverables_dir or not str(deliverables_dir).strip():
        return DeliverablesResult(copied=[], skipped=["deliverables_dir not set"])

    src = report_path.resolve()
    if not src.exists():
        return DeliverablesResult(copied=[], skipped=[f"missing: {src}"])

    try:
        dst_dir = _expand_path(str(deliverables_dir).strip())
        dst_dir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        return DeliverablesResult(copied=[], skipped=[f"failed to prepare deliverables dir: {e}"])

    tag = _sanitize_component(_pick_tag(case=case, xlsx_path=case_xlsx_path, tag_override=deliverables_tag))
    ts = _detect_run_ts(report_out_dir)
    base = "report"
    if (kind or "").strip().upper() == "EIA":
        base = "report_eia"
    elif (kind or "").strip().upper() == "DIA":
        base = "report_dia"

    dst = dst_dir / f"{base}.{tag}_{ts}{src.suffix}"
    try:
        shutil.copy2(src, dst)
    except Exception as e:
        return DeliverablesResult(copied=[], skipped=[f"copy failed: {e}"])

    return DeliverablesResult(copied=[dst], skipped=[])

