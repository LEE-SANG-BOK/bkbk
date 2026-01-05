#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import re
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Literal


Backend = Literal["word", "pages", "soffice", "auto"]


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _safe_slug(s: str) -> str:
    """
    Restrictive slug for filesystem-safe job_id.
    """
    s = s.strip()
    s = re.sub(r"[^A-Za-z0-9._-]+", "-", s)
    s = re.sub(r"-{2,}", "-", s).strip("-")
    return s or "job"


@dataclass(frozen=True)
class JobPaths:
    job_dir: Path
    in_dir: Path
    out_dir: Path
    log_dir: Path
    job_json: Path
    in_docx: Path
    expected_pdf: Path


def _resolve_job_paths(job_root: Path, *, job_id: str, docx: Path, out_pdf_name: str) -> JobPaths:
    job_dir = (job_root / job_id).resolve()
    in_dir = job_dir / "in"
    out_dir = job_dir / "out"
    log_dir = job_dir / "log"
    job_json = job_dir / "job.json"
    in_docx = in_dir / docx.name
    expected_pdf = out_dir / out_pdf_name
    return JobPaths(
        job_dir=job_dir,
        in_dir=in_dir,
        out_dir=out_dir,
        log_dir=log_dir,
        job_json=job_json,
        in_docx=in_docx,
        expected_pdf=expected_pdf,
    )


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> None:
    ap = argparse.ArgumentParser(
        description=(
            "Queue a DOCX->PDF postprocess job for an external/remote runner (Antigravity).\n"
            "This script only prepares an on-disk job folder (copy docx + job.json)."
        )
    )
    ap.add_argument("--docx", required=True, type=Path, help="Input report_*.docx path")
    ap.add_argument(
        "--job-root",
        type=Path,
        default=Path("output/_antigravity_jobs"),
        help="Job root directory (default: output/_antigravity_jobs)",
    )
    ap.add_argument(
        "--backend",
        choices=["auto", "word", "pages", "soffice"],
        default="auto",
        help="Preferred conversion backend on the remote runner",
    )
    ap.add_argument("--update-fields", action="store_true", help="Request field/TOC updates (Word-capable runner)")
    ap.add_argument("--out-pdf-name", type=str, default="", help="Output PDF file name (default: <docx.stem>.pdf)")
    ap.add_argument(
        "--job-id",
        type=str,
        default="",
        help="Override job_id (default: AG-<docx.stem>-<sha256[:8]>)",
    )
    ap.add_argument(
        "--wait",
        type=int,
        default=0,
        help="Wait N seconds for out/<pdf> to appear (0 = do not wait)",
    )
    ap.add_argument("--poll-sec", type=float, default=2.0, help="Polling interval when --wait > 0")

    args = ap.parse_args()

    docx = args.docx.expanduser().resolve()
    if not docx.exists():
        raise SystemExit(f"docx not found: {docx}")

    job_root = args.job_root.expanduser().resolve()
    job_root.mkdir(parents=True, exist_ok=True)

    docx_sha256 = _sha256_file(docx)

    out_pdf_name = args.out_pdf_name.strip() or f"{docx.stem}.pdf"
    job_id = args.job_id.strip()
    if not job_id:
        job_id = f"AG-{_safe_slug(docx.stem)}-{docx_sha256[:8]}"
    job_id = _safe_slug(job_id)

    p = _resolve_job_paths(job_root, job_id=job_id, docx=docx, out_pdf_name=out_pdf_name)
    p.in_dir.mkdir(parents=True, exist_ok=True)
    p.out_dir.mkdir(parents=True, exist_ok=True)
    p.log_dir.mkdir(parents=True, exist_ok=True)

    # Copy input docx into job/in/ (overwrite OK; job_id is sha-based by default)
    p.in_docx.write_bytes(docx.read_bytes())

    job_payload: dict[str, Any] = {
        "schema_version": "1.0",
        "job_id": job_id,
        "created_at": _utc_now_iso(),
        "requested": {
            "backend": args.backend,
            "update_fields": bool(args.update_fields),
            "input_docx": {
                "path": str(p.in_docx.relative_to(p.job_dir).as_posix()),
                "sha256": docx_sha256,
            },
            "output_pdf": {
                "path": str(p.expected_pdf.relative_to(p.job_dir).as_posix()),
            },
            "notes": "This job requests DOCX postprocess only. Do not generate content.",
        },
        "runner_contract": {
            "write_pdf": True,
            "write_log_json": True,
            "write_done_json": True,
            "log_dir": str(p.log_dir.relative_to(p.job_dir).as_posix()),
        },
    }
    _write_json(p.job_json, job_payload)

    print("OK queued Antigravity job")
    print(f"- job_dir: {p.job_dir}")
    print(f"- job_json: {p.job_json}")
    print(f"- in_docx: {p.in_docx} (sha256={docx_sha256})")
    print(f"- expected_pdf: {p.expected_pdf}")
    print("")
    print("Next (runner side): convert in/<docx> -> out/<pdf> and write log/done json under log/")

    if args.wait > 0:
        deadline = time.time() + float(args.wait)
        while time.time() < deadline:
            if p.expected_pdf.exists():
                print(f"OK found PDF: {p.expected_pdf}")
                return
            time.sleep(max(0.2, float(args.poll_sec)))
        raise SystemExit(f"Timed out waiting for: {p.expected_pdf}")


if __name__ == "__main__":
    main()

