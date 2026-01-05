#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import platform
import socket
import subprocess
import sys
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


def _read_json(path: Path) -> dict[str, Any]:
    obj = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(obj, dict):
        raise ValueError(f"job.json must be an object: {path}")
    return obj


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _as_rel(path: Path, base: Path) -> str:
    try:
        return str(path.resolve().relative_to(base.resolve())).replace("\\", "/")
    except Exception:
        return str(path)


def _trunc(text: str, limit: int = 8000) -> str:
    s = (text or "").strip()
    if len(s) <= limit:
        return s
    return s[:limit] + "\n...(truncated)..."

def _infer_user_action(*, stage: str, error: str, runner_output: str, backend: Backend) -> str:
    """
    Best-effort guidance for operators. Keep it short; never include secrets.
    """
    e = (error or "").lower()
    out = (runner_output or "").lower()

    if stage == "VALIDATE_INPUT":
        return "Ensure job/in contains the input DOCX and rerun the runner."
    if stage == "VERIFY_INPUT_HASH":
        return "Re-queue the job from the source DOCX (sha256 mismatch indicates wrong/changed input)."
    if stage == "CONVERT":
        if "microsoft word.app not found" in out or "word mode requires" in out:
            return "Install Microsoft Word on the runner, or re-queue with backend=pages/soffice."
        if "pages mode requires pages.app" in out or "pages.app not found" in out:
            return "Install Pages on the runner, or re-queue with backend=soffice."
        if "soffice not found" in out or "install libreoffice" in out:
            return "Install LibreOffice(soffice) on the runner, or re-queue with backend=pages/word."
        if "all conversion attempts failed" in out and backend == "auto":
            return "Retry with an explicit backend (--backend pages|word|soffice) based on runner OS."
        return "Inspect runner_output, fix runner environment, and retry."
    if stage == "VERIFY_OUTPUT":
        return "Conversion finished but output PDF is missing; retry with --force and inspect runner_output."
    if stage == "HASH_OUTPUT":
        return "Output PDF exists but hashing failed; verify file permissions and retry."

    return "Inspect conversion_log.json and retry after fixing the underlying issue."


def _infer_retryable(*, stage: str, error: str) -> bool:
    e = (error or "").lower()
    if stage == "VERIFY_INPUT_HASH":
        return False
    if "permission" in e:
        return True
    if stage in {"VALIDATE_INPUT", "CONVERT", "VERIFY_OUTPUT", "HASH_OUTPUT"}:
        return True
    return False


def _script_postprocess_path() -> Path:
    # sibling script
    return (Path(__file__).resolve().parent / "postprocess_docx_to_pdf.py").resolve()


def _run_postprocess(*, docx: Path, out_pdf: Path, backend: Backend, update_fields: bool) -> tuple[int, str]:
    script = _script_postprocess_path()
    cmd: list[str] = [
        str(sys.executable),
        str(script),
        "--docx",
        str(docx.resolve()),
        "--out-pdf",
        str(out_pdf.resolve()),
        "--mode",
        str(backend),
    ]
    if not update_fields:
        cmd.append("--no-update-fields")

    p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    return int(p.returncode), str(p.stdout or "")


@dataclass(frozen=True)
class _Job:
    job_dir: Path
    job_id: str
    backend: Backend
    update_fields: bool
    in_docx: Path
    out_pdf: Path
    expected_docx_sha256: str
    log_dir: Path


def _load_job(job_dir: Path) -> _Job:
    job_json = job_dir / "job.json"
    payload = _read_json(job_json)

    job_id = str(payload.get("job_id") or "").strip() or job_dir.name
    requested = payload.get("requested") if isinstance(payload.get("requested"), dict) else {}
    backend_raw = str((requested or {}).get("backend") or "auto").strip().lower()
    if backend_raw not in {"auto", "word", "pages", "soffice"}:
        raise ValueError(f"Unsupported requested.backend={backend_raw!r} (job_id={job_id})")
    backend: Backend = backend_raw  # type: ignore[assignment]
    update_fields = bool((requested or {}).get("update_fields", False))

    input_docx = (requested or {}).get("input_docx") if isinstance((requested or {}).get("input_docx"), dict) else {}
    input_rel = str((input_docx or {}).get("path") or "in/report.docx").strip()
    expected_docx_sha256 = str((input_docx or {}).get("sha256") or "").strip()

    output_pdf = (requested or {}).get("output_pdf") if isinstance((requested or {}).get("output_pdf"), dict) else {}
    output_rel = str((output_pdf or {}).get("path") or "out/report.pdf").strip()

    in_docx = (job_dir / input_rel).resolve()
    out_pdf = (job_dir / output_rel).resolve()

    runner_contract = payload.get("runner_contract") if isinstance(payload.get("runner_contract"), dict) else {}
    log_dir_rel = str((runner_contract or {}).get("log_dir") or "log").strip() or "log"
    log_dir = (job_dir / log_dir_rel).resolve()

    return _Job(
        job_dir=job_dir.resolve(),
        job_id=job_id,
        backend=backend,
        update_fields=update_fields,
        in_docx=in_docx,
        out_pdf=out_pdf,
        expected_docx_sha256=expected_docx_sha256,
        log_dir=log_dir,
    )


def _iter_job_dirs(job_root: Path) -> list[Path]:
    if not job_root.exists():
        return []
    dirs: list[Path] = []
    for p in sorted(job_root.iterdir()):
        if not p.is_dir():
            continue
        if (p / "job.json").exists():
            dirs.append(p)
    return dirs


def _process_job(job: _Job, *, dry_run: bool, force: bool) -> int:
    started = _utc_now_iso()
    t0 = time.time()

    conversion_log = job.log_dir / "conversion_log.json"
    done_json = job.log_dir / "done.json"

    warnings: list[str] = []

    stage = "INIT"
    docx_sha = ""
    sha_match = None
    status: str = "ERROR"
    err: str | None = None
    runner_out = ""
    retryable: bool | None = None
    user_action: str = ""

    try:
        stage = "VALIDATE_INPUT"
        if not job.in_docx.exists():
            raise FileNotFoundError(f"missing input docx: {job.in_docx}")

        stage = "VERIFY_INPUT_HASH"
        docx_sha = _sha256_file(job.in_docx)
        if job.expected_docx_sha256:
            sha_match = (docx_sha == job.expected_docx_sha256)
            if not sha_match:
                raise RuntimeError(
                    f"sha256 mismatch for input docx (expected={job.expected_docx_sha256}, actual={docx_sha})"
                )

        if job.update_fields and job.backend != "word":
            warnings.append("update_fields_requested_but_backend_not_word")

        stage = "CONVERT"
        if job.out_pdf.exists() and not force:
            status = "OK"
            warnings.append("skipped_conversion_out_pdf_exists")
        elif dry_run:
            status = "OK"
            warnings.append("dry_run_no_conversion")
        else:
            rc, out = _run_postprocess(
                docx=job.in_docx,
                out_pdf=job.out_pdf,
                backend=job.backend,
                update_fields=job.update_fields,
            )
            runner_out = out
            if rc != 0:
                raise RuntimeError(f"postprocess_docx_to_pdf failed (exit={rc})")

            stage = "VERIFY_OUTPUT"
            if not job.out_pdf.exists():
                raise RuntimeError(f"expected pdf not produced: {job.out_pdf}")

            status = "OK"
    except Exception as e:
        err = f"{type(e).__name__}: {e}"
        status = "ERROR"

    ended = _utc_now_iso()
    elapsed_ms = int(round((time.time() - t0) * 1000.0))

    pdf_sha = ""
    if status == "OK" and job.out_pdf.exists():
        try:
            stage = "HASH_OUTPUT"
            pdf_sha = _sha256_file(job.out_pdf)
        except Exception as e:
            warnings.append(f"sha256_pdf_failed:{type(e).__name__}")

    if status != "OK":
        retryable = _infer_retryable(stage=stage, error=err or "")
        user_action = _infer_user_action(stage=stage, error=err or "", runner_output=runner_out, backend=job.backend)

    log_payload: dict[str, Any] = {
        "schema_version": "1.0",
        "job_id": job.job_id,
        "status": status,
        "failure_stage": stage if status != "OK" else "",
        "retryable": retryable,
        "user_action": user_action,
        "started_at": started,
        "ended_at": ended,
        "elapsed_ms": elapsed_ms,
        "requested": {
            "backend": job.backend,
            "update_fields": bool(job.update_fields),
        },
        "runner": {
            "host": socket.gethostname(),
            "platform": platform.platform(),
            "python": sys.version.split()[0],
            "script": _as_rel(Path(__file__).resolve(), job.job_dir),
        },
        "input_docx": {
            "path": _as_rel(job.in_docx, job.job_dir),
            "sha256_expected": job.expected_docx_sha256,
            "sha256_actual": docx_sha,
            "sha256_match": sha_match,
        },
        "output_pdf": {
            "path": _as_rel(job.out_pdf, job.job_dir),
            "sha256": pdf_sha,
        },
        "warnings": warnings,
        "runner_output": _trunc(runner_out),
        "error": err or "",
    }
    _write_json(conversion_log, log_payload)

    done_payload: dict[str, Any] = {
        "schema_version": "1.0",
        "job_id": job.job_id,
        "status": status,
        "failure_stage": stage if status != "OK" else "",
        "retryable": retryable,
        "user_action": user_action,
        "finished_at": ended,
        "output_pdf": {"path": _as_rel(job.out_pdf, job.job_dir), "sha256": pdf_sha},
        "warnings": warnings,
        "error": err or "",
    }
    _write_json(done_json, done_payload)

    # Exit code: 0 OK, 2 when skipped, 1 on error
    if status != "OK":
        return 1
    if "skipped_conversion_out_pdf_exists" in warnings or "dry_run_no_conversion" in warnings:
        return 2
    return 0


def main() -> None:
    ap = argparse.ArgumentParser(
        description=(
            "Run Antigravity DOCX->PDF jobs (runner-side helper).\n"
            "Reads <job_dir>/job.json created by antigravity_queue_docx_to_pdf.py and writes log/ files."
        )
    )
    ap.add_argument(
        "--job-root",
        type=Path,
        default=Path("output/_antigravity_jobs"),
        help="Job root directory (default: output/_antigravity_jobs)",
    )
    ap.add_argument("--job-id", type=str, default="", help="Process only this job_id (folder name)")
    ap.add_argument("--dry-run", action="store_true", help="Validate and write logs without conversion")
    ap.add_argument("--force", action="store_true", help="Re-run even if out/pdf already exists")
    args = ap.parse_args()

    job_root = args.job_root.expanduser().resolve()
    job_id = args.job_id.strip()

    job_dirs: list[Path] = []
    if job_id:
        job_dirs = [job_root / job_id]
    else:
        job_dirs = _iter_job_dirs(job_root)

    if not job_dirs:
        raise SystemExit(f"No jobs found under: {job_root}")

    worst = 0
    for d in job_dirs:
        job_json = d / "job.json"
        if not job_json.exists():
            print(f"SKIP missing job.json: {d}")
            worst = max(worst, 2)
            continue
        job = _load_job(d)
        rc = _process_job(job, dry_run=bool(args.dry_run), force=bool(args.force))
        if rc == 0:
            print(f"OK processed job_id={job.job_id} out_pdf={_as_rel(job.out_pdf, job.job_dir)}")
        elif rc == 2:
            print(f"OK (skipped) job_id={job.job_id} out_pdf={_as_rel(job.out_pdf, job.job_dir)}")
        else:
            print(f"ERROR job_id={job.job_id} (see {job.log_dir}/conversion_log.json)")
        worst = max(worst, rc)

    raise SystemExit(worst if worst in {0, 1, 2} else 1)


if __name__ == "__main__":
    main()
