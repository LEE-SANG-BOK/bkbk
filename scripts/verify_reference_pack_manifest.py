#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path


def _sha256_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Verify sha256 entries in reference pack manifest.json (best-effort integrity check)."
    )
    ap.add_argument(
        "--pack-dir",
        type=Path,
        required=True,
        help="Reference pack dir (must contain manifest.json)",
    )
    ap.add_argument("--fail-on-missing", action="store_true", help="Exit non-zero if any file is missing/mismatched")
    args = ap.parse_args()

    pack_dir = args.pack_dir.expanduser().resolve()
    manifest_path = pack_dir / "manifest.json"
    if not manifest_path.exists():
        raise SystemExit(f"manifest.json not found: {manifest_path}")

    obj = json.loads(manifest_path.read_text(encoding="utf-8"))
    files = obj.get("files")
    if not isinstance(files, list):
        raise SystemExit("manifest.json must contain a top-level 'files' list")

    missing: list[str] = []
    mismatched: list[str] = []

    for entry in files:
        if not isinstance(entry, dict):
            continue
        rel = str(entry.get("file_path") or "").strip()
        expected = str(entry.get("sha256") or "").strip().lower()
        if not rel or not expected:
            continue
        p = (pack_dir / rel).resolve()
        if not p.exists():
            missing.append(rel)
            continue
        actual = _sha256_file(p)
        if actual.lower() != expected:
            mismatched.append(f"{rel} expected={expected} actual={actual}")

    if missing:
        print("MISSING:")
        for x in missing:
            print(f" - {x}")
    if mismatched:
        print("MISMATCHED:")
        for x in mismatched:
            print(f" - {x}")

    if not missing and not mismatched:
        print("OK: all manifest sha256 entries match")

    if args.fail_on_missing and (missing or mismatched):
        raise SystemExit(2)


if __name__ == "__main__":
    main()

