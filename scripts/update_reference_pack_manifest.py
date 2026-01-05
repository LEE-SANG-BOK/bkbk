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
        description=(
            "Update sha256 entries in reference pack manifest.json (best-effort). "
            "This is useful after regenerating pack files (xlsx/md/png) while keeping stable paths."
        )
    )
    ap.add_argument(
        "--pack-dir",
        type=Path,
        required=True,
        help="Reference pack dir (must contain manifest.json)",
    )
    ap.add_argument("--dry-run", action="store_true", help="Print planned changes without writing")
    ap.add_argument("--fail-on-missing", action="store_true", help="Exit non-zero if any file is missing")
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
    changed: list[str] = []
    unchanged = 0

    for entry in files:
        if not isinstance(entry, dict):
            continue
        rel = str(entry.get("file_path") or "").strip()
        if not rel:
            continue
        p = (pack_dir / rel).resolve()
        if not p.exists():
            missing.append(rel)
            continue
        actual = _sha256_file(p)
        expected = str(entry.get("sha256") or "").strip().lower()
        if expected != actual.lower():
            changed.append(f"{rel} {expected or '(empty)'} -> {actual}")
            entry["sha256"] = actual
        else:
            unchanged += 1

    if missing:
        print("MISSING:")
        for x in missing:
            print(f" - {x}")
    if changed:
        print("UPDATED:")
        for x in changed:
            print(f" - {x}")
    print(f"SUMMARY: unchanged={unchanged} updated={len(changed)} missing={len(missing)}")

    if args.fail_on_missing and missing:
        raise SystemExit(2)

    if args.dry_run:
        print("DRY-RUN: not writing manifest.json")
        return

    manifest_path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK wrote: {manifest_path}")


if __name__ == "__main__":
    main()
