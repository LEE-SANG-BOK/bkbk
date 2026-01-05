from __future__ import annotations

import hashlib
import json
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml

from eia_gen.services.data_requests.xlsx_io import (
    apply_rows_to_sheet,
    load_workbook,
    read_sheet_dicts,
    save_workbook,
)


@dataclass(frozen=True)
class PackSpec:
    version: int
    pack_id: str
    title: str
    created_at: str
    reference_case_xlsx: str
    reference_sources_yaml: str
    apply_sheets: list[str]
    copy_attachments: bool
    notes: str

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "PackSpec":
        return cls(
            version=int(d.get("version") or 1),
            pack_id=str(d.get("pack_id") or "").strip(),
            title=str(d.get("title") or "").strip(),
            created_at=str(d.get("created_at") or "").strip(),
            reference_case_xlsx=str(d.get("reference_case_xlsx") or "reference_case.xlsx").strip(),
            reference_sources_yaml=str(d.get("reference_sources_yaml") or "sources.yaml").strip(),
            apply_sheets=[str(s).strip() for s in (d.get("apply_sheets") or []) if str(s).strip()],
            copy_attachments=bool(d.get("copy_attachments") if d.get("copy_attachments") is not None else True),
            notes=str(d.get("notes") or "").strip(),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "version": self.version,
            "pack_id": self.pack_id,
            "title": self.title,
            "created_at": self.created_at,
            "reference_case_xlsx": self.reference_case_xlsx,
            "reference_sources_yaml": self.reference_sources_yaml,
            "apply_sheets": self.apply_sheets,
            "copy_attachments": self.copy_attachments,
            "notes": self.notes,
        }


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _sha256_file(p: Path) -> str:
    h = hashlib.sha256()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def _load_yaml(path: Path) -> dict[str, Any]:
    try:
        obj = yaml.safe_load(path.read_text(encoding="utf-8"))
    except Exception:
        obj = None
    return obj if isinstance(obj, dict) else {}


def _merge_sources_yaml(target_sources: Path, pack_sources: Path) -> list[str]:
    """Merge pack sources into case sources.yaml.

    Preserves top-level wrapper keys when present (version/project).
    """
    t = _load_yaml(target_sources)
    p = _load_yaml(pack_sources)

    t_sources = t.get("sources") if isinstance(t.get("sources"), list) else []
    p_sources = p.get("sources") if isinstance(p.get("sources"), list) else []

    existing_ids: set[str] = set()
    for s in t_sources:
        if not isinstance(s, dict):
            continue
        sid = s.get("source_id") or s.get("id")
        if isinstance(sid, str) and sid.strip():
            existing_ids.add(sid.strip())

    added: list[str] = []
    for s in p_sources:
        if not isinstance(s, dict):
            continue
        sid = s.get("source_id") or s.get("id")
        if not isinstance(sid, str) or not sid.strip():
            continue
        if sid.strip() in existing_ids:
            continue
        t_sources.append(s)
        existing_ids.add(sid.strip())
        added.append(sid.strip())

    # Keep wrapper keys from target, but ensure sources is updated.
    t["sources"] = t_sources
    target_sources.write_text(yaml.safe_dump(t, allow_unicode=True, sort_keys=False), encoding="utf-8")
    return added


def export_reference_pack(
    *,
    case_dir: Path,
    out_dir: Path,
    pack_id: str,
    title: str,
    apply_sheets: list[str],
    copy_attachments: bool = True,
) -> Path:
    """Export a reusable reference pack from an existing case folder."""
    case_dir = case_dir.resolve()
    out_dir = out_dir.resolve()

    xlsx = case_dir / "case.xlsx"
    sources = case_dir / "sources.yaml"

    if not xlsx.exists():
        raise FileNotFoundError(f"case.xlsx not found: {xlsx}")
    if not sources.exists():
        raise FileNotFoundError(f"sources.yaml not found: {sources}")

    pack_dir = out_dir / pack_id
    _ensure_dir(pack_dir)

    # Copy reference inputs
    ref_xlsx = pack_dir / "reference_case.xlsx"
    shutil.copy2(xlsx, ref_xlsx)

    ref_sources = pack_dir / "sources.yaml"
    shutil.copy2(sources, ref_sources)

    manifest: dict[str, Any] = {
        "pack_id": pack_id,
        "created_at": _now_iso(),
        "from_case_dir": str(case_dir),
        "files": [],
    }

    # Copy evidence/attachments referenced by ATTACHMENTS sheet.
    if copy_attachments:
        wb = load_workbook(xlsx)
        rows = read_sheet_dicts(wb, "ATTACHMENTS")
        for r in rows:
            fp = str(r.get("file_path") or "").strip()
            if not fp:
                continue
            src = (case_dir / fp).resolve()
            if not src.exists() or not src.is_file():
                continue
            dst = (pack_dir / fp).resolve()
            _ensure_dir(dst.parent)
            if not dst.exists():
                shutil.copy2(src, dst)
            manifest["files"].append(
                {
                    "file_path": fp.replace("\\", "/"),
                    "sha256": _sha256_file(dst),
                }
            )

    pack = PackSpec(
        version=1,
        pack_id=pack_id,
        title=title,
        created_at=_now_iso(),
        reference_case_xlsx=ref_xlsx.name,
        reference_sources_yaml=ref_sources.name,
        apply_sheets=apply_sheets,
        copy_attachments=bool(copy_attachments),
        notes="",
    )

    (pack_dir / "pack.yaml").write_text(
        yaml.safe_dump(pack.to_dict(), allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )
    (pack_dir / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    return pack_dir


def apply_reference_pack(
    *,
    xlsx: Path,
    sources: Path,
    pack_dir: Path,
) -> dict[str, Any]:
    """Apply a reference pack to a case.

    Rules (safe defaults):
    - Copy pack sheets only when the target sheet has no data rows.
    - Merge pack sources into the case sources.yaml (append-only).
    - Copy referenced attachment files into the case dir when missing.
    """
    xlsx = xlsx.resolve()
    sources = sources.resolve()
    case_dir = xlsx.parent.resolve()
    pack_dir = pack_dir.resolve()

    pack_path = pack_dir / "pack.yaml"
    if not pack_path.exists():
        raise FileNotFoundError(f"pack.yaml not found: {pack_path}")

    pack_obj = _load_yaml(pack_path)
    pack = PackSpec.from_dict(pack_obj)
    if not pack.pack_id:
        raise ValueError("pack.yaml missing pack_id")

    ref_xlsx = pack_dir / pack.reference_case_xlsx
    ref_sources = pack_dir / pack.reference_sources_yaml
    if not ref_xlsx.exists():
        raise FileNotFoundError(f"reference_case_xlsx not found: {ref_xlsx}")
    if not ref_sources.exists():
        raise FileNotFoundError(f"reference_sources_yaml not found: {ref_sources}")

    wb_target = load_workbook(xlsx)
    wb_ref = load_workbook(ref_xlsx)

    def _sheet_has_data(wb, sheet_name: str) -> bool:
        if sheet_name not in wb.sheetnames:
            return False
        ws = wb[sheet_name]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if any(v is not None and (not isinstance(v, str) or v.strip()) for v in r):
                return True
        return False

    applied_sheets: list[str] = []
    skipped_sheets: list[str] = []

    for sheet in pack.apply_sheets:
        sheet_name = str(sheet).strip()
        if not sheet_name:
            continue
        if sheet_name not in wb_ref.sheetnames:
            skipped_sheets.append(f"{sheet_name} (missing in pack)")
            continue
        if _sheet_has_data(wb_target, sheet_name):
            skipped_sheets.append(f"{sheet_name} (target not empty)")
            continue

        rows = read_sheet_dicts(wb_ref, sheet_name)
        if not rows:
            skipped_sheets.append(f"{sheet_name} (no rows)")
            continue

        apply_rows_to_sheet(
            wb_target,
            sheet_name=sheet_name,
            rows=rows,
            merge_strategy="REPLACE_SHEET",
            upsert_keys=None,
        )
        applied_sheets.append(sheet_name)

    # Always merge ATTACHMENTS (upsert by evidence_id), because copied sheets may reference evidence_id.
    ref_atts = read_sheet_dicts(wb_ref, "ATTACHMENTS")
    if ref_atts:
        apply_rows_to_sheet(
            wb_target,
            sheet_name="ATTACHMENTS",
            rows=ref_atts,
            merge_strategy="UPSERT_KEYS",
            upsert_keys=["evidence_id"],
        )

    save_workbook(wb_target, xlsx)

    added_sources = _merge_sources_yaml(target_sources=sources, pack_sources=ref_sources)

    copied_files: list[str] = []
    if pack.copy_attachments and ref_atts:
        # Copy only referenced files.
        for r in ref_atts:
            fp = str(r.get("file_path") or "").strip()
            if not fp:
                continue
            src = (pack_dir / fp).resolve()
            if not src.exists() or not src.is_file():
                continue
            dst = (case_dir / fp).resolve()
            _ensure_dir(dst.parent)
            if dst.exists():
                continue
            shutil.copy2(src, dst)
            copied_files.append(fp.replace("\\", "/"))

    return {
        "pack_id": pack.pack_id,
        "applied_sheets": applied_sheets,
        "skipped_sheets": skipped_sheets,
        "added_sources": added_sources,
        "copied_files": copied_files,
    }
