from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Any


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


@dataclass(frozen=True)
class DataRequest:
    req_id: str
    enabled: bool
    priority: int
    connector: str
    purpose: str
    src_id: str
    params_json: str
    params: dict[str, Any]
    output_sheet: str
    merge_strategy: str
    upsert_keys: list[str]
    run_mode: str
    last_run_at: str
    last_evidence_ids: list[str]
    note: str


@dataclass(frozen=True)
class Evidence:
    evidence_id: str
    evidence_type: str
    title: str
    file_path: str
    used_in: str
    data_origin: str
    src_id: str
    note: str = ""

