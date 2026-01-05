from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class TableData:
    caption: str
    headers: list[str]
    rows: list[list[str]]
    source_ids: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class FigureData:
    file_path: str | None
    caption: str
    source_ids: list[str] = field(default_factory=list)
    # Optional v2 figure controls (from `case.xlsx` FIGURES sheet)
    width_mm: float | None = None
    crop: str | None = None
    gen_method: str | None = None
