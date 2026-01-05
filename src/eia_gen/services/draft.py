from __future__ import annotations

from typing import Any

from pydantic import BaseModel, ConfigDict, Field


class TableDraft(BaseModel):
    model_config = ConfigDict(extra="ignore")

    table_id: str | None = None
    caption: str
    headers: list[str]
    rows: list[list[str]]
    source_ids: list[str] = Field(default_factory=list)


class FigureDraft(BaseModel):
    model_config = ConfigDict(extra="ignore")

    figure_id: str | None = None
    file_path: str | None = None
    caption: str
    source_ids: list[str] = Field(default_factory=list)


class SectionDraft(BaseModel):
    model_config = ConfigDict(extra="ignore")

    section_id: str
    title: str
    paragraphs: list[str] = Field(default_factory=list)
    tables: list[TableDraft] = Field(default_factory=list)
    figures: list[FigureDraft] = Field(default_factory=list)
    todos: list[str] = Field(default_factory=list)
    meta: dict[str, Any] = Field(default_factory=dict)


class ReportDraft(BaseModel):
    model_config = ConfigDict(extra="ignore")

    sections: list[SectionDraft] = Field(default_factory=list)
