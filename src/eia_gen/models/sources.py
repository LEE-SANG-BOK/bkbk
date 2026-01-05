from __future__ import annotations

from typing import Any

from pydantic import AliasChoices, BaseModel, ConfigDict, Field, model_validator


class SourceEntry(BaseModel):
    model_config = ConfigDict(extra="allow")

    # Accept both v1 (`id`) and v2 (`source_id`) keys.
    id: str = Field(validation_alias=AliasChoices("id", "source_id"))
    # Accept both v1 (`type`) and v2 (`kind`) naming.
    type: str | None = Field(default=None, validation_alias=AliasChoices("type", "kind"))
    title: str | None = None
    publisher: str | None = None
    # v2 commonly uses issued_date / published_date.
    date: str | None = Field(
        default=None,
        validation_alias=AliasChoices("date", "issued_date", "published_date", "published_at", "issued_at"),
    )
    pages: str | None = None
    access_date: str | None = Field(
        default=None, validation_alias=AliasChoices("access_date", "accessed_date", "accessed_at")
    )
    period: str | None = None
    station_name: str | None = None
    station_coord: dict[str, float | None] | None = None
    station_distance_km: float | None = None
    file_path: str | None = Field(default=None, validation_alias=AliasChoices("file_path", "local_file"))
    file_or_url: str | None = Field(default=None, validation_alias=AliasChoices("file_or_url", "url"))
    coverage: str | None = None
    note: str | None = None
    notes: str | None = None
    confidential: bool | None = None


class SourceRegistry(BaseModel):
    # v2 sources.yaml commonly includes top-level metadata (version/project/etc).
    # Keep it in model_extra so round-trips and downstream tooling can access it.
    model_config = ConfigDict(extra="allow")

    sources: list[SourceEntry] = Field(default_factory=list)

    @model_validator(mode="before")
    @classmethod
    def _coerce(cls, v: Any) -> Any:
        # Accept:
        # 1) {"sources": [ ... ]}
        # 2) [ ... ]
        if v is None:
            return {"sources": []}
        if isinstance(v, list):
            return {"sources": v}
        if isinstance(v, dict):
            if "sources" in v and isinstance(v["sources"], list):
                return v
            # sometimes stored as {"entries": [...]}
            if "entries" in v and isinstance(v["entries"], list):
                out = dict(v)
                out["sources"] = out.pop("entries")
                return out
            # try treat dict as single entry?
            return {"sources": [v]}
        raise ValueError("Invalid sources.yaml structure")

    @model_validator(mode="after")
    def _validate_unique_ids(self) -> "SourceRegistry":
        seen: set[str] = set()
        for s in self.sources:
            if s.id in seen:
                raise ValueError(f"duplicate source id: {s.id}")
            seen.add(s.id)
        return self

    def has(self, source_id: str) -> bool:
        return any(s.id == source_id for s in self.sources)

    def get(self, source_id: str) -> SourceEntry | None:
        for s in self.sources:
            if s.id == source_id:
                return s
        return None
