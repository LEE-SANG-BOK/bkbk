from __future__ import annotations

from typing import Any

from pydantic import BaseModel, ConfigDict, Field, model_validator


class TextField(BaseModel):
    """Text value with optional source IDs and confidentiality flag.

    Accepts either:
    - plain string
    - dict: {"t": "...", "src": [...], "note": "...", "confidential": true}
    """

    model_config = ConfigDict(extra="ignore")

    t: str = ""
    src: list[str] = Field(default_factory=list)
    note: str | None = None
    confidential: bool | None = None

    @model_validator(mode="before")
    @classmethod
    def _coerce(cls, v: Any) -> Any:
        if v is None:
            return {"t": ""}
        if isinstance(v, str):
            return {"t": v}
        if isinstance(v, dict):
            # allow either "t" or "text"
            if "t" not in v and "text" in v:
                v = {**v, "t": v.get("text", "")}
            return v
        return {"t": str(v)}

    def is_empty(self) -> bool:
        return self.t.strip() == ""

    def text_or_placeholder(self, placeholder: str = "【작성자 기입 필요】") -> str:
        return self.t.strip() if self.t.strip() else placeholder


class QuantityField(BaseModel):
    """Numeric value with unit and optional source IDs.

    Accepts either:
    - plain int/float
    - dict: {"v": 123, "u": "m2", "src": [...], "note": "..."}
    """

    model_config = ConfigDict(extra="ignore")

    v: float | None = None
    u: str | None = None
    src: list[str] = Field(default_factory=list)
    note: str | None = None

    @model_validator(mode="before")
    @classmethod
    def _coerce(cls, v: Any) -> Any:
        if v is None:
            return {"v": None}
        if isinstance(v, (int, float)):
            return {"v": float(v)}
        if isinstance(v, dict):
            # allow either "v" or "value"
            if "v" not in v and "value" in v:
                v = {**v, "v": v.get("value")}
            return v
        # best effort
        try:
            return {"v": float(v)}
        except Exception:
            return {"v": None, "note": f"non-numeric: {v!r}"}

    def is_empty(self) -> bool:
        return self.v is None


def normalize_source_ids(ids: list[str] | None) -> list[str]:
    if not ids:
        return []
    return [s.strip() for s in ids if isinstance(s, str) and s.strip()]

