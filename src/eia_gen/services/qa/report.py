from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field


Severity = Literal["ERROR", "WARN", "INFO"]


class RuleResult(BaseModel):
    model_config = ConfigDict(extra="ignore")

    rule_id: str
    severity: Severity
    message: str
    fix_hint: str | None = None
    path: str | None = None
    # Optional linkage metadata for exporting VALIDATION_SUMMARY in source_register.xlsx.
    related_anchor: str | None = None
    related_sheet: str | None = None
    related_row_id: str | None = None


class ValidationReport(BaseModel):
    model_config = ConfigDict(extra="ignore")

    results: list[RuleResult] = Field(default_factory=list)
    stats: dict[str, int] = Field(default_factory=dict)

    @property
    def failures(self) -> list[RuleResult]:
        return [r for r in self.results if r.severity == "ERROR"]

    @property
    def warnings(self) -> list[RuleResult]:
        return [r for r in self.results if r.severity == "WARN"]
