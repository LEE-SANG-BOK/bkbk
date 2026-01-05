from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field


class DocNumbering(BaseModel):
    model_config = ConfigDict(extra="ignore")

    chapter_style: str = "제{n}장"
    # NOTE: `label` is always passed during formatting.
    # - If a spec provides `label` (e.g. "2.3.2"), it will be used.
    # - Otherwise we auto-generate label (e.g. "1-1").
    table_style: str = "표 {label}"
    figure_style: str = "그림 {label}"


class DocProfile(BaseModel):
    model_config = ConfigDict(extra="ignore")

    report_type: str = "소규모환경영향평가서"
    project_type: str = "관광농원"
    language: str = "ko"
    numbering: DocNumbering = Field(default_factory=DocNumbering)
    forbidden_phrases: list[str] = Field(default_factory=list)


class ScopingRules(BaseModel):
    model_config = ConfigDict(extra="ignore")

    categories: list[str] = Field(default_factory=lambda: ["중점", "현황", "제외"])
    require_exclude_reason_if_excluded: bool = True


class PriorOmissionRules(BaseModel):
    model_config = ConfigDict(extra="ignore")

    enable: bool = True
    basis: str = "시행령 제60조"
    rule_ref_source_id: str | None = None


class SmallScaleRules(BaseModel):
    model_config = ConfigDict(extra="ignore")

    default_no_manager_no_post_monitoring: bool = True
    conditional_sections: dict[str, str] = Field(default_factory=dict)


class GlobalRules(BaseModel):
    model_config = ConfigDict(extra="ignore")

    scoping: ScopingRules = Field(default_factory=ScopingRules)
    prior_assessment_omission: PriorOmissionRules = Field(default_factory=PriorOmissionRules)
    small_scale_special: SmallScaleRules = Field(default_factory=SmallScaleRules)


class SectionOutputs(BaseModel):
    model_config = ConfigDict(extra="ignore")

    tables: list[str] = Field(default_factory=list)
    figures: list[str] = Field(default_factory=list)


class SectionQA(BaseModel):
    model_config = ConfigDict(extra="ignore")

    error_if_missing: list[str] = Field(default_factory=list)
    warn_if_missing: list[str] = Field(default_factory=list)


SectionMode = Literal["deterministic", "llm", "hybrid"]


class Section(BaseModel):
    model_config = ConfigDict(extra="ignore")

    id: str
    heading: str
    anchor: str | None = None
    mode: SectionMode = "hybrid"
    input_paths: list[str] = Field(default_factory=list)
    outputs: SectionOutputs = Field(default_factory=SectionOutputs)
    condition: str | None = None
    qa: SectionQA = Field(default_factory=SectionQA)
    rules: dict[str, Any] = Field(default_factory=dict)


class SectionsSpec(BaseModel):
    model_config = ConfigDict(extra="ignore")

    version: str = "1.0"
    doc_profile: DocProfile = Field(default_factory=DocProfile)
    global_rules: GlobalRules = Field(default_factory=GlobalRules)
    sections: list[Section] = Field(default_factory=list)


class TableColumn(BaseModel):
    model_config = ConfigDict(extra="allow")

    key: str
    title: str
    type: str
    path: str | None = None
    unit_path: str | None = None
    enum: list[str] | None = None
    allow_empty: bool | None = None


class TableValidation(BaseModel):
    model_config = ConfigDict(extra="allow")

    id: str
    type: str
    severity: Literal["ERROR", "WARN", "INFO"] = "ERROR"


class TableSpec(BaseModel):
    model_config = ConfigDict(extra="ignore")

    id: str
    anchor: str
    caption: str
    chapter: int
    # Optional fixed label (e.g. "2.3.2") to match a sample document numbering scheme.
    label: str | None = None
    # When set, overrides global defaults.include_src_column for this table.
    include_src_column: bool | None = None
    # Static tables: embed headers/rows directly in the spec.
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)
    source_ids: list[str] = Field(default_factory=list)
    data_path: str | None = None
    mode: str | None = None
    columns: list[TableColumn] = Field(default_factory=list)
    validations: list[TableValidation] = Field(default_factory=list)
    rows_definition: list[dict[str, Any]] = Field(default_factory=list)


class TableDefaults(BaseModel):
    model_config = ConfigDict(extra="ignore")

    include_src_column: bool = True
    src_render: str = "join"
    empty_cell: str = "[작성자 기입 필요]"


class TableSpecs(BaseModel):
    model_config = ConfigDict(extra="ignore")

    version: str = "1.0"
    defaults: TableDefaults = Field(default_factory=TableDefaults)
    tables: list[TableSpec] = Field(default_factory=list)


class RequiredIf(BaseModel):
    model_config = ConfigDict(extra="ignore")

    scoping_item_id: str | None = None
    category_in: list[str] = Field(default_factory=list)


class FigureSpec(BaseModel):
    model_config = ConfigDict(extra="ignore")

    id: str
    anchor: str
    caption: str
    chapter: int
    # Optional fixed label (e.g. "7.4.2") to match a sample document numbering scheme.
    label: str | None = None
    required: bool = False
    required_if: RequiredIf | None = None
    asset_type: str
    constraints: dict[str, Any] = Field(default_factory=dict)
    must_include: list[str] = Field(default_factory=list)


class FigureDefaults(BaseModel):
    model_config = ConfigDict(extra="ignore")

    require_source_ids: bool = True
    empty_caption: str = "[작성자 캡션 기입]"


class FigureSpecs(BaseModel):
    model_config = ConfigDict(extra="ignore")

    version: str = "1.0"
    defaults: FigureDefaults = Field(default_factory=FigureDefaults)
    figures: list[FigureSpec] = Field(default_factory=list)


class TemplateInsert(BaseModel):
    model_config = ConfigDict(extra="ignore")

    type: Literal["section", "table", "figure"]
    id: str
    conditional: bool = False


class TemplateAnchor(BaseModel):
    model_config = ConfigDict(extra="ignore")

    anchor: str
    insert: TemplateInsert


class TemplateMap(BaseModel):
    model_config = ConfigDict(extra="ignore")

    version: str = "1.0"
    template_file: str
    anchors: list[TemplateAnchor] = Field(default_factory=list)


class SpecBundle(BaseModel):
    model_config = ConfigDict(extra="ignore")

    sections: SectionsSpec
    tables: TableSpecs
    figures: FigureSpecs
    template_map: TemplateMap
