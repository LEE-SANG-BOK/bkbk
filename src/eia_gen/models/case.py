from __future__ import annotations

from enum import Enum
from typing import Any

from pydantic import BaseModel, ConfigDict, Field, model_validator

from eia_gen.models.fields import QuantityField, TextField


class ScopingClass(str, Enum):
    FOCUS = "FOCUS"
    BASELINE = "BASELINE"
    EXCLUDE = "EXCLUDE"


def _normalize_scoping_class(value: str) -> ScopingClass:
    v = (value or "").strip().upper()
    if v in {"FOCUS", "중점", "중점평가", "중점평가항목"}:
        return ScopingClass.FOCUS
    if v in {"BASELINE", "현황", "현황조사", "현황조사항목"}:
        return ScopingClass.BASELINE
    if v in {"EXCLUDE", "제외", "평가제외", "평가제외항목"}:
        return ScopingClass.EXCLUDE
    raise ValueError(f"invalid scoping class: {value!r}")


class Meta(BaseModel):
    model_config = ConfigDict(extra="allow")

    template_version: str = "1.0"
    project_type: str = "관광농원"
    report_type: str = "소규모환경영향평가서"
    language: str = "ko"


class Cover(BaseModel):
    model_config = ConfigDict(extra="allow")

    project_name: TextField = Field(default_factory=TextField)
    client_name: TextField = Field(default_factory=TextField)
    proponent_name: TextField = Field(default_factory=TextField)
    author_org: TextField = Field(default_factory=TextField)
    submit_date: TextField = Field(default_factory=TextField)
    approving_authority: TextField = Field(default_factory=TextField)
    consultation_agency: TextField = Field(default_factory=TextField)


class SummaryInputs(BaseModel):
    model_config = ConfigDict(extra="allow")

    key_issues: list[TextField] = Field(default_factory=list)
    key_measures: list[TextField] = Field(default_factory=list)


class AdminDistrict(BaseModel):
    model_config = ConfigDict(extra="allow")

    sido: TextField = Field(default_factory=TextField)
    sigungu: TextField = Field(default_factory=TextField)
    eupmyeon: TextField = Field(default_factory=TextField)


class CenterCoord(BaseModel):
    model_config = ConfigDict(extra="allow")

    epsg: int = 4326
    lat: QuantityField = Field(default_factory=QuantityField)
    lon: QuantityField = Field(default_factory=QuantityField)


class Location(BaseModel):
    model_config = ConfigDict(extra="allow")

    address: TextField = Field(default_factory=TextField)
    admin: AdminDistrict = Field(default_factory=AdminDistrict)
    center_coord: CenterCoord = Field(default_factory=CenterCoord)


class Parcel(BaseModel):
    model_config = ConfigDict(extra="allow")

    pnu: TextField = Field(default_factory=TextField)
    jibun: TextField = Field(default_factory=TextField)
    land_category: TextField = Field(default_factory=TextField)
    zoning: TextField = Field(default_factory=TextField)
    area_m2: QuantityField = Field(default_factory=QuantityField)
    note: TextField = Field(default_factory=TextField)


class AreaInfo(BaseModel):
    model_config = ConfigDict(extra="allow")

    total_area_m2: QuantityField = Field(default_factory=QuantityField)
    parcels: list[Parcel] = Field(default_factory=list)
    zoning_area_m2: dict[str, QuantityField] = Field(default_factory=dict)


class Facility(BaseModel):
    model_config = ConfigDict(extra="allow")

    category: TextField = Field(default_factory=TextField)
    name: TextField = Field(default_factory=TextField)
    qty: QuantityField = Field(default_factory=QuantityField)
    area_m2: QuantityField = Field(default_factory=QuantityField)
    capacity_person: QuantityField = Field(default_factory=QuantityField)
    note: TextField = Field(default_factory=TextField)


class ContentsScale(BaseModel):
    model_config = ConfigDict(extra="allow")

    facilities: list[Facility] = Field(default_factory=list)
    land_use_plan_summary: dict[str, QuantityField] = Field(default_factory=dict)


class Milestone(BaseModel):
    model_config = ConfigDict(extra="allow")

    phase: TextField = Field(default_factory=TextField)
    start: TextField = Field(default_factory=TextField)  # YYYY-MM
    end: TextField = Field(default_factory=TextField)


class Schedule(BaseModel):
    model_config = ConfigDict(extra="allow")

    milestones: list[Milestone] = Field(default_factory=list)


class PermitItem(BaseModel):
    model_config = ConfigDict(extra="allow")

    name: TextField = Field(default_factory=TextField)
    status: TextField = Field(default_factory=TextField)
    authority: TextField = Field(default_factory=TextField)
    note: TextField = Field(default_factory=TextField)


class LegalPermits(BaseModel):
    model_config = ConfigDict(extra="allow")

    permit_list: list[PermitItem] = Field(default_factory=list)


class ProjectOverview(BaseModel):
    model_config = ConfigDict(extra="allow")

    purpose_need: TextField = Field(default_factory=TextField)
    location: Location = Field(default_factory=Location)
    area: AreaInfo = Field(default_factory=AreaInfo)
    contents_scale: ContentsScale = Field(default_factory=ContentsScale)
    schedule: Schedule = Field(default_factory=Schedule)
    legal_permits: LegalPermits = Field(default_factory=LegalPermits)


class ApplicabilityCalc(BaseModel):
    model_config = ConfigDict(extra="allow")

    판정결론: TextField = Field(default_factory=TextField)
    근거요약: TextField = Field(default_factory=TextField)


class ApplicabilityItem(BaseModel):
    model_config = ConfigDict(extra="allow")

    subject: TextField = Field(default_factory=TextField)  # true/false/unknown
    basis: dict[str, TextField] = Field(default_factory=dict)
    calc: ApplicabilityCalc = Field(default_factory=ApplicabilityCalc)


class PriorAssessmentRef(BaseModel):
    model_config = ConfigDict(extra="allow")

    title: TextField = Field(default_factory=TextField)
    consulted_date: TextField = Field(default_factory=TextField)
    covered_items: list[str] = Field(default_factory=list)
    omission_basis: dict[str, TextField] = Field(default_factory=dict)


class PriorAssessments(BaseModel):
    model_config = ConfigDict(extra="allow")

    strategic_eia: dict[str, Any] = Field(default_factory=dict)


class InfluenceArea(BaseModel):
    model_config = ConfigDict(extra="allow")

    radius_m: QuantityField = Field(default_factory=lambda: QuantityField(v=500.0, u="m"))
    justification: TextField = Field(default_factory=TextField)


class SurveyPlan(BaseModel):
    model_config = ConfigDict(extra="allow")

    influence_area: InfluenceArea = Field(default_factory=InfluenceArea)
    methods: dict[str, TextField] = Field(default_factory=dict)


class ScopingItem(BaseModel):
    model_config = ConfigDict(extra="allow")

    item_id: str
    item_name: str
    category: TextField = Field(default_factory=TextField)  # 중점/현황/제외
    exclude_reason: TextField = Field(default_factory=TextField)
    baseline_method: TextField = Field(default_factory=TextField)
    prediction_method: TextField = Field(default_factory=TextField)
    src_expected: list[str] = Field(default_factory=list)

    @property
    def scoping_class(self) -> ScopingClass:
        return _normalize_scoping_class(self.category.t)


class BaselineTopographyGeology(BaseModel):
    model_config = ConfigDict(extra="allow")

    elevation_range_m: TextField = Field(default_factory=TextField)
    mean_slope_deg: QuantityField = Field(default_factory=QuantityField)
    geology_summary: TextField = Field(default_factory=TextField)
    soil_summary: TextField = Field(default_factory=TextField)


class SpeciesEntry(BaseModel):
    model_config = ConfigDict(extra="allow")

    species_ko: TextField = Field(default_factory=TextField)
    scientific: TextField = Field(default_factory=TextField)
    protected: TextField = Field(default_factory=TextField)
    note: TextField = Field(default_factory=TextField)
    evidence: TextField = Field(default_factory=TextField)


class BaselineEcology(BaseModel):
    model_config = ConfigDict(extra="allow")

    survey_dates: list[TextField] = Field(default_factory=list)
    flora_list: list[SpeciesEntry] = Field(default_factory=list)
    fauna_list: list[SpeciesEntry] = Field(default_factory=list)


class StreamEntry(BaseModel):
    model_config = ConfigDict(extra="allow")

    name: TextField = Field(default_factory=TextField)
    distance_m: QuantityField = Field(default_factory=QuantityField)
    flow_direction: TextField = Field(default_factory=TextField)
    note: TextField = Field(default_factory=TextField)


class BaselineWaterEnvironment(BaseModel):
    model_config = ConfigDict(extra="allow")

    streams: list[StreamEntry] = Field(default_factory=list)
    water_quality: dict[str, Any] = Field(default_factory=dict)


class BaselineAirQuality(BaseModel):
    model_config = ConfigDict(extra="allow")

    station_name: TextField = Field(default_factory=TextField)
    pm10_ugm3: QuantityField = Field(default_factory=QuantityField)
    pm25_ugm3: QuantityField = Field(default_factory=QuantityField)
    ozone_ppm: QuantityField = Field(default_factory=QuantityField)


class NoiseReceptor(BaseModel):
    model_config = ConfigDict(extra="allow")

    name: TextField = Field(default_factory=TextField)
    distance_m: QuantityField = Field(default_factory=QuantityField)
    baseline_day_db: QuantityField = Field(default_factory=QuantityField)
    baseline_night_db: QuantityField = Field(default_factory=QuantityField)
    measured: TextField = Field(default_factory=TextField)  # true/false


class BaselineNoiseVibration(BaseModel):
    model_config = ConfigDict(extra="allow")

    receptors: list[NoiseReceptor] = Field(default_factory=list)


class Viewpoint(BaseModel):
    model_config = ConfigDict(extra="allow")

    vp_id: TextField = Field(default_factory=TextField)
    location_desc: TextField = Field(default_factory=TextField)
    photo_asset_id: TextField = Field(default_factory=TextField)
    note: TextField = Field(default_factory=TextField)


class BaselineLanduseLandscape(BaseModel):
    model_config = ConfigDict(extra="allow")

    current_landcover_summary: TextField = Field(default_factory=TextField)
    protected_areas_overlap: TextField = Field(default_factory=TextField)
    key_viewpoints: list[Viewpoint] = Field(default_factory=list)


class BaselinePopulationTraffic(BaseModel):
    model_config = ConfigDict(extra="allow")

    nearest_village: TextField = Field(default_factory=TextField)
    distance_to_village_m: QuantityField = Field(default_factory=QuantityField)
    access_roads: list[TextField] = Field(default_factory=list)
    expected_vehicles_per_day: QuantityField = Field(default_factory=QuantityField)


class Baseline(BaseModel):
    model_config = ConfigDict(extra="allow")

    topography_geology: BaselineTopographyGeology = Field(default_factory=BaselineTopographyGeology)
    ecology: BaselineEcology = Field(default_factory=BaselineEcology)
    water_environment: BaselineWaterEnvironment = Field(default_factory=BaselineWaterEnvironment)
    air_quality: BaselineAirQuality = Field(default_factory=BaselineAirQuality)
    noise_vibration: BaselineNoiseVibration = Field(default_factory=BaselineNoiseVibration)
    landuse_landscape: BaselineLanduseLandscape = Field(default_factory=BaselineLanduseLandscape)
    population_traffic: BaselinePopulationTraffic = Field(default_factory=BaselinePopulationTraffic)


class ImpactPrediction(BaseModel):
    model_config = ConfigDict(extra="allow")

    construction: dict[str, Any] = Field(default_factory=dict)
    operation: dict[str, Any] = Field(default_factory=dict)


class MitigationMeasure(BaseModel):
    model_config = ConfigDict(extra="allow")

    measure_id: str
    title: TextField = Field(default_factory=TextField)
    phase: TextField = Field(default_factory=TextField)  # 공사/운영
    description: TextField = Field(default_factory=TextField)
    location_ref: TextField = Field(default_factory=TextField)
    design_params: TextField = Field(default_factory=TextField)
    monitoring: TextField = Field(default_factory=TextField)
    related_impacts: list[str] = Field(default_factory=list)


class Mitigation(BaseModel):
    model_config = ConfigDict(extra="allow")

    measures: list[MitigationMeasure] = Field(default_factory=list)


class ConditionTrackerItem(BaseModel):
    model_config = ConfigDict(extra="allow")

    item: TextField = Field(default_factory=TextField)
    measure_id: TextField = Field(default_factory=TextField)
    when: TextField = Field(default_factory=TextField)
    evidence: TextField = Field(default_factory=TextField)
    responsible: TextField = Field(default_factory=TextField)


class ManagementPlan(BaseModel):
    model_config = ConfigDict(extra="allow")

    implementation_register: list[ConditionTrackerItem] = Field(default_factory=list)


class ResidentOpinion(BaseModel):
    model_config = ConfigDict(extra="allow")

    applicable: TextField = Field(default_factory=TextField)  # true/false
    summary: TextField = Field(default_factory=TextField)
    responses: list[TextField] = Field(default_factory=list)


class Asset(BaseModel):
    model_config = ConfigDict(extra="allow")

    asset_id: str
    type: str  # location_map / landuse_plan / layout_plan / drainage_map / photo / simulation
    file_path: str
    caption: TextField = Field(default_factory=TextField)
    source_ids: list[str] = Field(default_factory=list)
    shooting_date: TextField = Field(default_factory=TextField)
    viewpoint: TextField = Field(default_factory=TextField)


class Case(BaseModel):
    model_config = ConfigDict(extra="allow")

    meta: Meta = Field(default_factory=Meta)
    cover: Cover = Field(default_factory=Cover)
    summary_inputs: SummaryInputs = Field(default_factory=SummaryInputs)
    project_overview: ProjectOverview = Field(default_factory=ProjectOverview)

    applicability: dict[str, Any] = Field(default_factory=dict)
    prior_assessments: dict[str, Any] = Field(default_factory=dict)

    survey_plan: SurveyPlan = Field(default_factory=SurveyPlan)
    scoping_matrix: list[ScopingItem] = Field(default_factory=list)
    baseline: Baseline = Field(default_factory=Baseline)
    impact_prediction: ImpactPrediction = Field(default_factory=ImpactPrediction)
    mitigation: Mitigation = Field(default_factory=Mitigation)
    management_plan: ManagementPlan = Field(default_factory=ManagementPlan)
    resident_opinion: ResidentOpinion = Field(default_factory=ResidentOpinion)
    assets: list[Asset] = Field(default_factory=list)

    @model_validator(mode="after")
    def _validate_scoping(self) -> "Case":
        seen: set[str] = set()
        for item in self.scoping_matrix:
            if item.item_id in seen:
                raise ValueError(f"duplicate scoping item_id: {item.item_id}")
            seen.add(item.item_id)
            if item.scoping_class == ScopingClass.EXCLUDE and item.exclude_reason.is_empty():
                raise ValueError(
                    f"scoping item {item.item_id} is EXCLUDE but exclude_reason is empty"
                )
        return self

