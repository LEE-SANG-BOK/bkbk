#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any


try:
    import yaml
except Exception as e:  # pragma: no cover
    raise SystemExit(
        "Missing python deps. Install project deps first (see eia-gen/pyproject.toml).\n"
        f"Import error: {e}"
    )


@dataclass(frozen=True)
class SheetHint:
    sheet: str
    columns: list[str]
    reason: str


def _load_validation_report(path: Path) -> dict[str, Any]:
    obj = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(obj, dict) or "results" not in obj or not isinstance(obj["results"], list):
        raise SystemExit("validation_report JSON must be an object with a top-level 'results' list")
    return obj


def _load_sections_yaml(path: Path) -> dict[str, dict[str, Any]]:
    y = yaml.safe_load(path.read_text(encoding="utf-8"))
    sections = y.get("sections") or []
    out: dict[str, dict[str, Any]] = {}
    for s in sections:
        sid = str(s.get("id") or "").strip()
        if not sid:
            continue
        out[sid] = {
            "heading": str(s.get("heading") or ""),
            "input_paths": list(s.get("input_paths") or []),
            "outputs_tables": list(((s.get("outputs") or {}).get("tables") or [])),
            "outputs_figures": list(((s.get("outputs") or {}).get("figures") or [])),
        }
    return out


def _load_table_specs_yaml(path: Path) -> dict[str, dict[str, Any]]:
    y = yaml.safe_load(path.read_text(encoding="utf-8"))
    tables = y.get("tables") or []
    out: dict[str, dict[str, Any]] = {}
    for t in tables:
        tid = str(t.get("id") or "").strip()
        if not tid:
            continue
        out[tid] = t
    return out


def _guess_report_kind(results: list[dict[str, Any]]) -> str:
    # If any path looks like DIA section id, treat as DIA.
    for r in results:
        p = str(r.get("path") or "")
        if p.startswith("DIA"):
            return "DIA"
    return "EIA"


def _guess_report_kind_with_fallback(report_path: Path, results: list[dict[str, Any]]) -> str:
    kind = _guess_report_kind(results)
    if kind == "DIA":
        return "DIA"

    # Fallback: empty/clean reports have no section ids in results; infer from filename.
    name = report_path.name.lower()
    if "validation_report_dia" in name or name.endswith("_dia.json") or name.startswith("dia"):
        return "DIA"
    if "validation_report_eia" in name or name.endswith("_eia.json") or name.startswith("eia"):
        return "EIA"

    return "EIA"


# NOTE: This mapping is intentionally conservative.
# Source-of-truth references:
# - Sheet/column list: src/eia_gen/services/xlsx/case_template_v2.py
# - Best-effort mapping: docs/10_case_xlsx_v2_mapping.md
_PATH_HINTS: list[tuple[str, list[SheetHint]]] = [
    (
        "cover",
        [
            SheetHint(
                sheet="PROJECT",
                columns=["project_name", "submit_date", "approving_authority", "proponent_name", "client_name", "src_id"],
                reason="표지/커버 필드(프로젝트명/제출일/기관 등)",
            ),
        ],
    ),
    (
        "project_overview.location",
        [
            SheetHint(
                sheet="LOCATION",
                columns=[
                    "address_road",
                    "address_jibeon",
                    "admin_si",
                    "admin_eupmyeon",
                    "center_lat",
                    "center_lon",
                    "boundary_file",
                    "src_id",
                ],
                reason="위치/좌표/경계(주소+중심점+경계파일)",
            )
        ],
    ),
    (
        "project_overview.area",
        [
            SheetHint(sheet="PROJECT", columns=["total_area_m2"], reason="총면적(대표값)"),
            SheetHint(
                sheet="PARCELS",
                columns=["parcel_no", "jimok", "zoning", "area_m2", "note", "src_id"],
                reason="지번별 현황(지목/용도지역/면적) + 근거(src_id)",
            ),
        ],
    ),
    (
        "project_overview.contents_scale",
        [
            SheetHint(
                sheet="FACILITIES",
                columns=["type", "name", "qty", "qty_unit", "area_m2", "capacity_person", "note", "src_id"],
                reason="시설/규모(수량/면적/수용인원/비고) + 근거(src_id)",
            ),
            SheetHint(
                sheet="PROJECT",
                columns=["main_facilities_summary", "max_occupancy", "expected_visitors_day"],
                reason="요약용 대표 값(보고서 텍스트/표)",
            ),
        ],
    ),
    (
        "project_overview.schedule",
        [
            SheetHint(sheet="SCHEDULE", columns=["sch_id", "phase", "start_date", "end_date", "src_id"], reason="공정/기간"),
        ],
    ),
    (
        "survey_plan",
        [
            SheetHint(
                sheet="FIELD_SURVEY_LOG",
                columns=["survey_id", "survey_date", "survey_time_range", "surveyors", "weather", "scope", "route_desc", "src_id"],
                reason="현장조사 범위/방법/로그(조사계획 근거)",
            )
        ],
    ),
    (
        "summary_inputs",
        [
            SheetHint(
                sheet="ENV_MITIGATION",
                columns=["target_item", "measure", "src_id"],
                reason="요약(key_issues/key_measures) best-effort 파생(ENV_MITIGATION 기반)",
            )
        ],
    ),
    (
        "scoping_matrix",
        [
            SheetHint(
                sheet="ENV_SCOPING",
                columns=["item_id", "category", "status", "reason", "src_id"],
                reason="스코핑 매트릭스(중점/현황/제외)",
            )
        ],
    ),
    (
        "mitigation.measures",
        [
            SheetHint(
                sheet="ENV_MITIGATION",
                columns=["mit_id", "stage", "target_item", "measure", "src_id"],
                reason="저감대책(요약/본문/부록)",
            )
        ],
    ),
    (
        "baseline.air_quality",
        [
            SheetHint(
                sheet="ENV_BASE_AIR",
                columns=["air_id", "station_name", "period_start", "period_end", "pollutant", "value_avg", "src_id"],
                reason="대기질(측정소/기간/항목)",
            )
        ],
    ),
    (
        "baseline.water_environment",
        [
            SheetHint(
                sheet="ENV_BASE_WATER",
                columns=["water_id", "waterbody_name", "parameter", "value", "src_id"],
                reason="수환경(수계/수질)",
            )
        ],
    ),
    (
        "baseline.noise_vibration",
        [
            SheetHint(
                sheet="ENV_BASE_NOISE",
                columns=["noise_id", "point_name", "receptor_type", "day_leq", "night_leq", "src_id"],
                reason="소음·진동(대표 지점)",
            )
        ],
    ),
    (
        "baseline.topography_geology",
        [
            SheetHint(
                sheet="ENV_BASE_GEO",
                columns=["geo_id", "topic", "summary", "source_map", "src_id"],
                reason="지형·지질/토양 요약(주제별 요약+근거)",
            )
        ],
    ),
    (
        "baseline.population_traffic",
        [
            SheetHint(
                sheet="ENV_BASE_SOCIO",
                columns=["socio_id", "admin_code", "admin_name", "year", "population_total", "households", "housing_total", "src_id"],
                reason="인구/주거(행정구역+연도별 통계)",
            )
        ],
    ),
    (
        "baseline.ecology",
        [
            SheetHint(
                sheet="ENV_ECO_EVENTS",
                columns=["eco_event_id", "date", "season", "survey_team", "weather", "area_desc", "evidence_id"],
                reason="생태 조사 이벤트/일정(간접 근거)",
            ),
            SheetHint(
                sheet="ENV_ECO_OBS",
                columns=["obs_id", "eco_event_id", "taxon_group", "scientific_name", "korean_name", "count", "protected_status", "src_id"],
                reason="생태 관찰 기록(식물/동물)",
            ),
        ],
    ),
    (
        "baseline.landuse_landscape",
        [
            SheetHint(
                sheet="ENV_LANDSCAPE",
                columns=["view_id", "viewpoint_name", "description", "photo_fig_id", "src_id"],
                reason="경관/조망점(조망점/사진 연결)",
            )
        ],
    ),
    (
        "assets",
        [
            SheetHint(sheet="FIGURES", columns=["fig_id", "file_path", "width_mm", "src_id"], reason="그림 입력/삽입 위치"),
            SheetHint(sheet="ATTACHMENTS", columns=["evidence_id", "file_path", "related_fig_id", "src_id"], reason="첨부/증빙 연결"),
            SheetHint(sheet="SSOT_PAGE_OVERRIDES", columns=["sample_page", "override_file_path", "override_page"], reason="샘플 페이지 치환"),
        ],
    ),
    (
        "management_plan",
        [
            SheetHint(
                sheet="ENV_MANAGEMENT",
                columns=["cond_id", "condition_text", "compliance_plan", "evidence_id", "status", "note"],
                reason="협의의견(조건) 이행관리(행별 1조건)",
            ),
            SheetHint(
                sheet="ENV_MITIGATION",
                columns=["mit_id", "stage", "target_item", "measure", "src_id"],
                reason="연계 대책(대책ID/단계/내용) 연결(선택)",
            ),
        ],
    ),
    # DIA
    (
        "disaster.target_area",
        [
            SheetHint(
                sheet="DRR_TARGET_AREA",
                columns=["concept", "upstream_area_km2", "downstream_to", "affected_neighborhood", "map_fig_id", "data_origin", "src_id"],
                reason="평가대상지역 개념/범위",
            )
        ],
    ),
    (
        "disaster.target_area_parts",
        [
            SheetHint(
                sheet="DRR_TARGET_AREA_PARTS",
                columns=["part", "included", "reason", "exclude_reason", "geom_ref", "figure_id", "data_origin", "src_id"],
                reason="대상지역 4파트(포함/제외 사유)",
            ),
        ],
    ),
    (
        "disaster.scoping_matrix",
        [
            SheetHint(
                sheet="DRR_SCOPING",
                columns=["hazard_type", "include", "reason", "method", "data_origin", "src_id"],
                reason="재해 스코핑(대상/제외 + 사유/방법)",
            ),
        ],
    ),
    (
        "disaster.rainfall",
        [
            SheetHint(
                sheet="DRR_HYDRO_RAIN",
                columns=[
                    "source_basis",
                    "return_period_yr",
                    "duration_hr",
                    "rainfall_mm",
                    "intensity_formula",
                    "temporal_dist",
                    "data_origin",
                    "evidence_id",
                    "src_id",
                ],
                reason="강우 입력(지속시간/빈도/강우량)",
            ),
        ],
    ),
    (
        "disaster.hazard_history",
        [
            SheetHint(
                sheet="DRR_BASE_HAZARD",
                columns=["hazard_type", "occurred", "interview_done", "interview_summary", "photo_fig_id", "evidence_id", "data_origin", "src_id"],
                reason="재해발생 현황(탐문/현장사진/자료조사)",
            ),
        ],
    ),
    (
        "disaster.interviews",
        [
            SheetHint(
                sheet="DRR_INTERVIEWS",
                columns=["respondent_id", "residence_years", "location_desc", "summary", "photo_fig_id", "evidence_id", "data_origin", "src_id"],
                reason="주민탐문(익명) 요약",
            ),
        ],
    ),
    (
        "disaster.runoff_basins",
        [
            SheetHint(
                sheet="DRR_HYDRO_MODEL",
                columns=[
                    "hydro_id",
                    "model",
                    "cn_or_c",
                    "tc_min",
                    "k_storage",
                    "peak_cms_before",
                    "peak_cms_after",
                    "vol_m3_before",
                    "vol_m3_after",
                    "critical_duration_hr",
                    "data_origin",
                    "evidence_id",
                    "src_id",
                ],
                reason="유역/유출 해석 입력(사업 전/후 비교)",
            ),
        ],
    ),
    (
        "disaster.sediment_erosion",
        [
            SheetHint(
                sheet="DRR_SEDIMENT",
                columns=[
                    "sed_id",
                    "method",
                    "r_factor",
                    "k_factor",
                    "ls_factor",
                    "c_factor",
                    "p_factor",
                    "soil_loss_t_ha_yr_before",
                    "soil_loss_t_ha_yr_after",
                    "data_origin",
                    "evidence_id",
                    "src_id",
                ],
                reason="토사유출/침식(USLE/RUSLE) 입력",
            ),
        ],
    ),
    (
        "disaster.slope_landslide",
        [
            SheetHint(
                sheet="DRR_SLOPE",
                columns=["slope_id", "has_slope_work", "height_m", "risk_grade", "mitigation_ref", "exclude", "exclude_reason", "data_origin", "evidence_id"],
                reason="사면/산사태 위험요소(해당 시)",
            ),
        ],
    ),
    (
        "disaster.drainage_facilities",
        [
            SheetHint(
                sheet="UTILITIES",
                columns=["util_type", "capacity", "discharge_point", "drawing_ref", "src_id"],
                reason="배수/우수 시설(도면/수량집계표 기반) → DIA 배수시설 표로 연결",
            ),
        ],
    ),
    (
        "disaster.measures",
        [
            SheetHint(
                sheet="DRR_MITIGATION",
                columns=["hazard_type", "description", "design_ref", "maintenance_ref", "data_origin", "src_id"],
                reason="재해 저감대책(공사/운영 단계)",
            ),
        ],
    ),
    (
        "disaster.maintenance_ledger",
        [
            SheetHint(
                sheet="DRR_MAINTENANCE",
                columns=["facility_name", "inspection_cycle", "maintenance_method", "responsible", "ledger_template", "evidence_id", "src_id"],
                reason="유지관리대장(부록)",
            ),
        ],
    ),
]

# Some QA rules report a sheet name (e.g., ENV_BASE_WATER) as `path` instead of a section id.
# Provide a best-effort alias so next-actions remains useful for "sheet-level" blockers.
_SHEET_PATH_ALIASES: dict[str, str] = {
    # EIA
    "ENV_SCOPING": "scoping_matrix",
    "ENV_MANAGEMENT": "management_plan",
    "ENV_BASE_AIR": "baseline.air_quality",
    "ENV_BASE_WATER": "baseline.water_environment",
    "ENV_BASE_NOISE": "baseline.noise_vibration",
    "ENV_BASE_GEO": "baseline.topography_geology",
    "ENV_BASE_SOCIO": "baseline.population_traffic",
    "ENV_ECO_EVENTS": "baseline.ecology",
    "ENV_ECO_OBS": "baseline.ecology",
    "ENV_LANDSCAPE": "baseline.landuse_landscape",
    # DIA (best-effort)
    "DRR_HYDRO_RAIN": "disaster.rainfall",
    "DRR_SCOPING": "disaster.scoping_matrix",
    "DRR_INTERVIEWS": "disaster.interviews",
    "DRR_HYDRO_MODEL": "disaster.runoff_basins",
    "DRR_SEDIMENT": "disaster.sediment_erosion",
    "DRR_SLOPE": "disaster.slope_landslide",
    "DRR_MITIGATION": "disaster.measures",
    "DRR_MAINTENANCE": "disaster.maintenance_ledger",
}


def _hints_for_input_path(input_path: str) -> list[SheetHint]:
    p = (input_path or "").strip()
    if not p:
        return []

    # Special-case: 'baseline' is intentionally too broad; match only exact.
    if p == "baseline":
        return [
            SheetHint(sheet="ENV_BASE_GEO", columns=["topic", "summary", "src_id"], reason="지형·지질/토양 요약"),
            SheetHint(sheet="ENV_BASE_AIR", columns=["station_name", "pollutant", "value_avg", "src_id"], reason="대기질(대표값)"),
            SheetHint(sheet="ENV_BASE_WATER", columns=["waterbody_name", "parameter", "value", "src_id"], reason="수환경(대표값)"),
            SheetHint(sheet="ENV_BASE_NOISE", columns=["point_name", "day_leq", "night_leq", "src_id"], reason="소음·진동(대표값)"),
            SheetHint(sheet="ENV_BASE_SOCIO", columns=["admin_name", "year", "population_total", "src_id"], reason="인구/주거(대표값)"),
        ]

    out: list[SheetHint] = []
    for prefix, hints in _PATH_HINTS:
        if p == prefix or p.startswith(prefix + "."):
            out.extend(hints)
    return out


def _extra_hints_from_message(msg: str) -> list[SheetHint]:
    m = (msg or "")
    hints: list[SheetHint] = []
    if "key_issues" in m:
        hints.append(SheetHint(sheet="ENV_MITIGATION", columns=["target_item"], reason="요약 핵심 이슈(key_issues)"))
    if "key_measures" in m:
        hints.append(SheetHint(sheet="ENV_MITIGATION", columns=["measure"], reason="요약 핵심 저감대책(key_measures)"))
    return hints


def build_next_actions(
    *,
    report_kind: str,
    sections: dict[str, dict[str, Any]],
    table_specs: dict[str, dict[str, Any]] | None,
    results: list[dict[str, Any]],
) -> dict[str, Any]:
    def _sev_rank(sev: str) -> int:
        s = (sev or "").strip().upper()
        return {"ERROR": 3, "WARN": 2, "INFO": 1}.get(s, 0)

    # Multiple QA rules can emit overlapping items for the same `path` (e.g., placeholder + missing sheets).
    # For usability, aggregate by path and union their sheet hints.
    order: list[str] = []
    items_by_key: dict[str, dict[str, Any]] = {}

    for r in results:
        path = str(r.get("path") or "").strip()
        rule_id = str(r.get("rule_id") or "").strip()
        severity = str(r.get("severity") or "").strip()
        message = str(r.get("message") or "").strip()

        section = sections.get(path)
        table_spec = table_specs.get(path) if table_specs else None
        input_paths = list(section.get("input_paths") or []) if section else []

        hints: list[SheetHint] = []
        if path == "DATA_REQUESTS" or rule_id.startswith("DATA_REQUESTS"):
            hints.append(
                SheetHint(
                    sheet="DATA_REQUESTS",
                    columns=["req_id", "enabled", "connector", "params_json", "output_sheet", "run_mode", "note"],
                    reason="DATA_REQUESTS 활성화/파라미터/실행 설정",
                )
            )

        if (not section) and path in _SHEET_PATH_ALIASES:
            hints.extend(_hints_for_input_path(_SHEET_PATH_ALIASES[path]))

        # Table-level findings (e.g., W-TBL-PLACEHOLDER-001 path=TBL-...).
        # Provide best-effort hints from the table's data_path/rows_definition.
        if (not section) and isinstance(table_spec, dict):
            data_path = str(table_spec.get("data_path") or "").strip()
            if data_path:
                hints.extend(_hints_for_input_path(data_path))
            rows_def = table_spec.get("rows_definition") or []
            if isinstance(rows_def, list):
                for group in rows_def:
                    if not isinstance(group, dict):
                        continue
                    items = group.get("items") or []
                    if not isinstance(items, list):
                        continue
                    for item in items:
                        if not isinstance(item, dict):
                            continue
                        ip = str(item.get("path") or "").strip()
                        if ip:
                            hints.extend(_hints_for_input_path(ip))

        if rule_id == "W-SRC-TBD-001" or "S-TBD" in message or "SRC-TBD" in message:
            hints.append(
                SheetHint(
                    sheet="sources.yaml",
                    columns=["id", "type", "title", "date", "url", "local_file"],
                    reason="임시 출처(S-TBD/SRC-TBD) → 실제 출처로 교체",
                )
            )

        for ip in input_paths:
            hints.extend(_hints_for_input_path(str(ip)))
        hints.extend(_extra_hints_from_message(message))

        # Fallback: if the section has deterministic tables but no (or unmapped) input_paths,
        # derive hints from table specs. This keeps the mapping best-effort without mutating SSOT specs.
        if not hints and section and table_specs:
            for table_id in section.get("outputs_tables") or []:
                t = table_specs.get(str(table_id))
                if not isinstance(t, dict):
                    continue
                data_path = str(t.get("data_path") or "").strip()
                if data_path:
                    hints.extend(_hints_for_input_path(data_path))
                rows_def = t.get("rows_definition") or []
                if isinstance(rows_def, list):
                    for group in rows_def:
                        if not isinstance(group, dict):
                            continue
                        items = group.get("items") or []
                        if not isinstance(items, list):
                            continue
                        for item in items:
                            if not isinstance(item, dict):
                                continue
                            ip = str(item.get("path") or "").strip()
                            if ip:
                                hints.extend(_hints_for_input_path(ip))

        # De-dup by (sheet, columns)
        seen: set[tuple[str, tuple[str, ...]]] = set()
        hints2: list[dict[str, Any]] = []
        for h in hints:
            key = (h.sheet, tuple(h.columns))
            if key in seen:
                continue
            seen.add(key)
            hints2.append({"sheet": h.sheet, "columns": h.columns, "reason": h.reason})

        if not hints2:
            continue

        key = path or f"{rule_id}:{message}"
        if key not in items_by_key:
            heading = section.get("heading") if section else ""
            if (not heading) and isinstance(table_spec, dict):
                heading = str(table_spec.get("caption") or "").strip()

            items_by_key[key] = {
                "severity": severity,
                "rule_ids": [rule_id] if rule_id else [],
                "rule_id_set": {rule_id} if rule_id else set(),
                "messages": [message] if message else [],
                "message_set": {message} if message else set(),
                "path": path,
                "heading": heading,
                "sheet_hints_map": {(h["sheet"], tuple(h["columns"])): h for h in hints2},
            }
            order.append(key)
            continue

        cur = items_by_key[key]
        if _sev_rank(severity) > _sev_rank(str(cur.get("severity") or "")):
            cur["severity"] = severity
        if rule_id and rule_id not in cur.get("rule_id_set", set()):
            cur["rule_ids"].append(rule_id)
            cur["rule_id_set"].add(rule_id)
        if message and message not in cur.get("message_set", set()):
            cur["messages"].append(message)
            cur["message_set"].add(message)
        shm = cur.get("sheet_hints_map") or {}
        for h in hints2:
            hk = (h["sheet"], tuple(h["columns"]))
            if hk not in shm:
                shm[hk] = h
        cur["sheet_hints_map"] = shm

    out_items: list[dict[str, Any]] = []
    for key in order:
        cur = items_by_key[key]
        rule_ids = list(cur.get("rule_ids") or [])
        messages = list(cur.get("messages") or [])
        out_items.append(
            {
                "severity": str(cur.get("severity") or ""),
                "rule_id": " / ".join(rule_ids) if rule_ids else "",
                "path": str(cur.get("path") or ""),
                "heading": str(cur.get("heading") or ""),
                "message": " / ".join(messages) if messages else "",
                "sheet_hints": list((cur.get("sheet_hints_map") or {}).values()),
            }
        )

    return {
        "report_kind": report_kind,
        "count": len(out_items),
        "items": out_items,
        "note": "best-effort mapping (sheet/columns) from spec.input_paths + v2 mapping doc",
    }


def main() -> None:
    ap = argparse.ArgumentParser(
        description=(
            "QA next-actions helper: read validation_report_*.json and print sheet/column hints (best-effort). "
            "Designed to improve usability without touching core QA code."
        )
    )
    ap.add_argument("--validation-report", required=True, type=Path, help="validation_report_eia.json or validation_report_dia.json")
    ap.add_argument(
        "--kind",
        choices=["auto", "EIA", "DIA"],
        default="auto",
        help="Override report kind (default: auto).",
    )
    ap.add_argument("--out-json", type=Path, default=None, help="Optional output JSON path")
    ap.add_argument("--out-md", type=Path, default=None, help="Optional output markdown path")

    args = ap.parse_args()

    report_path = args.validation_report.expanduser().resolve()
    obj = _load_validation_report(report_path)
    results = obj.get("results") or []

    kind = str(args.kind).strip().upper()
    if kind not in {"EIA", "DIA"}:
        kind = _guess_report_kind_with_fallback(report_path, results)

    root = Path(__file__).resolve().parents[1]
    spec_path = root / "spec" / "sections.yaml" if kind == "EIA" else root / "spec_dia" / "sections.yaml"
    sections = _load_sections_yaml(spec_path)

    table_specs_path = root / "spec" / "table_specs.yaml" if kind == "EIA" else root / "spec_dia" / "table_specs.yaml"
    table_specs: dict[str, dict[str, Any]] | None = None
    if table_specs_path.exists():
        table_specs = _load_table_specs_yaml(table_specs_path)

    next_actions = build_next_actions(report_kind=kind, sections=sections, table_specs=table_specs, results=results)

    # stdout summary
    print(f"Report: {report_path}")
    print(f"Kind: {kind} (spec={spec_path})")
    print(f"Next-actions items: {next_actions['count']}")

    for item in next_actions["items"][:40]:
        sev = item["severity"]
        path = item["path"]
        heading = item["heading"]
        print(f"- [{sev}] {path} {heading}")
        for h in item["sheet_hints"]:
            cols = ", ".join(h["columns"])
            print(f"    -> {h['sheet']}: {cols}  ({h['reason']})")

    if args.out_json:
        out_json = args.out_json.expanduser().resolve()
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(json.dumps(next_actions, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"OK wrote JSON: {out_json}")

    if args.out_md:
        out_md = args.out_md.expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)

        lines: list[str] = []
        lines.append(f"# QA Next Actions ({kind})")
        lines.append("")
        lines.append(f"- report: `{report_path}`")
        lines.append(f"- spec: `{spec_path}`")
        lines.append("")
        for item in next_actions["items"]:
            lines.append(f"## {item['path']} — {item['heading']}")
            lines.append(f"- severity: {item['severity']}")
            lines.append(f"- rule_id: {item['rule_id']}")
            lines.append(f"- message: {item['message']}")
            lines.append("- sheet_hints:")
            for h in item["sheet_hints"]:
                cols = ", ".join(h["columns"])
                lines.append(f"  - `{h['sheet']}`: {cols} ({h['reason']})")
            lines.append("")

        out_md.write_text("\n".join(lines), encoding="utf-8")
        print(f"OK wrote MD: {out_md}")


if __name__ == "__main__":
    main()
