from __future__ import annotations

from typing import Any

from eia_gen.models.case import Case
from eia_gen.models.fields import QuantityField, TextField


def _text_fact(field: TextField) -> dict[str, Any]:
    text = field.t.strip()
    missing = text == ""
    return {
        "text": text if text else "【작성자 기입 필요】",
        "missing": missing,
        "source_ids": field.src or (["S-TBD"] if missing else []),
        "note": field.note,
        "confidential": field.confidential,
    }


def _quantity_fact(field: QuantityField, default_unit: str | None = None) -> dict[str, Any]:
    missing = field.v is None
    unit = field.u or default_unit
    value = field.v
    return {
        "value": value,
        "unit": unit,
        "missing": missing,
        "source_ids": field.src or (["S-TBD"] if missing else []),
        "note": field.note,
    }


def build_facts(case: Case, section_id: str) -> dict[str, Any]:
    # Spec(SSOT) section IDs aliasing (v1)
    section_id = {
        "CH1_PURPOSE": "CH1_OVERVIEW",
        "CH1_LOCATION_AREA": "CH1_OVERVIEW",
        "CH1_SCALE": "CH1_OVERVIEW",
        "CH1_SCHEDULE": "CH1_OVERVIEW",
        "CH1_APPLICABILITY": "CH1_PERMITS",
        "CH2_TOPO": "CH2_NAT_TG",
        "CH2_ECO": "CH2_NAT_ECO",
        "CH2_WATER": "CH2_NAT_WATER",
        "CH2_AIR": "CH2_LIFE_AIR",
        "CH2_NOISE": "CH2_LIFE_NOISE",
        "CH2_LANDUSE": "CH2_SOC_LANDUSE",
        "CH2_LANDSCAPE": "CH2_SOC_LANDSCAPE",
        "CH2_POP_TRAFFIC": "CH2_SOC_POP",
        "CH4_MITIGATION": "CH4_TEXT",
        "CH5_TRACKER": "CH5_TEXT",
    }.get(section_id, section_id)

    po = case.project_overview

    extra = case.model_extra or {}
    ssot_overrides = extra.get("ssot_page_overrides")
    appendix_inserts = extra.get("appendix_inserts")
    dia_auto_generated = extra.get("dia_auto_generated")

    base: dict[str, Any] = {
        "meta": {
            "project_type": case.meta.project_type,
            "report_type": case.meta.report_type,
            "template_version": case.meta.template_version,
        },
        "project": {
            "project_name": _text_fact(case.cover.project_name),
            "purpose_need": _text_fact(po.purpose_need),
            "address": _text_fact(po.location.address),
            "total_area_m2": _quantity_fact(po.area.total_area_m2, default_unit="m2"),
        },
        "survey_plan": {
            "radius_m": _quantity_fact(case.survey_plan.influence_area.radius_m, default_unit="m"),
            "justification": _text_fact(case.survey_plan.influence_area.justification),
            "methods": {k: _text_fact(v) for k, v in case.survey_plan.methods.items()},
        },
    }

    if isinstance(ssot_overrides, list) and ssot_overrides:
        base["ssot_page_overrides"] = ssot_overrides
    if isinstance(appendix_inserts, list) and appendix_inserts:
        base["appendix_inserts"] = appendix_inserts
    if isinstance(dia_auto_generated, dict) and dia_auto_generated:
        base["dia_auto_generated"] = dia_auto_generated

    # Extra blocks (best-effort)
    disaster = getattr(case, "disaster", None)
    if isinstance(disaster, dict) and disaster:
        base["disaster"] = disaster

    if section_id == "CH0_SUMMARY":
        base["summary_inputs"] = {
            "key_issues": [_text_fact(x) for x in case.summary_inputs.key_issues],
            "key_measures": [_text_fact(x) for x in case.summary_inputs.key_measures],
        }
        base["facilities"] = [
            {
                "category": _text_fact(f.category),
                "name": _text_fact(f.name),
                "qty": _quantity_fact(f.qty),
                "area_m2": _quantity_fact(f.area_m2, default_unit="m2"),
                "capacity_person": _quantity_fact(f.capacity_person, default_unit="명"),
            }
            for f in po.contents_scale.facilities
        ]
        return base

    if section_id in {"CH0_COVER", "DIA0_COVER"}:
        base["cover"] = {
            "project_name": _text_fact(case.cover.project_name),
            "author_org": _text_fact(case.cover.author_org),
            "submit_date": _text_fact(case.cover.submit_date),
            "approving_authority": _text_fact(case.cover.approving_authority),
            "consultation_agency": _text_fact(case.cover.consultation_agency),
        }
        return base

    if section_id == "CH1_OVERVIEW":
        base["admin"] = {
            "sido": _text_fact(po.location.admin.sido),
            "sigungu": _text_fact(po.location.admin.sigungu),
            "eupmyeon": _text_fact(po.location.admin.eupmyeon),
            "center_coord": {
                "lat": _quantity_fact(po.location.center_coord.lat, default_unit="deg"),
                "lon": _quantity_fact(po.location.center_coord.lon, default_unit="deg"),
            },
        }
        base["parcels"] = [
            {
                "jibun": _text_fact(p.jibun),
                "pnu": _text_fact(p.pnu),
                "land_category": _text_fact(p.land_category),
                "zoning": _text_fact(p.zoning),
                "area_m2": _quantity_fact(p.area_m2, default_unit="m2"),
            }
            for p in po.area.parcels
        ]
        base["facilities"] = [
            {
                "category": _text_fact(f.category),
                "name": _text_fact(f.name),
                "qty": _quantity_fact(f.qty),
                "area_m2": _quantity_fact(f.area_m2, default_unit="m2"),
                "capacity_person": _quantity_fact(f.capacity_person, default_unit="명"),
                "note": _text_fact(f.note),
            }
            for f in po.contents_scale.facilities
        ]
        base["schedule"] = [
            {"phase": _text_fact(m.phase), "start": _text_fact(m.start), "end": _text_fact(m.end)}
            for m in po.schedule.milestones
        ]
        base["permits"] = [
            {
                "name": _text_fact(p.name),
                "status": _text_fact(p.status),
                "authority": _text_fact(p.authority),
                "note": _text_fact(p.note),
            }
            for p in po.legal_permits.permit_list
        ]
        return base

    if section_id == "CH1_PERMITS":
        base["cover"] = {
            "approving_authority": _text_fact(case.cover.approving_authority),
            "consultation_agency": _text_fact(case.cover.consultation_agency),
        }
        base["applicability_raw"] = case.applicability
        base["prior_assessments_raw"] = case.prior_assessments
        base["permits"] = [
            {
                "name": _text_fact(p.name),
                "status": _text_fact(p.status),
                "authority": _text_fact(p.authority),
                "note": _text_fact(p.note),
            }
            for p in po.legal_permits.permit_list
        ]
        return base

    if section_id == "CH2_METHOD":
        return base

    if section_id == "CH2_BASELINE_SUMMARY":
        # Table-only section; include baseline payload so rule-based narrative can cite the same sources.
        base["baseline"] = case.baseline.model_dump()
        return base

    if section_id == "CH2_NAT_TG":
        tg = case.baseline.topography_geology
        base["baseline"] = {
            "elevation_range_m": _text_fact(tg.elevation_range_m),
            "mean_slope_deg": _quantity_fact(tg.mean_slope_deg, default_unit="deg"),
            "geology_summary": _text_fact(tg.geology_summary),
            "soil_summary": _text_fact(tg.soil_summary),
        }
        return base

    if section_id == "CH2_NAT_ECO":
        eco = case.baseline.ecology
        base["baseline"] = {
            "survey_dates": [_text_fact(d) for d in eco.survey_dates],
            "flora_list": [
                {
                    "species_ko": _text_fact(x.species_ko),
                    "scientific": _text_fact(x.scientific),
                    "protected": _text_fact(x.protected),
                    "note": _text_fact(x.note),
                }
                for x in eco.flora_list
            ],
            "fauna_list": [
                {
                    "species_ko": _text_fact(x.species_ko),
                    "scientific": _text_fact(x.scientific),
                    "protected": _text_fact(x.protected),
                    "evidence": _text_fact(x.evidence),
                    "note": _text_fact(x.note),
                }
                for x in eco.fauna_list
            ],
        }
        base["assets"] = [
            {
                "asset_id": a.asset_id,
                "type": a.type,
                "file_path": a.file_path,
                "caption": _text_fact(a.caption),
                "source_ids": a.source_ids,
                "viewpoint": _text_fact(a.viewpoint),
            }
            for a in case.assets
        ]
        return base

    if section_id == "CH2_NAT_WATER":
        w = case.baseline.water_environment
        base["baseline"] = {
            "streams": [
                {
                    "name": _text_fact(s.name),
                    "distance_m": _quantity_fact(s.distance_m, default_unit="m"),
                    "flow_direction": _text_fact(s.flow_direction),
                    "note": _text_fact(s.note),
                }
                for s in w.streams
            ],
            "water_quality": w.water_quality,
        }
        return base

    if section_id == "CH2_LIFE_AIR":
        a = case.baseline.air_quality
        base["baseline"] = {
            "station_name": _text_fact(a.station_name),
            "pm10_ugm3": _quantity_fact(a.pm10_ugm3, default_unit="µg/m3"),
            "pm25_ugm3": _quantity_fact(a.pm25_ugm3, default_unit="µg/m3"),
            "ozone_ppm": _quantity_fact(a.ozone_ppm, default_unit="ppm"),
        }
        return base

    if section_id == "CH2_LIFE_NOISE":
        nv = case.baseline.noise_vibration
        base["baseline"] = {
            "receptors": [
                {
                    "name": _text_fact(r.name),
                    "distance_m": _quantity_fact(r.distance_m, default_unit="m"),
                    "baseline_day_db": _quantity_fact(r.baseline_day_db, default_unit="dB(A)"),
                    "baseline_night_db": _quantity_fact(r.baseline_night_db, default_unit="dB(A)"),
                    "measured": _text_fact(r.measured),
                }
                for r in nv.receptors
            ]
        }
        return base

    if section_id == "CH2_LIFE_ODOR":
        # v1: 상세 입력 구조가 아직 없으므로 baseline 전체를 전달
        base["baseline"] = case.baseline.model_dump()
        return base

    if section_id == "CH2_SOC_LANDUSE":
        ll = case.baseline.landuse_landscape
        base["baseline"] = {
            "current_landcover_summary": _text_fact(ll.current_landcover_summary),
            "protected_areas_overlap": _text_fact(ll.protected_areas_overlap),
        }
        base["assets"] = [
            {
                "asset_id": a.asset_id,
                "type": a.type,
                "file_path": a.file_path,
                "caption": _text_fact(a.caption),
                "source_ids": a.source_ids,
            }
            for a in case.assets
        ]
        return base

    if section_id == "CH2_SOC_LANDSCAPE":
        ll = case.baseline.landuse_landscape
        base["baseline"] = {
            "viewpoints": [
                {
                    "vp_id": _text_fact(v.vp_id),
                    "location_desc": _text_fact(v.location_desc),
                    "photo_asset_id": _text_fact(v.photo_asset_id),
                    "note": _text_fact(v.note),
                }
                for v in ll.key_viewpoints
            ]
        }
        base["assets"] = [
            {
                "asset_id": a.asset_id,
                "type": a.type,
                "file_path": a.file_path,
                "caption": _text_fact(a.caption),
                "source_ids": a.source_ids,
                "viewpoint": _text_fact(a.viewpoint),
            }
            for a in case.assets
        ]
        return base

    if section_id == "CH2_SOC_POP":
        pt = case.baseline.population_traffic
        base["baseline"] = {
            "nearest_village": _text_fact(pt.nearest_village),
            "distance_to_village_m": _quantity_fact(pt.distance_to_village_m, default_unit="m"),
            "access_roads": [_text_fact(r) for r in pt.access_roads],
            "expected_vehicles_per_day": _quantity_fact(pt.expected_vehicles_per_day, default_unit="대/일"),
        }
        return base

    if section_id == "CH3_SCOPING":
        base["scoping_matrix"] = [
            {
                "item_id": s.item_id,
                "item_name": s.item_name,
                "category": _text_fact(s.category),
                "exclude_reason": _text_fact(s.exclude_reason),
                "baseline_method": _text_fact(s.baseline_method),
                "prediction_method": _text_fact(s.prediction_method),
                "src_expected": s.src_expected,
            }
            for s in case.scoping_matrix
        ]
        return base

    if section_id == "CH3_CONSTRUCTION":
        base["scoping_matrix"] = [
            {
                "item_id": s.item_id,
                "item_name": s.item_name,
                "class": s.scoping_class.value,
                "exclude_reason": _text_fact(s.exclude_reason),
                "baseline_method": _text_fact(s.baseline_method),
                "prediction_method": _text_fact(s.prediction_method),
            }
            for s in case.scoping_matrix
        ]
        base["impact_prediction_raw"] = case.impact_prediction.construction
        base["mitigation_measures"] = [
            {
                "measure_id": m.measure_id,
                "phase": _text_fact(m.phase),
                "title": _text_fact(m.title),
                "description": _text_fact(m.description),
                "related_impacts": m.related_impacts,
            }
            for m in case.mitigation.measures
            if m.phase.t.strip() == "공사"
        ]
        return base

    if section_id == "CH3_OPERATION":
        base["scoping_matrix"] = [
            {
                "item_id": s.item_id,
                "item_name": s.item_name,
                "class": s.scoping_class.value,
                "exclude_reason": _text_fact(s.exclude_reason),
                "baseline_method": _text_fact(s.baseline_method),
                "prediction_method": _text_fact(s.prediction_method),
            }
            for s in case.scoping_matrix
        ]
        base["impact_prediction_raw"] = case.impact_prediction.operation
        base["mitigation_measures"] = [
            {
                "measure_id": m.measure_id,
                "phase": _text_fact(m.phase),
                "title": _text_fact(m.title),
                "description": _text_fact(m.description),
                "related_impacts": m.related_impacts,
            }
            for m in case.mitigation.measures
            if m.phase.t.strip() == "운영"
        ]
        return base

    if section_id == "CH4_TEXT":
        base["mitigation_measures"] = [
            {
                "measure_id": m.measure_id,
                "phase": _text_fact(m.phase),
                "title": _text_fact(m.title),
                "description": _text_fact(m.description),
                "monitoring": _text_fact(m.monitoring),
                "related_impacts": m.related_impacts,
            }
            for m in case.mitigation.measures
        ]
        return base

    if section_id == "CH5_TEXT":
        base["condition_tracker"] = [
            {
                "item": _text_fact(x.item),
                "measure_id": _text_fact(x.measure_id),
                "when": _text_fact(x.when),
                "evidence": _text_fact(x.evidence),
                "responsible": _text_fact(x.responsible),
            }
            for x in case.management_plan.implementation_register
        ]
        return base

    if section_id == "CH7_CONCLUSION":
        base["scoping_matrix"] = [
            {"item_id": s.item_id, "item_name": s.item_name, "class": s.scoping_class.value}
            for s in case.scoping_matrix
        ]
        base["mitigation_measures"] = [
            {
                "measure_id": m.measure_id,
                "phase": _text_fact(m.phase),
                "title": _text_fact(m.title),
                "related_impacts": m.related_impacts,
            }
            for m in case.mitigation.measures
        ]
        base["management_plan"] = case.management_plan.model_dump()
        return base

    # default: minimal
    return base
