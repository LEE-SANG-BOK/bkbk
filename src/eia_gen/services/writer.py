from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path
from typing import Any

from eia_gen.models.case import Case
from eia_gen.models.case import ScopingClass
from eia_gen.models.sources import SourceRegistry
from eia_gen.services.citations import ensure_citation, format_citations, normalize_ids
from eia_gen.services.conditions import eval_condition
from eia_gen.services.draft import FigureDraft, ReportDraft, SectionDraft, TableDraft
from eia_gen.services.facts import build_facts
from eia_gen.services.llm.base import LLMClient
from eia_gen.services.sections import SECTION_SPECS, SectionSpec
from eia_gen.spec.models import SpecBundle


def _fact_text(f: dict[str, Any], placeholder: str = "【작성자 기입 필요】") -> str:
    t = (f.get("text") or "").strip()
    return t if t else placeholder


def _fact_value_with_unit(f: dict[str, Any]) -> str:
    v = f.get("value")
    u = f.get("unit")
    if v is None:
        return "【작성자 기입 필요】"
    if isinstance(v, (int, float)):
        if float(v).is_integer():
            vv = f"{int(v):,}"
        else:
            vv = f"{v:,.2f}"
    else:
        vv = str(v)
    return f"{vv}{u or ''}"


def _any_text(x: Any) -> str:
    if x is None:
        return ""
    # TextField / QuantityField (pydantic)
    t = getattr(x, "t", None)
    if isinstance(t, str):
        return t.strip()
    v = getattr(x, "v", None)
    if v is not None:
        return str(v).strip()
    if isinstance(x, dict):
        if "text" in x:
            return str(x.get("text") or "").strip()
        if "t" in x:
            return str(x.get("t") or "").strip()
        if "value" in x:
            return str(x.get("value") or "").strip()
        if "v" in x:
            return str(x.get("v") or "").strip()
    return str(x).strip()


def _origin_label(origin: str) -> str:
    o = (origin or "").strip().upper()
    if not o:
        return ""
    mapping = {
        "FIELD_SURVEY": "현지확인",
        "FIELD": "현지확인",
        "INTERVIEW": "탐문/인터뷰",
        "OFFICIAL_DB": "공공DB",
        "PUBLIC_API": "공공DB",
        "API": "공공DB",
        "LITERATURE": "문헌",
        "CLIENT_PROVIDED": "사업자 제공",
        "CLIENT_DOC": "사업자 제공",
        "MODEL_OUTPUT": "모델/계산",
        "MODEL": "모델/계산",
        "GIS_LAYER": "공간분석",
        "GIS": "공간분석",
    }
    return mapping.get(o, origin.strip())


def _summarize_origins(rows: Any) -> str:
    if not isinstance(rows, list):
        return ""
    labels: list[str] = []
    for r in rows:
        if not isinstance(r, dict):
            continue
        lbl = _origin_label(_any_text(r.get("data_origin")))
        if lbl:
            labels.append(lbl)
    labels = sorted(set(labels))
    return ", ".join(labels)


def _collect_source_ids(*facts: Any) -> list[str]:
    ids: list[str] = []

    def visit(x: Any) -> None:
        if isinstance(x, dict):
            if "source_ids" in x and isinstance(x["source_ids"], list):
                ids.extend([str(s) for s in x["source_ids"]])
            # xlsx/raw dict fields often store citations in "src"
            src = x.get("src")
            if isinstance(src, list):
                ids.extend([str(s) for s in src])
            elif isinstance(src, str):
                ids.append(src)
            for v in x.values():
                visit(v)
        elif isinstance(x, list):
            for v in x:
                visit(v)

    for f in facts:
        visit(f)
    return normalize_ids(ids)


_TBD_SOURCE_IDS = {"S-TBD", "SRC-TBD"}


def _filter_tbd_source_ids(ids: list[str]) -> list[str]:
    return [s for s in normalize_ids(ids) if s not in _TBD_SOURCE_IDS]


def _collect_source_ids_no_tbd(*facts: Any, fallback: list[str] | None = None) -> list[str]:
    ids = _filter_tbd_source_ids(_collect_source_ids(*facts))
    if ids:
        return ids
    return _filter_tbd_source_ids(fallback or [])


_ITEM_SECTION_TO_ITEM_ID: dict[str, str] = {
    "CH2_NAT_TG": "NAT_TG",
    "CH2_NAT_ECO": "NAT_ECO",
    "CH2_NAT_WATER": "NAT_WATER",
    "CH2_LIFE_AIR": "LIFE_AIR",
    "CH2_LIFE_NOISE": "LIFE_NOISE",
    "CH2_LIFE_ODOR": "LIFE_ODOR",
    "CH2_SOC_LANDUSE": "SOC_LANDUSE",
    "CH2_SOC_LANDSCAPE": "SOC_LANDSCAPE",
    "CH2_SOC_POP": "SOC_POP",
}

_SSOT_CHANGWON_SAMPLE_SOURCE_ID = "S-CHANGWON-SAMPLE"
_SSOT_CHANGWON_SAMPLE_PDF_BASENAME = "25. 창원 마산합포구 진저명 관광농원 기허가 샘플링.pdf"
_SSOT_CHANGWON_REFERENCE_PACK_ID = "CHANGWON_JINJEON_APPROVED_2025"
_SSOT_SAMPLE_PDF_ENV = "EIA_GEN_SSOT_SAMPLE_PDF"


def _guess_repo_root() -> Path | None:
    # eia-gen/src/eia_gen/services/writer.py -> eia-gen/
    try:
        return Path(__file__).resolve().parents[3]
    except Exception:
        return None


def _resolve_existing_file_path(raw: str, *, repo_root: Path | None) -> Path | None:
    s = (raw or "").strip()
    if not s:
        return None
    p = Path(s).expanduser()
    if p.is_absolute():
        return p.resolve() if p.exists() else None
    if repo_root is not None:
        cand = (repo_root / p).expanduser()
        if cand.exists():
            return cand.resolve()
    # last resort: relative to current working dir
    return p.resolve() if p.exists() else None


def _resolve_ssot_changwon_sample_pdf_path(sources: SourceRegistry | None) -> str:
    """Resolve the Changwon sample PDF path for SSOT PDF_PAGE embedding.

    Precedence:
    1) env: EIA_GEN_SSOT_SAMPLE_PDF
    2) sources.yaml: source_id=S-CHANGWON-SAMPLE local_file/file_path
    3) repo-relative fallbacks: reference_packs/..., spec_dia/...
    """
    repo_root = _guess_repo_root()

    env = os.getenv(_SSOT_SAMPLE_PDF_ENV, "").strip()
    p = _resolve_existing_file_path(env, repo_root=repo_root) if env else None
    if p is not None:
        return str(p)

    configured: str | None = None
    if sources is not None:
        entry = sources.get(_SSOT_CHANGWON_SAMPLE_SOURCE_ID)
        if entry and (entry.file_path or "").strip():
            configured = str(entry.file_path or "").strip()
            p2 = _resolve_existing_file_path(configured, repo_root=repo_root)
            if p2 is not None:
                return str(p2)

    basename = Path(configured).name if configured else _SSOT_CHANGWON_SAMPLE_PDF_BASENAME
    candidates: list[Path] = []
    if repo_root is not None:
        candidates.extend(
            [
                repo_root / "reference_packs" / _SSOT_CHANGWON_REFERENCE_PACK_ID / "assets" / basename,
                repo_root / "reference_packs" / _SSOT_CHANGWON_REFERENCE_PACK_ID / basename,
                repo_root / "spec_dia" / basename,
                repo_root / "spec_dia" / _SSOT_CHANGWON_SAMPLE_PDF_BASENAME,
            ]
        )

    for cand in candidates:
        try:
            if cand.exists():
                return str(cand.resolve())
        except Exception:
            continue

    return configured or f"spec_dia/{_SSOT_CHANGWON_SAMPLE_PDF_BASENAME}"


_SSOT_CHANGWON_SAMPLE_PDF_PAGE_RANGES: dict[str, tuple[int, int]] = {
    # NOTE: pages are 1-based and inclusive ranges within the sample PDF.
    "SSOT_CH2_REUSE_PDF": (1, 41),
    "SSOT_CH3_REUSE_PDF": (42, 49),
    "SSOT_CH4_REUSE_PDF": (50, 56),
    "SSOT_CH5_REUSE_PDF": (57, 64),
    # NOTE: The sample PDF's Chapter 7 cover starts at p070 (see split TOC under
    # output/pdf_split/changwon_sample_gingerfarm_2025/).
    "SSOT_CH6_REUSE_PDF": (65, 69),
    "SSOT_CH7_REUSE_PDF": (70, 322),
    "SSOT_CH8_REUSE_PDF": (323, 328),
}


def _ssot_page_override_map(raw: Any) -> dict[int, dict[str, Any]]:
    """Return sample_pdf_page -> override mapping (best-effort).

    Stored in v2 case.xlsx sheet `SSOT_PAGE_OVERRIDES` and loaded into Case extras as:
      facts["ssot_page_overrides"] = [
        { "sample_page": 264, "file_path": "attachments/normalized/ATT-0001__....pdf", "page": 11, ... },
      ]
    """
    if not isinstance(raw, list):
        return {}

    out: dict[int, dict[str, Any]] = {}
    for row in raw:
        if not isinstance(row, dict):
            continue

        sample_page_raw = row.get("sample_page")
        try:
            sample_page = int(sample_page_raw)
        except Exception:
            continue

        file_path = str(row.get("file_path") or row.get("override_file_path") or "").strip()
        if not file_path:
            continue

        page_raw = row.get("page") or row.get("override_page")
        try:
            page = int(page_raw)
        except Exception:
            continue
        if page <= 0:
            continue

        width_mm: float | None
        width_mm_raw = row.get("width_mm")
        try:
            width_mm = float(width_mm_raw) if width_mm_raw not in (None, "") else None
        except Exception:
            width_mm = None

        dpi: int | None
        dpi_raw = row.get("dpi")
        try:
            dpi = int(dpi_raw) if dpi_raw not in (None, "") else None
        except Exception:
            dpi = None

        crop = str(row.get("crop") or "").strip() or None

        src_raw = row.get("src_ids") or row.get("src_id") or row.get("src")
        src_ids: list[str] = []
        if isinstance(src_raw, str):
            src_ids = [s.strip() for s in src_raw.replace(",", ";").split(";") if s.strip()]
        elif isinstance(src_raw, list):
            src_ids = [str(s).strip() for s in src_raw if str(s).strip()]
        src_ids = normalize_ids(src_ids)

        out[sample_page] = {
            "file_path": file_path,
            "page": page,
            "width_mm": width_mm,
            "dpi": dpi,
            "crop": crop,
            "src_ids": src_ids,
        }

    return out


def _extract_prior_omission(case: Case) -> dict[str, Any]:
    """Best-effort extraction of omission rule from case.yaml.

    Supported locations (v1 best effort):
    - case.prior_assessment_linkage.omission_rule (if present as extra)
    """

    def _truthy(v: Any) -> bool:
        if v is None:
            return False
        if isinstance(v, bool):
            return v
        if isinstance(v, (int, float)):
            return v != 0
        if isinstance(v, dict) and "t" in v:
            return _truthy(v.get("t"))
        s = str(v).strip().lower()
        return s in {"true", "t", "yes", "y", "1"}

    allow_omit = False
    omit_item_ids: list[str] = []
    legal_basis_text = ""

    # Gather candidate locations (accept multiple template variants)
    candidates: list[dict[str, Any]] = []

    extra = case.model_extra or {}
    pal = extra.get("prior_assessment_linkage")
    if isinstance(pal, dict):
        om = pal.get("omission_rule")
        if isinstance(om, dict):
            candidates.append(om)
        candidates.append(pal)

    pa = case.prior_assessments
    coverage_level: str = ""
    if isinstance(pa, dict):
        om = pa.get("omission_rule")
        if isinstance(om, dict):
            candidates.append(om)
        seia = pa.get("strategic_eia")
        if isinstance(seia, dict):
            coverage_level = str(seia.get("coverage_level") or "").strip().lower()
            om2 = seia.get("omission_rule")
            if isinstance(om2, dict):
                candidates.append(om2)
            ob = seia.get("omission_basis")
            if isinstance(ob, dict):
                candidates.append(ob)
            candidates.append(seia)

    # Merge: first non-empty wins
    for c in candidates:
        if not allow_omit:
            if "allow_omit" in c:
                allow_omit = _truthy(c.get("allow_omit"))
            elif "article_60_applicable" in c:
                allow_omit = _truthy(c.get("article_60_applicable"))

        if not omit_item_ids:
            raw = c.get("omit_item_ids") or c.get("reviewed_item_ids") or c.get("reviewed_item_ids")
            if isinstance(raw, list):
                omit_item_ids = [str(x) for x in raw if str(x).strip()]

        if not legal_basis_text:
            legal_basis_text = str(
                c.get("legal_basis_text")
                or c.get("omission_detail")
                or c.get("omission_basis_text")
                or c.get("legal_basis")
                or ""
            ).strip()

    # If omission is enabled and coverage is "full", allow omitting all item-sections when not specified.
    if allow_omit and not omit_item_ids and coverage_level == "full":
        omit_item_ids = [s.item_id for s in case.scoping_matrix]

    return {
        "allow_omit": allow_omit,
        "omit_item_ids": omit_item_ids,
        "legal_basis_text": legal_basis_text,
    }


def _omitted_section(spec: SectionSpec, legal_basis_text: str) -> SectionDraft:
    basis = legal_basis_text.strip() or "【작성자 기입 필요】(생략 적용 근거)"
    paragraph = f"선행평가에서 이미 검토된 항목으로 판단되어 본 절은 생략한다. 근거: {basis}."
    paragraph = ensure_citation(paragraph, ["S-TBD"])
    todos = []
    if "【작성자 기입 필요】" in basis:
        todos.append("선행평가 생략 근거(조문/선행평가 증빙) 입력 필요")
    return SectionDraft(section_id=spec.section_id, title=spec.title, paragraphs=[paragraph], todos=todos)


def _excluded_section(spec: SectionSpec, item_name: str, exclude_reason: str, src_ids: list[str]) -> SectionDraft:
    reason = exclude_reason.strip() or "【작성자 기입 필요】(제외 사유)"
    paragraph = f"본 항목({item_name})은(는) 다음 사유로 평가에서 제외한다: {reason}."
    paragraph = ensure_citation(paragraph, src_ids)
    todos = []
    if "【작성자 기입 필요】" in reason:
        todos.append(f"{spec.section_id}: 제외 사유 입력 필요")
    return SectionDraft(section_id=spec.section_id, title=spec.title, paragraphs=[paragraph], todos=todos)


def _rule_based_section(spec: SectionSpec, facts: dict[str, Any], *, sources: SourceRegistry | None = None) -> SectionDraft:
    todos: list[str] = []
    paras: list[str] = []

    sid = spec.section_id
    # Support both legacy(v1) and current(SSOT) section IDs with one implementation,
    # while preserving `sid` in output drafts.
    sid_norm = {
        "CH2_NAT_TG": "CH2_TOPO",
        "CH2_NAT_ECO": "CH2_ECO",
        "CH2_NAT_WATER": "CH2_WATER",
        "CH2_LIFE_AIR": "CH2_AIR",
        "CH2_LIFE_NOISE": "CH2_NOISE",
        "CH2_LIFE_ODOR": "CH2_ODOR",
        "CH2_SOC_LANDUSE": "CH2_LANDUSE",
        "CH2_SOC_LANDSCAPE": "CH2_LANDSCAPE",
        "CH2_SOC_POP": "CH2_POP_TRAFFIC",
        "CH4_TEXT": "CH4_MITIGATION",
        "CH5_TEXT": "CH5_TRACKER",
    }.get(sid, sid)

    project = facts.get("project", {})
    p_name = _fact_text(project.get("project_name", {}))
    address = _fact_text(project.get("address", {}))
    total_area = _fact_value_with_unit(project.get("total_area_m2", {}))

    if sid == "CH0_COVER":
        cover = facts.get("cover", {})
        project_name = _fact_text(cover.get("project_name", {}))
        author_org = _fact_text(cover.get("author_org", {}), placeholder="").strip()
        submit_date = _fact_text(cover.get("submit_date", {}), placeholder="").strip()
        approving = _fact_text(cover.get("approving_authority", {}), placeholder="").strip()
        consult = _fact_text(cover.get("consultation_agency", {}), placeholder="").strip()

        paras.append(f"사업명: {project_name}")
        if author_org:
            paras.append(f"작성기관: {author_org}")
        if submit_date:
            paras.append(f"제출일: {submit_date}")
        if approving and consult:
            paras.append(f"승인기관: {approving}, 협의기관: {consult}")
        elif approving:
            paras.append(f"승인기관: {approving}")
        elif consult:
            paras.append(f"협의기관: {consult}")

        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(cover, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "DIA0_COVER":
        cover = facts.get("cover", {})
        project_name = _fact_text(cover.get("project_name", {}))
        author_org = _fact_text(cover.get("author_org", {}), placeholder="").strip()
        submit_date = _fact_text(cover.get("submit_date", {}), placeholder="").strip()

        paras.append(f"사업명: {project_name}")
        if author_org:
            paras.append(f"작성기관: {author_org}")
        if submit_date:
            paras.append(f"제출일: {submit_date}")
        paras.append("문서종류: 소규모재해영향평가서(재해영향성검토서).")
        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(cover, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "CH0_SUMMARY":
        key_issues = facts.get("summary_inputs", {}).get("key_issues", [])
        key_measures = facts.get("summary_inputs", {}).get("key_measures", [])

        issues_text = ", ".join(_fact_text(x) for x in key_issues if _fact_text(x) != "【작성자 기입 필요】")
        measures_text = ", ".join(
            _fact_text(x) for x in key_measures if _fact_text(x) != "【작성자 기입 필요】"
        )

        if not issues_text:
            issues_text = "【작성자 기입 필요】(핵심 이슈)"
            todos.append("요약: 핵심 이슈(key_issues) 입력 필요")
        if not measures_text:
            measures_text = "【작성자 기입 필요】(핵심 저감대책)"
            todos.append("요약: 핵심 저감대책(key_measures) 입력 필요")

        paras.append(f"사업명: {p_name}, 위치: {address}, 면적: {total_area}.")
        paras.append(f"핵심 이슈: {issues_text}.")
        paras.append(f"핵심 저감대책: {measures_text}.")

        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(project, facts.get("summary_inputs", {}), fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "DIA0_SUMMARY":
        # Best-effort summary from inputs; avoid unconditional placeholders.
        sp = facts.get("survey_plan", {})
        disaster = facts.get("disaster", {}) if isinstance(facts.get("disaster", {}), dict) else {}

        radius = _fact_value_with_unit(sp.get("radius_m", {}))
        radius_ok = "【작성자 기입 필요】" not in radius

        just = _fact_text(sp.get("justification", {}), placeholder="").strip()
        if not just or "【작성자 기입 필요】" in just:
            just = (
                f"사업지 배수계통 및 주변 지형·하천 여건을 고려하여 영향권 반경 {radius}를 설정하였다."
                if radius_ok
                else "사업지 배수계통 및 주변 지형·하천 여건을 고려하여 영향권을 설정하였다."
            )
        just = just.rstrip().rstrip(".")

        paras.append(f"사업명: {p_name}, 위치: {address}, 면적: {total_area}.")
        if radius_ok:
            paras.append(f"평가대상지역 설정(참고): 영향권 반경 {radius}.")
        paras.append(f"설정 사유: {just}.")

        parts = disaster.get("target_area_parts") if isinstance(disaster, dict) else None
        if isinstance(parts, list) and parts:
            included_parts: list[str] = []
            excluded_parts: list[str] = []
            unknown_parts: list[str] = []
            for p in parts:
                if not isinstance(p, dict):
                    continue
                part = _any_text(p.get("part")) or "미기재"
                included = _any_text(p.get("included")).upper()
                if included in {"YES", "Y", "O", "TRUE", "1"}:
                    included_parts.append(part)
                elif included in {"NO", "N", "X", "FALSE", "0"}:
                    excluded_parts.append(part)
                else:
                    unknown_parts.append(part)
            bits: list[str] = []
            if included_parts:
                bits.append(f"포함({', '.join(included_parts)})")
            if excluded_parts:
                bits.append(f"제외({', '.join(excluded_parts)})")
            if unknown_parts:
                bits.append(f"추가검토({', '.join(unknown_parts)})")
            if bits:
                paras.append(f"평가대상지역 구성(요약): {', '.join(bits)}.")

        rainfall = disaster.get("rainfall") if isinstance(disaster, dict) else None
        if isinstance(rainfall, list) and rainfall:
            descs: list[str] = []
            for r in rainfall:
                if not isinstance(r, dict):
                    continue
                station = _any_text(r.get("station_name"))
                rain_mm = _fact_value_with_unit(r.get("rainfall_mm", {}))
                if "【작성자 기입 필요】" in rain_mm:
                    rain_mm = ""
                dur = _fact_value_with_unit(r.get("duration_min", {}))
                if "【작성자 기입 필요】" in dur:
                    dur = ""
                freq = _fact_value_with_unit(r.get("frequency_year", {}))
                if "【작성자 기입 필요】" in freq:
                    freq = ""
                parts2: list[str] = []
                if station:
                    parts2.append(station)
                if dur:
                    parts2.append(f"{dur} 지속")
                if freq:
                    parts2.append(f"{freq} 빈도")
                if rain_mm:
                    parts2.append(rain_mm)
                if parts2:
                    descs.append(" / ".join(parts2))
            if descs:
                paras.append(f"설계강우(요약): {'; '.join(descs)}.")

        drainage = disaster.get("drainage_facilities") if isinstance(disaster, dict) else None
        if isinstance(drainage, list) and drainage:
            df_bits: list[str] = []
            for d in drainage[:5]:
                if not isinstance(d, dict):
                    continue
                fid = _any_text(d.get("facility_id"))
                typ = _any_text(d.get("type"))
                cap = _any_text(d.get("capacity"))
                discharge_to = _any_text(d.get("discharge_to"))
                seg: list[str] = [x for x in [fid, typ, cap] if x]
                if discharge_to:
                    seg.append(f"방류: {discharge_to}")
                if seg:
                    df_bits.append(" / ".join(seg))
            if df_bits:
                paras.append(f"주요 배수시설(요약): {'; '.join(df_bits)}.")

        mledger = disaster.get("maintenance_ledger") if isinstance(disaster, dict) else None
        if isinstance(mledger, list) and mledger:
            assets: list[str] = []
            for m in mledger:
                if not isinstance(m, dict):
                    continue
                asset = _any_text(m.get("asset_id")) or _any_text(m.get("facility_name"))
                if asset:
                    assets.append(asset)
            assets = sorted(set(assets))
            if assets:
                paras.append(
                    f"유지관리(요약): {', '.join(assets)} 등에 대해 점검주기/점검항목을 유지관리대장에 제시한다."
                )

        interviews = disaster.get("interviews") if isinstance(disaster, dict) else None
        if isinstance(interviews, list) and interviews:
            n = 0
            n_10y = 0
            for it in interviews:
                if not isinstance(it, dict):
                    continue
                n += 1
                years = getattr(it.get("residence_years"), "v", None)
                try:
                    if years is not None and float(years) >= 10:
                        n_10y += 1
                except Exception:
                    pass
            paras.append(f"주민탐문(현황): {n}건(거주 10년 이상 {n_10y}건).")

        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(
            project,
            sp,
            parts,
            rainfall,
            drainage,
            mledger,
            interviews,
            fallback=fallback_src,
        )
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "DIA1_PROJECT":
        purpose = _fact_text(project.get("purpose_need", {}))
        paras.append(f"본 사업({p_name})은 {purpose}를 목적으로 계획되었다.")
        paras.append(f"사업지는 {address}에 위치하며, 사업면적은 {total_area}이다.")

        facilities = facts.get("facilities", [])
        fac_bits: list[str] = []
        if isinstance(facilities, list) and facilities:
            for f in facilities:
                if not isinstance(f, dict):
                    continue
                name = _any_text(f.get("name"))
                if not name or "【작성자 기입 필요】" in name:
                    continue
                seg: list[str] = [name]
                area = _fact_value_with_unit(f.get("area_m2", {}))
                if "【작성자 기입 필요】" not in area:
                    seg.append(f"면적 {area}")
                cap = _any_text(f.get("capacity"))
                if cap and "【작성자 기입 필요】" not in cap:
                    seg.append(f"수용 {cap}")
                fac_bits.append(" / ".join(seg))
        if fac_bits:
            paras.append(f"주요 시설계획(요약): {'; '.join(fac_bits[:8])}.")
        else:
            paras.append("주요 시설계획은 시설계획표를 기준으로 정리한다.")

        schedule = facts.get("schedule", [])
        sched_bits: list[str] = []
        if isinstance(schedule, list) and schedule:
            for m in schedule:
                if not isinstance(m, dict):
                    continue
                phase = _any_text(m.get("phase"))
                start = _any_text(m.get("start")) or _any_text(m.get("start_date"))
                end = _any_text(m.get("end")) or _any_text(m.get("end_date"))
                if not phase or "【작성자 기입 필요】" in phase:
                    continue
                if start and end:
                    sched_bits.append(f"{phase}({start}~{end})")
                else:
                    sched_bits.append(phase)
        if sched_bits:
            paras.append(f"사업 추진 일정(요약): {', '.join(sched_bits)}.")
        else:
            paras.append("사업 추진 일정은 공정표를 기준으로 정리한다.")

        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(project, facilities, schedule, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "DIA2_TARGET_AREA":
        sp = facts.get("survey_plan", {})
        radius = _fact_value_with_unit(sp.get("radius_m", {}))
        just = _fact_text(sp.get("justification", {}), placeholder="").strip()
        if not just or "【작성자 기입 필요】" in just:
            just = (
                f"사업지 배수계통 및 주변 지형·하천 여건을 고려하여 영향권 반경 {radius}를 설정하였다."
                if "【작성자 기입 필요】" not in radius
                else "사업지 배수계통 및 주변 지형·하천 여건을 고려하여 영향권을 설정하였다."
            )
        just = just.rstrip().rstrip(".")
        paras.append("평가대상지역은 사업지 및 상·하류 영향권(주변지역 포함)을 고려하여 설정한다.")
        if "【작성자 기입 필요】" not in radius:
            paras.append(f"설정 범위(참고): 영향권 반경 {radius}.")
        paras.append(f"설정 사유: {just}.")
        # If available, summarize the 4-part target area (PROJECT/UPSTREAM/DOWNSTREAM/SURROUNDING).
        disaster = facts.get("disaster", {}) if isinstance(facts.get("disaster", {}), dict) else {}
        parts = disaster.get("target_area_parts") if isinstance(disaster, dict) else None
        if isinstance(parts, list) and parts:
            included_parts: list[str] = []
            excluded_parts: list[str] = []
            missing_reasons = False
            for p in parts:
                if not isinstance(p, dict):
                    continue
                part = _any_text(p.get("part"))
                included = _any_text(p.get("included")).upper()
                reason = _any_text(p.get("reason"))
                exclude_reason = _any_text(p.get("exclude_reason"))
                if included in {"YES", "Y", "O", "TRUE", "1"}:
                    included_parts.append(part or "미기재")
                    if not reason:
                        missing_reasons = True
                elif included in {"NO", "N", "X", "FALSE", "0"}:
                    excluded_parts.append(part or "미기재")
                    if not exclude_reason:
                        missing_reasons = True
            if included_parts:
                paras.append(f"입력된 대상지역 구성(포함): {', '.join(included_parts)}.")
            if excluded_parts:
                paras.append(f"입력된 대상지역 구성(제외): {', '.join(excluded_parts)}.")
        paras.append("평가대상지역 설정도 및 검토항목 표를 첨부한다.")
        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(sp, parts, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "DIA3_BASELINE":
        disaster = facts.get("disaster", {}) if isinstance(facts.get("disaster", {}), dict) else {}
        origins_hazard = _summarize_origins(disaster.get("hazard_history")) if isinstance(disaster, dict) else ""
        origins_interviews = _summarize_origins(disaster.get("interviews")) if isinstance(disaster, dict) else ""
        origins_drainage = _summarize_origins(disaster.get("drainage_facilities")) if isinstance(disaster, dict) else ""
        origins_rain = _summarize_origins(disaster.get("rainfall")) if isinstance(disaster, dict) else ""
        origins_bits = [x for x in [origins_hazard, origins_interviews, origins_drainage, origins_rain] if x]
        origins = ", ".join(sorted(set([p for chunk in origins_bits for p in chunk.split(", ") if p])))
        origins_suffix = f"(자료원: {origins})" if origins else ""

        paras.append(
            f"재해 관련 기초현황은 강우자료, 지형·배수체계, 하천/방류처 및 재해발생 현황(자료조사/탐문/현장사진)을 기반으로 정리한다. {origins_suffix}".strip()
        )

        interviews = disaster.get("interviews") if isinstance(disaster, dict) else None
        if isinstance(interviews, list) and interviews:
            n = 0
            n_10y = 0
            for it in interviews:
                if not isinstance(it, dict):
                    continue
                n += 1
                years = getattr(it.get("residence_years"), "v", None)
                try:
                    if years is not None and float(years) >= 10:
                        n_10y += 1
                except Exception:
                    pass
            paras.append(f"주민탐문(익명코드) 입력: {n}건(거주 10년 이상 {n_10y}건).")
        else:
            paras.append("주민탐문 자료는 현장조사/탐문 결과로 보완한다.")

        rainfall = disaster.get("rainfall") if isinstance(disaster, dict) else None
        if isinstance(rainfall, list) and rainfall:
            stations = sorted(
                {s for s in (_any_text(r.get("station_name")) for r in rainfall if isinstance(r, dict)) if s}
            )
            if stations:
                paras.append(f"강우자료(관측소): {', '.join(stations)}.")
        else:
            paras.append("강우자료는 공공DB(관측소) 자료를 기반으로 정리한다.")

        drainage = disaster.get("drainage_facilities") if isinstance(disaster, dict) else None
        if isinstance(drainage, list) and drainage:
            df_ids = sorted(
                {s for s in (_any_text(d.get('facility_id')) for d in drainage if isinstance(d, dict)) if s}
            )
            if df_ids:
                paras.append(f"배수시설(식별자): {', '.join(df_ids)}.")
        else:
            paras.append("배수시설 현황은 도면 및 현장확인 자료를 기반으로 정리한다.")

        paras.append("강우/재해발생 현황/배수시설 현황은 표·도면·증빙으로 첨부한다.")

        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(disaster, project, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "DIA4_ANALYSIS":
        disaster = facts.get("disaster", {}) if isinstance(facts.get("disaster", {}), dict) else {}
        runoff = disaster.get("runoff_basins") if isinstance(disaster, dict) else None
        model = ""
        if isinstance(runoff, list) and runoff:
            for b in runoff:
                if not isinstance(b, dict):
                    continue
                model = _any_text(b.get("model"))
                if model:
                    break
        model = model or "유출(수문)모형"

        paras.append(
            f"재해영향 분석은 사업 전/후 유출량 변화 및 배수체계 영향, 토사유출/침식, 사면·산사태 위험요소(해당 시)를 중심으로 검토한다."
        )
        paras.append(f"유출 검토는 {model}을(를) 적용하여 입력값 및 산정 결과를 표로 제시한다.")

        if isinstance(runoff, list) and runoff:
            summaries: list[str] = []
            for b in runoff[:5]:
                if not isinstance(b, dict):
                    continue
                bid = _any_text(b.get("basin_id"))
                pre = _fact_value_with_unit(b.get("pre_peak_cms", {}))
                post = _fact_value_with_unit(b.get("post_peak_cms", {}))
                if "【작성자 기입 필요】" in pre or "【작성자 기입 필요】" in post:
                    continue
                if bid:
                    summaries.append(f"{bid}: 전 {pre} → 후 {post}")
            if summaries:
                paras.append(f"검토 결과(요약): {'; '.join(summaries)}.")

        paras.append("유역/유출 해석, 토사유출/침식, 사면/산사태 검토 표를 첨부한다.")

        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(disaster, runoff, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "DIA5_MITIGATION":
        disaster = facts.get("disaster", {}) if isinstance(facts.get("disaster", {}), dict) else {}
        drainage = disaster.get("drainage_facilities") if isinstance(disaster, dict) else None
        mledger = disaster.get("maintenance_ledger") if isinstance(disaster, dict) else None

        paras.append("저감대책은 공사/운영 단계별로 정리하고, 배수시설/저류/침사지 등 시설 계획과 연계하여 기술한다.")
        if isinstance(drainage, list) and drainage:
            paras.append("사업지 배수시설(집수정/우수관로/침사지 등)의 설치·정비 및 방류처 관리로 내수·월류 위험을 저감한다.")
        if isinstance(mledger, list) and mledger:
            paras.append("유지관리대장(점검주기/점검항목/증빙)을 통해 저감시설의 성능을 지속적으로 관리한다.")
        paras.append("재해 저감대책 표를 첨부한다.")

        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(disaster, drainage, mledger, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "DIA6_MAINTENANCE":
        auto = facts.get("dia_auto_generated")
        if isinstance(auto, dict) and auto.get("maintenance_ledger"):
            todos.append("DIA 유지관리대장: DRR_MAINTENANCE 미입력으로 placeholder 기반 자동 생성(내용 확정 필요)")
        paras.append("유지관리계획 및 유지관리대장은 재해저감시설(배수시설 등)에 대해 점검주기/점검항목/증빙을 포함하여 작성한다.")
        paras.append("유지관리대장 표를 첨부한다.")
        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(facts.get("disaster", {}), fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "APPX_SOURCES":
        paras.append("본 부록은 입력된 출처원장(sources.yaml) 및 산출된 출처관리표(Source Register)를 기준으로 출처·근거를 정리하였다.")
        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(facts, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "DIA7_CONCLUSION":
        sp = facts.get("survey_plan", {})
        radius = _fact_value_with_unit(sp.get("radius_m", {}))
        radius_ok = "【작성자 기입 필요】" not in radius

        paras.append("평가대상지역 설정, 재해영향 분석 결과, 저감대책 및 유지관리계획을 종합하여 결론을 정리한다.")
        if radius_ok:
            paras.append(f"평가대상지역은 사업지 및 영향권(반경 {radius})을 기준으로 설정하였다.")
        paras.append("배수체계 및 유출 특성을 고려한 저감대책과 유지관리계획을 이행하여 재해영향을 최소화한다.")

        disaster = facts.get("disaster", {}) if isinstance(facts.get("disaster", {}), dict) else {}
        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(project, sp, disaster, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid in {"CH1_OVERVIEW", "CH1_PURPOSE"}:
        purpose = _fact_text(project.get("purpose_need", {}))
        if "【작성자 기입 필요】" in purpose:
            todos.append("1장: 목적/필요성(purpose_need) 입력 필요")

        paras.append(f"본 사업({p_name})은 {purpose}를 목적으로 계획되었다.")
        if sid != "CH1_PURPOSE":
            paras.append(f"사업지는 {address}에 위치하며, 사업면적은 {total_area}이다.")
        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(project, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid in {"CH1_PERMITS", "CH1_APPLICABILITY"}:
        cover = facts.get("cover", {})
        approving = _fact_text(cover.get("approving_authority", {}))
        consult = _fact_text(cover.get("consultation_agency", {}))
        paras.append(f"승인기관: {approving}, 협의기관: {consult}.")
        paras.append("소규모환경영향평가 대상 여부 및 법적 근거는 입력된 자료를 기준으로 정리한다.")
        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(cover, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        if "【작성자 기입 필요】" in approving or "【작성자 기입 필요】" in consult:
            todos.append("1장: 승인/협의기관 입력 필요")
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "CH1_LOCATION_AREA":
        paras.append(f"사업지는 {address}에 위치하며, 사업면적은 {total_area}이다.")
        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(project, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "CH1_SCALE":
        facilities = facts.get("facilities", [])
        has_any = any(
            _fact_text((f or {}).get("name", {})) != "【작성자 기입 필요】"
            and _fact_value_with_unit((f or {}).get("area_m2", {})) != "【작성자 기입 필요】"
            for f in (facilities or [])
        )
        if has_any:
            paras.append("사업 계획은 입력된 시설계획(시설별 면적/수용 등)을 기준으로 정리한다.")
        else:
            paras.append("사업 계획은 입력된 시설계획(시설별 면적/수용 등)을 기준으로 정리한다. 【작성자 기입 필요】")
            todos.append("1장: 시설계획(시설명/면적/수용) 상세 입력 확인")
        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(facilities, project, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "CH1_SCHEDULE":
        schedule = facts.get("schedule", [])
        has_any = any(
            _fact_text((m or {}).get("phase", {})) != "【작성자 기입 필요】"
            and _fact_text((m or {}).get("start", {})) != "【작성자 기입 필요】"
            and _fact_text((m or {}).get("end", {})) != "【작성자 기입 필요】"
            for m in (schedule or [])
        )
        if has_any:
            paras.append("사업 추진 일정은 입력된 공정표를 기준으로 정리한다.")
        else:
            paras.append("사업 추진 일정은 입력된 공정표를 기준으로 정리한다. 【작성자 기입 필요】")
        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(schedule, project, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid in _SSOT_CHANGWON_SAMPLE_PDF_PAGE_RANGES:
        # NOTE: This block is for "max reuse" mode (Changwon sample-based SSOT).
        # We embed full PDF pages as images to achieve the strongest format fidelity quickly.
        pdf = _resolve_ssot_changwon_sample_pdf_path(sources)
        start_page, end_page = _SSOT_CHANGWON_SAMPLE_PDF_PAGE_RANGES[sid]
        pages = list(range(start_page, end_page + 1))
        width_mm = 170
        try:
            dpi = int(os.getenv("EIA_GEN_SSOT_PDF_DPI", "120"))
        except ValueError:
            dpi = 120
        src_ids = ["S-CHANGWON-SAMPLE"]
        overrides = _ssot_page_override_map(facts.get("ssot_page_overrides"))

        appendix_rows: list[dict[str, Any]] = []
        if sid == "SSOT_CH8_REUSE_PDF":
            raw_appendix = facts.get("appendix_inserts")
            if isinstance(raw_appendix, list):
                appendix_rows = [r for r in raw_appendix if isinstance(r, dict)]
                appendix_rows.sort(
                    key=lambda r: (
                        int(r.get("order") or 0),
                        str(r.get("ins_id") or ""),
                    )
                )

        for i, sample_page in enumerate(pages):
            # If we append extra pages at the end (appendix inserts), keep a page break after the last sample page.
            br = 1 if (i < len(pages) - 1 or appendix_rows) else 0
            o_auth: str | None = None
            o = overrides.get(sample_page)
            if o:
                o_pdf = str(o.get("file_path") or pdf).strip() or pdf
                o_page = int(o.get("page") or sample_page)
                o_width = float(o.get("width_mm") or width_mm)
                o_dpi = int(o.get("dpi") or dpi)
                o_crop = str(o.get("crop") or "").strip() or None
                o_src = o.get("src_ids") or src_ids
                if not isinstance(o_src, list) or not all(isinstance(x, str) for x in o_src):
                    o_src = src_ids
            else:
                o_pdf = pdf
                o_page = sample_page
                o_width = float(width_mm)
                o_dpi = int(dpi)
                o_crop = None
                o_src = src_ids
                o_auth = "SSOT_SAMPLE"

            crop_part = f"|crop={o_crop}" if o_crop else ""
            auth_part = f"|auth={o_auth}" if o_auth else ""
            directive = (
                f"[[PDF_PAGE:{o_pdf}|id=SP{sample_page:03d}|page={o_page}|width_mm={o_width}|dpi={o_dpi}{crop_part}{auth_part}|break={br}]]"
            )
            paras.append(ensure_citation(directive, o_src))

        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "CH2_METHOD":
        sp = facts.get("survey_plan", {})
        radius = _fact_value_with_unit(sp.get("radius_m", {}))
        just = _fact_text(sp.get("justification", {}), placeholder="").strip()
        if (not just) or ("【작성자 기입 필요】" in just):
            just = f"사업지 배수계통 및 주변 지형·하천 여건을 고려하여 영향권 반경 {radius}를 설정하였다."
        paras.append(f"조사범위는 사업지 및 영향권(반경 {radius})을 기준으로 설정하였다.")
        paras.append(f"조사범위 설정 사유: {just}.")
        paras.append("조사방법은 문헌/공공DB 자료를 우선 활용하고, 현장조사가 필요한 항목은 별도의 현장조사(증빙자료 포함)로 보완한다.")
        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(sp, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid_norm == "CH2_TOPO":
        b = facts.get("baseline", {}) if isinstance(facts.get("baseline", {}), dict) else {}
        elev_f = b.get("elevation_range_m", {}) if isinstance(b.get("elevation_range_m", {}), dict) else {}
        slope_f = b.get("mean_slope_deg", {}) if isinstance(b.get("mean_slope_deg", {}), dict) else {}
        geo_f = b.get("geology_summary", {}) if isinstance(b.get("geology_summary", {}), dict) else {}
        soil_f = b.get("soil_summary", {}) if isinstance(b.get("soil_summary", {}), dict) else {}

        elev = _fact_text(elev_f)
        slope = _fact_value_with_unit(slope_f)
        geology = _fact_text(geo_f)
        soil = _fact_text(soil_f)

        elev_missing = bool(elev_f.get("missing", True))
        slope_missing = bool(slope_f.get("missing", True))
        geology_missing = bool(geo_f.get("missing", True))
        soil_missing = bool(soil_f.get("missing", True))

        paras.append("지형·지질 현황은 지형도/지질도/토양도 등 기존자료를 활용하여 정리하였다.")
        if not elev_missing:
            paras.append(f"표고 분포: {elev}.")
        if not slope_missing:
            paras.append(f"평균 경사도: {slope}.")
        if not geology_missing:
            paras.append(f"지질 개요: {geology}.")
        if not soil_missing:
            paras.append(f"토양 개요: {soil}.")

        if elev_missing or geology_missing:
            paras.append("지형·지질 기초자료가 미확보되어 추가 자료 확인이 필요하다. 【작성자 기입 필요】")
            todos.append(f"{sid}: ENV_BASE_GEO(표고/지질/토양 요약) 입력 또는 발급본/문헌 근거 연결 필요")

        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(b, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid_norm == "CH2_ECO":
        b = facts.get("baseline", {}) if isinstance(facts.get("baseline", {}), dict) else {}
        dates = b.get("survey_dates") if isinstance(b.get("survey_dates"), list) else []
        flora = b.get("flora_list") if isinstance(b.get("flora_list"), list) else []
        fauna = b.get("fauna_list") if isinstance(b.get("fauna_list"), list) else []

        has_dates = any(isinstance(d, dict) and not bool(d.get("missing", True)) for d in dates)
        has_any_species = bool(flora) or bool(fauna)

        paras.append("동·식물상(자연생태) 현황은 기존 문헌 및 현지확인 결과를 종합하여 정리한다.")
        if has_dates:
            date_txt = ", ".join(_fact_text(d, placeholder="").strip() for d in dates if isinstance(d, dict))
            date_txt = date_txt.strip().strip(",")
            if date_txt:
                paras.append(f"조사일자(기록): {date_txt}.")

        if not has_any_species:
            paras.append("생태 관찰/문헌 근거가 미확보되어 추가 조사 또는 기존자료 확보가 필요하다. 【작성자 기입 필요】")
            todos.append(f"{sid}: ENV_ECO_EVENTS/ENV_ECO_OBS(조사일자/관찰종) 입력 또는 문헌 근거 연결 필요")

        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(b, facts.get("assets", {}), fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid_norm == "CH2_WATER":
        b = facts.get("baseline", {}) if isinstance(facts.get("baseline", {}), dict) else {}
        streams = b.get("streams") if isinstance(b.get("streams"), list) else []
        wq = b.get("water_quality") if isinstance(b.get("water_quality"), dict) else {}

        has_stream = bool(streams)
        has_wq = bool(wq)

        paras.append("수환경 현황은 사업지 주변 수계 및 수질자료를 활용하여 정리하였다.")
        if has_stream:
            names = []
            for s in streams:
                if not isinstance(s, dict):
                    continue
                nm = _fact_text((s.get("name") or {}), placeholder="").strip()
                if nm and "【작성자 기입 필요】" not in nm:
                    names.append(nm)
            if names:
                paras.append(f"주요 수계: {', '.join(sorted(set(names)))}.")

        if has_wq:
            # Keep it simple: list known keys present in the input.
            keys = sorted(k for k, v in wq.items() if v is not None)
            if keys:
                paras.append(f"수질 항목({len(keys)}종)은 입력된 자료를 기준으로 제시한다.")

        if not has_stream and not has_wq:
            paras.append("수계/수질 기초자료가 미확보되어 추가 자료 확인이 필요하다. 【작성자 기입 필요】")
            todos.append(f"{sid}: ENV_BASE_WATER(수계/수질) 입력 또는 공공DB/문헌 근거 연결 필요")

        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(b, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid_norm == "CH2_AIR":
        b = facts.get("baseline", {}) if isinstance(facts.get("baseline", {}), dict) else {}
        station_f = b.get("station_name", {}) if isinstance(b.get("station_name", {}), dict) else {}
        pm10_f = b.get("pm10_ugm3", {}) if isinstance(b.get("pm10_ugm3", {}), dict) else {}
        pm25_f = b.get("pm25_ugm3", {}) if isinstance(b.get("pm25_ugm3", {}), dict) else {}
        o3_f = b.get("ozone_ppm", {}) if isinstance(b.get("ozone_ppm", {}), dict) else {}

        station = _fact_text(station_f)
        pm10 = _fact_value_with_unit(pm10_f)
        pm25 = _fact_value_with_unit(pm25_f)
        o3 = _fact_value_with_unit(o3_f)

        station_missing = bool(station_f.get("missing", True))
        has_any_value = any(not bool(f.get("missing", True)) for f in [pm10_f, pm25_f, o3_f])

        paras.append("대기질 현황은 사업지 인근 대기측정망 자료를 활용하여 정리하였다.")
        if not station_missing:
            paras.append(f"적용 측정소: {station}.")
        if has_any_value:
            parts: list[str] = []
            if not bool(pm10_f.get("missing", True)):
                parts.append(f"PM10 {pm10}")
            if not bool(pm25_f.get("missing", True)):
                parts.append(f"PM2.5 {pm25}")
            if not bool(o3_f.get("missing", True)):
                parts.append(f"O₃ {o3}")
            if parts:
                paras.append("최근 관측값(평균): " + ", ".join(parts) + ".")

        if station_missing or (not has_any_value):
            paras.append("대기질 기초자료가 미확보되어 추가 자료 확인이 필요하다. 【작성자 기입 필요】")
            todos.append(f"{sid}: ENV_BASE_AIR(측정소/오염물질 평균) 입력 또는 AIRKOREA 자동수집 필요")

        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(b, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid_norm == "CH2_NOISE":
        b = facts.get("baseline", {}) if isinstance(facts.get("baseline", {}), dict) else {}
        receptors = b.get("receptors") if isinstance(b.get("receptors"), list) else []

        has_receptors = bool(receptors)
        paras.append("소음·진동 현황은 주변 수음시설 및 기존자료를 활용하여 정리한다.")

        if not has_receptors:
            paras.append("소음·진동 기초자료가 미확보되어 추가 자료 확인이 필요하다. 【작성자 기입 필요】")
            todos.append(f"{sid}: ENV_BASE_NOISE(주요 수음지점/주간·야간 등가소음도) 입력 또는 문헌 근거 연결 필요")

        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(b, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid_norm == "CH2_ODOR":
        # v2: 별도 입력 구조가 약하므로, 오수/폐기물 등 운영계획 기반의 기본 서술로 제공한다.
        paras.append("악취 영향은 공사 및 운영 단계에서의 오수·폐기물 관리 상태에 따라 발생할 수 있으므로, 관련 시설의 적정 운영 및 청결관리를 통해 저감한다.")
        paras.append("악취 관련 민원 발생 가능성이 확인되는 경우, 원인시설 점검 및 저감대책을 추가로 수립·이행한다.")
        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(facts, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid_norm == "CH2_LANDUSE":
        b = facts.get("baseline", {}) if isinstance(facts.get("baseline", {}), dict) else {}
        landcover_f = b.get("current_landcover_summary", {}) if isinstance(b.get("current_landcover_summary", {}), dict) else {}
        overlap_f = b.get("protected_areas_overlap", {}) if isinstance(b.get("protected_areas_overlap", {}), dict) else {}

        landcover = _fact_text(landcover_f)
        overlap = _fact_text(overlap_f)
        landcover_missing = bool(landcover_f.get("missing", True))
        overlap_missing = bool(overlap_f.get("missing", True))

        paras.append("토지이용 현황은 토지이용계획확인서, 항공사진 및 관련 도면을 활용하여 정리하였다.")
        if not landcover_missing:
            paras.append(f"현황 요약: {landcover}.")
        if not overlap_missing:
            paras.append(f"보호지역 중첩 여부: {overlap}.")

        if landcover_missing and overlap_missing:
            paras.append("토지이용 현황 요약이 미확보되어 추가 자료 확인이 필요하다. 【작성자 기입 필요】")
            todos.append(f"{sid}: 토지이용 현황 요약(작성자 입력) 또는 관련 도면/발급본 근거 연결 필요")

        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(b, facts.get("assets", {}), fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid_norm == "CH2_LANDSCAPE":
        b = facts.get("baseline", {}) if isinstance(facts.get("baseline", {}), dict) else {}
        vps = b.get("viewpoints") if isinstance(b.get("viewpoints"), list) else []
        has_vps = bool(vps)

        paras.append("경관 현황은 주요 조망점 및 주변 토지이용 여건을 고려하여 정리하였다.")
        if not has_vps:
            paras.append("주요 조망점 정보가 미확보되어 추가 자료 확인이 필요하다. 【작성자 기입 필요】")
            todos.append(f"{sid}: ENV_LANDSCAPE(조망점/사진) 입력 또는 현지확인 근거 연결 필요")

        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(b, facts.get("assets", {}), fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid_norm == "CH2_POP_TRAFFIC":
        b = facts.get("baseline", {}) if isinstance(facts.get("baseline", {}), dict) else {}
        village_f = b.get("nearest_village", {}) if isinstance(b.get("nearest_village", {}), dict) else {}
        dist_f = b.get("distance_to_village_m", {}) if isinstance(b.get("distance_to_village_m", {}), dict) else {}
        veh_f = b.get("expected_vehicles_per_day", {}) if isinstance(b.get("expected_vehicles_per_day", {}), dict) else {}

        village = _fact_text(village_f)
        dist = _fact_value_with_unit(dist_f)
        veh = _fact_value_with_unit(veh_f)

        village_missing = bool(village_f.get("missing", True))
        dist_missing = bool(dist_f.get("missing", True))
        veh_missing = bool(veh_f.get("missing", True))

        paras.append("인구·주거/교통 현황은 주변 정주여건 및 접근도로 여건을 고려하여 정리하였다.")
        if not village_missing:
            if not dist_missing:
                paras.append(f"인접 취락: {village} (이격거리 {dist}).")
            else:
                paras.append(f"인접 취락: {village}.")
        if not veh_missing:
            paras.append(f"예상 교통량(참고): {veh}.")

        if village_missing and veh_missing:
            paras.append("인접 취락/교통량 등 기초자료가 미확보되어 추가 자료 확인이 필요하다. 【작성자 기입 필요】")
            todos.append(f"{sid}: 인접 취락/접근도로/교통량(작성자 입력) 또는 문헌 근거 연결 필요")

        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(b, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "CH2_BASELINE_SUMMARY":
        paras.append("환경현황 요약표는 항목별 현황자료를 종합하여 제시한다.")
        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(facts.get("baseline", {}), fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "CH3_SCOPING":
        rows = facts.get("scoping_matrix", [])
        has_any = isinstance(rows, list) and any(isinstance(r, dict) for r in rows)
        paras.append("평가항목은 중점/현황/제외로 구분하고, 제외 시 사유를 제시한다.")
        if not has_any:
            paras.append("평가항목(스코핑) 입력이 없어 표/본문이 미완성일 수 있다. 【작성자 기입 필요】")
            todos.append("CH3_SCOPING: ENV_SCOPING(평가항목 선정/제외사유/방법) 입력 필요")
        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(rows, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid in {"CH3_CONSTRUCTION", "CH3_OPERATION"}:
        phase = "공사" if sid == "CH3_CONSTRUCTION" else "운영"
        scoping = facts.get("scoping_matrix", [])
        measures = facts.get("mitigation_measures", [])

        has_scoping = isinstance(scoping, list) and len(scoping) > 0
        has_measures = isinstance(measures, list) and len(measures) > 0

        paras.append(f"{phase} 단계에서 발생 가능한 환경영향 요인 및 저감방안을 항목별로 검토한다.")
        if has_measures:
            titles = []
            for m in measures:
                if not isinstance(m, dict):
                    continue
                t = _fact_text(m.get("title", {}), placeholder="").strip()
                if t and "【작성자 기입 필요】" not in t:
                    titles.append(t)
            if titles:
                paras.append(f"{phase} 단계 주요 저감대책(입력 기준): {', '.join(sorted(set(titles)))}.")

        if not has_scoping:
            paras.append("평가항목(스코핑) 입력이 없어 영향 검토 범위가 불명확하다. 【작성자 기입 필요】")
            todos.append(f"{sid}: ENV_SCOPING 입력 필요")
        if not has_measures:
            paras.append(f"{phase} 단계 저감대책 입력이 부족하여 보완이 필요하다. 【작성자 기입 필요】")
            todos.append(f"{sid}: ENV_MITIGATION({phase} 단계 저감대책) 입력 필요")

        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(scoping, measures, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid_norm == "CH4_MITIGATION":
        measures = facts.get("mitigation_measures", [])
        has_any = isinstance(measures, list) and len(measures) > 0
        paras.append("저감방안은 공사/운영 단계별로 정리하고, 영향예측 결과와 연계하여 기술한다.")
        if not has_any:
            paras.append("저감대책 입력이 부족하여 표/본문이 미완성일 수 있다. 【작성자 기입 필요】")
            todos.append(f"{sid}: ENV_MITIGATION(저감대책) 입력 필요")
        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(measures, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid_norm == "CH5_TRACKER":
        rows = facts.get("condition_tracker", [])
        has_any = isinstance(rows, list) and len(rows) > 0
        paras.append("협의의견(조건) 이행관리대장을 중심으로 협의내용 이행계획을 정리한다.")
        if not has_any:
            paras.append("협의조건/이행계획 입력이 없어 이행관리대장이 미완성일 수 있다. 【작성자 기입 필요】")
            todos.append(f"{sid}: ENV_MANAGEMENT(협의조건/이행관리) 입력 필요")
        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(rows, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    if sid == "CH7_CONCLUSION":
        scoping = facts.get("scoping_matrix", [])
        measures = facts.get("mitigation_measures", [])
        has_scoping = isinstance(scoping, list) and len(scoping) > 0
        has_measures = isinstance(measures, list) and len(measures) > 0

        paras.append("항목별 영향요인, 저감대책 및 이행계획을 종합하여 결론을 정리한다.")
        if not has_scoping or not has_measures:
            paras.append("평가항목/저감대책 입력이 부족하여 결론이 미완성일 수 있다. 【작성자 기입 필요】")
            if not has_scoping:
                todos.append("CH7_CONCLUSION: ENV_SCOPING 입력 필요")
            if not has_measures:
                todos.append("CH7_CONCLUSION: ENV_MITIGATION 입력 필요")

        fallback_src = _collect_source_ids_no_tbd(project)
        src = _collect_source_ids_no_tbd(scoping, measures, fallback=fallback_src)
        paras = [ensure_citation(p, src) for p in paras]
        return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)

    # Default minimal scaffold for remaining sections (LLM 권장)
    paras.append(
        f"{spec.title}은(는) 입력된 현황자료 및 계획자료를 바탕으로 작성한다. 【작성자 기입 필요】"
    )
    fallback_src = _collect_source_ids_no_tbd(facts.get("project", {}))
    src = _collect_source_ids_no_tbd(facts, fallback=fallback_src)
    paras = [ensure_citation(x, src) for x in paras]
    todos.append(f"{sid}: 섹션별 상세 입력/검토 필요")
    return SectionDraft(section_id=sid, title=spec.title, paragraphs=paras, todos=todos)


@dataclass
class WriterOptions:
    use_llm: bool = True


class ReportWriter:
    def __init__(self, sources: SourceRegistry, llm: LLMClient | None, options: WriterOptions | None = None):
        self._sources = sources
        self._llm = llm
        self._options = options or WriterOptions()

    def generate(self, case: Case) -> ReportDraft:
        sections: list[SectionDraft] = []

        scoping_by_id = {s.item_id: s for s in case.scoping_matrix}
        omission = _extract_prior_omission(case)

        for spec in SECTION_SPECS:
            # Conditional omit/exclude for item sections.
            item_id = _ITEM_SECTION_TO_ITEM_ID.get(spec.section_id)
            if item_id and omission.get("allow_omit") and item_id in set(omission.get("omit_item_ids") or []):
                draft = _omitted_section(spec, omission.get("legal_basis_text", ""))
                sections.append(draft)
                continue

            if item_id and item_id in scoping_by_id:
                item = scoping_by_id[item_id]
                if item.scoping_class == ScopingClass.EXCLUDE:
                    draft = _excluded_section(
                        spec,
                        item.item_name,
                        item.exclude_reason.t,
                        item.exclude_reason.src or ["S-TBD"],
                    )
                    sections.append(draft)
                    continue

            facts = build_facts(case, spec.section_id)
            draft = self._generate_section(spec, facts)
            draft.paragraphs = [ensure_citation(p) for p in draft.paragraphs]
            sections.append(draft)

        return ReportDraft(sections=sections)

    def _generate_section(self, spec: SectionSpec, facts: dict[str, Any]) -> SectionDraft:
        if not self._options.use_llm or self._llm is None:
            return _rule_based_section(spec, facts, sources=self._sources)

        try:
            draft = self._llm.generate_section(spec, facts)
        except Exception as e:
            fallback = _rule_based_section(spec, facts, sources=self._sources)
            fallback.todos.append(f"LLM 실패로 규칙기반 생성: {type(e).__name__}")
            return fallback

        # Post-process: ensure citations, and collect TODOs if missing.
        fixed: list[str] = []
        for p in draft.paragraphs:
            if "【작성자 기입 필요】" in p and "S-TBD" not in p:
                draft.todos.append(f"{spec.section_id}: 누락 입력 존재(본문 확인)")
            fixed.append(ensure_citation(p))
        draft.paragraphs = fixed
        if not draft.section_id:
            draft.section_id = spec.section_id
        if not draft.title:
            draft.title = spec.title
        return draft


def _spec_section_to_llm_spec(section_id: str, heading: str) -> SectionSpec:
    return SectionSpec(section_id=section_id, title=heading, requirements="")


_SPEC_SECTION_TO_ITEM_ID: dict[str, str] = {
    "CH2_TOPO": "NAT_TG",
    "CH2_ECO": "NAT_ECO",
    "CH2_WATER": "NAT_WATER",
    "CH2_AIR": "LIFE_AIR",
    "CH2_NOISE": "LIFE_NOISE",
    "CH2_ODOR": "LIFE_ODOR",
    "CH2_LANDUSE": "SOC_LANDUSE",
    "CH2_LANDSCAPE": "SOC_LANDSCAPE",
    "CH2_POP_TRAFFIC": "SOC_POP",
}


class SpecReportWriter:
    """Spec(SSOT: spec/*.yaml) 기반 섹션 생성기."""

    def __init__(
        self,
        spec: SpecBundle,
        sources: SourceRegistry,
        llm: LLMClient | None,
        options: WriterOptions | None = None,
    ) -> None:
        self._spec = spec
        self._sources = sources
        self._llm = llm
        self._options = options or WriterOptions()

    def generate(self, case: Case) -> ReportDraft:
        sections: list[SectionDraft] = []
        scoping_by_id = {s.item_id: s for s in case.scoping_matrix}
        omission = _extract_prior_omission(case)
        tables_by_id = {t.id: t for t in self._spec.tables.tables}
        figures_by_id = {f.id: f for f in self._spec.figures.figures}

        from eia_gen.services.figures.spec_figures import resolve_figure
        from eia_gen.services.tables.spec_tables import build_table

        for sec in self._spec.sections.sections:
            # section-level condition
            if sec.condition and not eval_condition(case, sec.condition):
                continue

            llm_spec = _spec_section_to_llm_spec(sec.id, sec.heading)

            # prior omission / exclude are only for item sections
            item_id = _SPEC_SECTION_TO_ITEM_ID.get(sec.id)
            if item_id and omission.get("allow_omit") and item_id in set(omission.get("omit_item_ids") or []):
                sections.append(_omitted_section(llm_spec, omission.get("legal_basis_text", "")))
                continue

            if item_id and item_id in scoping_by_id:
                item = scoping_by_id[item_id]
                if item.scoping_class == ScopingClass.EXCLUDE:
                    sections.append(
                        _excluded_section(
                            llm_spec,
                            item.item_name,
                            item.exclude_reason.t,
                            item.exclude_reason.src or ["S-TBD"],
                        )
                    )
                    continue

            facts = build_facts(case, sec.id)
            if sec.mode == "deterministic" or not self._options.use_llm or self._llm is None:
                draft = _rule_based_section(llm_spec, facts, sources=self._sources)
            else:
                draft = self._generate_section(llm_spec, facts)

            draft.paragraphs = [ensure_citation(p) for p in draft.paragraphs]

            # Attach deterministic table/figure payloads for:
            # - draft exports
            # - source_register.xlsx claim-level traceability
            for table_id in sec.outputs.tables:
                t_spec = tables_by_id.get(table_id)
                if not t_spec:
                    continue
                tdata = build_table(case, self._sources, t_spec, self._spec.tables.defaults)
                draft.tables.append(
                    TableDraft(
                        table_id=table_id,
                        caption=tdata.caption,
                        headers=tdata.headers,
                        rows=tdata.rows,
                        source_ids=tdata.source_ids,
                    )
                )
            for fig_id in sec.outputs.figures:
                f_spec = figures_by_id.get(fig_id)
                if not f_spec:
                    continue
                fdata = resolve_figure(case, f_spec)
                draft.figures.append(
                    FigureDraft(
                        figure_id=fig_id,
                        file_path=fdata.file_path,
                        caption=fdata.caption,
                        source_ids=fdata.source_ids,
                    )
                )
            sections.append(draft)

        return ReportDraft(sections=sections)

    def _generate_section(self, spec: SectionSpec, facts: dict[str, Any]) -> SectionDraft:
        try:
            return (
                self._llm.generate_section(spec, facts)
                if self._llm
                else _rule_based_section(spec, facts, sources=self._sources)
            )
        except Exception as e:
            fallback = _rule_based_section(spec, facts, sources=self._sources)
            fallback.todos.append(f"LLM 실패로 규칙기반 생성: {type(e).__name__}")
            return fallback
