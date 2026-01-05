from __future__ import annotations

import json
from pathlib import Path
import re

from docx import Document
from docx.shared import Inches, Mm
from docx.text.paragraph import Paragraph

from eia_gen.models.case import Case
from eia_gen.models.sources import SourceRegistry
from eia_gen.services.citations import format_citations
from eia_gen.services.docx.types import TableData
from eia_gen.services.draft import ReportDraft, SectionDraft
from eia_gen.services.figures.derived_evidence import guess_case_dir_from_derived_dir, record_derived_evidence
from eia_gen.services.figures.materialize import MaterializeOptions, materialize_figure_image_result

_GEN_METHOD_PAGE_RE = re.compile(r"(?i)(?:^|\b)(?:PDF_PAGE|FROM_PDF_PAGE|PAGE)\s*[:=]\s*(\d+)\b")


def _set_paragraph_text(paragraph: Paragraph, text: str) -> None:
    paragraph.clear()
    paragraph.add_run(text)


def _insert_paragraph_after(paragraph: Paragraph, text: str = "", style: str | None = None) -> Paragraph:
    # Based on python-docx internal API; safe enough for controlled insertion.
    from docx.oxml import OxmlElement
    from docx.text.paragraph import Paragraph as ParagraphCls

    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = ParagraphCls(new_p, paragraph._parent)
    if style:
        new_para.style = style
    if text:
        new_para.add_run(text)
    return new_para


def _insert_table_after(paragraph: Paragraph, headers: list[str], rows: list[list[str]]) -> None:
    doc = paragraph.part.document
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
    for r in rows:
        cells = table.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = v

    paragraph._p.addnext(table._tbl)


def _add_caption(paragraph: Paragraph, caption: str) -> Paragraph:
    try:
        paragraph.style = "Caption"
    except Exception:
        # If the style doesn't exist, keep default.
        pass
    _set_paragraph_text(paragraph, caption)
    return paragraph


def _find_asset(case: Case, asset_id: str) -> tuple[str | None, str, list[str]]:
    for a in case.assets:
        if a.asset_id == asset_id:
            return a.file_path, a.caption.text_or_placeholder(), a.source_ids
    return None, f"【첨부 필요】({asset_id})", ["S-TBD"]


def _find_asset_obj(case: Case, asset_id: str):
    for a in case.assets:
        if a.asset_id == asset_id:
            return a
    return None


def _table_parcels(case: Case) -> TableData:
    headers = ["지번", "지목", "용도지역", "면적(m2)", "비고"]
    rows: list[list[str]] = []
    for p in case.project_overview.area.parcels:
        area = "" if p.area_m2.v is None else f"{int(p.area_m2.v):,}"
        rows.append(
            [
                p.jibun.text_or_placeholder(),
                p.land_category.text_or_placeholder(),
                p.zoning.text_or_placeholder(),
                area,
                p.note.text_or_placeholder("") if p.note.t.strip() else "",
            ]
        )
    caption = "지번별 면적표"
    return TableData(caption=caption, headers=headers, rows=rows, source_ids=["S-TBD"])


def _table_facilities(case: Case) -> TableData:
    headers = ["구분", "시설명", "수량", "면적(m2)", "수용", "비고"]
    rows: list[list[str]] = []
    for f in case.project_overview.contents_scale.facilities:
        qty = "" if f.qty.v is None else f"{int(f.qty.v):,}{f.qty.u or ''}"
        area = "" if f.area_m2.v is None else f"{int(f.area_m2.v):,}"
        cap = "" if f.capacity_person.v is None else f"{int(f.capacity_person.v):,}{f.capacity_person.u or ''}"
        rows.append(
            [
                f.category.text_or_placeholder(),
                f.name.text_or_placeholder(),
                qty,
                area,
                cap,
                f.note.text_or_placeholder("") if f.note.t.strip() else "",
            ]
        )
    caption = "시설별 규모"
    return TableData(caption=caption, headers=headers, rows=rows, source_ids=["S-TBD"])


def _table_scoping(case: Case) -> TableData:
    headers = ["항목코드", "항목", "구분(중점/현황/제외)", "제외 사유", "현황조사 방법", "예측/평가 방법"]
    rows: list[list[str]] = []
    for s in case.scoping_matrix:
        rows.append(
            [
                s.item_id,
                s.item_name,
                s.category.t,
                s.exclude_reason.text_or_placeholder("") if s.exclude_reason.t.strip() else "",
                s.baseline_method.text_or_placeholder(""),
                s.prediction_method.text_or_placeholder(""),
            ]
        )
    caption = "평가항목 선정(스코핑)"
    return TableData(caption=caption, headers=headers, rows=rows, source_ids=["S-TBD"])


def _table_mitigation(case: Case) -> TableData:
    headers = ["ID", "단계", "영향(연계항목)", "대책", "관리/모니터링"]
    rows: list[list[str]] = []
    for m in case.mitigation.measures:
        rows.append(
            [
                m.measure_id,
                m.phase.text_or_placeholder(),
                ", ".join(m.related_impacts) if m.related_impacts else "",
                m.title.text_or_placeholder(),
                m.monitoring.text_or_placeholder("") if m.monitoring.t.strip() else "",
            ]
        )
    caption = "저감방안(영향-대책-관리)"
    return TableData(caption=caption, headers=headers, rows=rows, source_ids=["S-TBD"])


def _table_condition_tracker(case: Case) -> TableData:
    headers = ["협의의견/조건", "조치(대책ID)", "시기", "증빙", "담당(해당 시)"]
    rows: list[list[str]] = []
    for x in case.management_plan.implementation_register:
        rows.append(
            [
                x.item.text_or_placeholder(),
                x.measure_id.text_or_placeholder(),
                x.when.text_or_placeholder(),
                x.evidence.text_or_placeholder(),
                x.responsible.text_or_placeholder("") if x.responsible.t.strip() else "",
            ]
        )
    caption = "협의조건 이행관리대장"
    return TableData(caption=caption, headers=headers, rows=rows, source_ids=["S-TBD"])


def _table_source_register(sources: SourceRegistry) -> TableData:
    headers = ["Source ID", "유형", "자료명", "제공기관", "기준일/기간", "파일/링크", "비고"]
    rows: list[list[str]] = []
    for s in sources.sources:
        rows.append(
            [
                s.id,
                s.type or "",
                s.title or "",
                s.publisher or "",
                s.period or s.date or "",
                s.file_path or s.file_or_url or "",
                s.note or s.notes or "",
            ]
        )
    caption = "출처/근거 관리표(Source Register)"
    return TableData(caption=caption, headers=headers, rows=rows)


def _assets_by_types(case: Case, types: set[str]) -> list[str]:
    return [a.asset_id for a in case.assets if a.type in types]


def _resolve_existing_path(file_path: str | None, search_dirs: list[Path]) -> Path | None:
    fp = (file_path or "").strip()
    if not fp:
        return None
    # Allow optional "#page=N" / "?page=N" fragments in file_path (PDF page hint).
    fp = re.sub(r"(?i)(?:[#?@]page=)\\d+\\b", "", fp)
    p = Path(fp).expanduser()
    if p.is_absolute():
        return p if p.exists() else None
    for base in search_dirs:
        cand = (base / p).expanduser()
        if cand.exists():
            return cand
    return p if p.exists() else None


def _asset_width_mm(asset) -> float | None:
    try:
        v = getattr(asset, "width_mm", None)
        if v is None or v == "":
            return None
        return float(v)
    except Exception:
        return None


def _insert_asset_picture(
    p: Paragraph,
    *,
    case: Case,
    asset,
    resolved_path: Path,
    derived_dir: Path,
    report_anchor: str,
) -> None:
    width_mm = _asset_width_mm(asset)
    width = Mm(width_mm) if (width_mm and width_mm > 0) else Inches(6.0)
    width_mm_eff = None
    try:
        width_mm_eff = float(width.inches) * 25.4
    except Exception:
        width_mm_eff = width_mm
    crop = (getattr(asset, "crop", None) or "").strip() or None
    gen_method = (getattr(asset, "gen_method", None) or "").strip() or None
    asset_type = (getattr(asset, "type", None) or "").strip() or None
    raw_fp = (getattr(asset, "file_path", None) or "").strip()
    # Allow '#page=N' style hints in file_path for PDF selection even when resolved_path strips it.
    if resolved_path.suffix.lower() == ".pdf" and raw_fp and not _GEN_METHOD_PAGE_RE.search(gen_method or ""):
        m = re.search(r"(?i)(?:[#?@]page=)(\d+)\b", raw_fp)
        if m:
            try:
                page = max(1, int(m.group(1)))
                gen_method = f"{gen_method} PDF_PAGE:{page}".strip() if gen_method else f"PDF_PAGE:{page}"
            except Exception:
                pass

    # Apply v2 figure controls (PDF rasterize / crop / resize) when requested.
    try:
        if resolved_path.suffix.lower() == ".pdf" or crop or (width_mm and width_mm > 0) or resolved_path.suffix.lower() in {
            ".png",
            ".jpg",
            ".jpeg",
        }:
            mat_dir = derived_dir / "figures" / "_materialized"
            mat_res = materialize_figure_image_result(
                resolved_path,
                MaterializeOptions(
                    out_dir=mat_dir,
                    fig_id=getattr(asset, "asset_id", "FIG"),
                    gen_method=gen_method,
                    crop=crop,
                    width_mm=width_mm_eff,
                    asset_type=asset_type,
                ),
                include_meta=True,
            )
            mat = mat_res.path

            case_dir = guess_case_dir_from_derived_dir(derived_dir)
            src_ids = [str(s) for s in (getattr(asset, "source_ids", None) or []) if str(s).strip()]
            if not src_ids:
                src_ids = ["S-TBD"]

            note_obj = mat_res.meta if isinstance(getattr(mat_res, "meta", None), dict) else {"kind": "MATERIALIZE"}
            if isinstance(note_obj, dict) and case_dir is not None:
                note_obj = dict(note_obj)
                try:
                    note_obj["src_rel_path"] = str(resolved_path.resolve().relative_to(case_dir.resolve())).replace(
                        "\\", "/"
                    )
                except Exception:
                    note_obj["src_rel_path"] = resolved_path.name
            note = json.dumps(note_obj, ensure_ascii=False, sort_keys=True)

            evidence_type = "derived_png" if mat.suffix.lower() == ".png" else "derived_jpg"
            record_derived_evidence(
                case,
                derived_path=mat,
                related_fig_id=str(getattr(asset, "asset_id", "") or report_anchor).strip() or report_anchor,
                report_anchor=report_anchor,
                src_ids=src_ids,
                evidence_type=evidence_type,
                title=str(getattr(getattr(asset, "caption", None), "t", "") or report_anchor).strip() or report_anchor,
                note=note,
                pdf_page=getattr(mat_res, "pdf_page", None),
                pdf_page_source=getattr(mat_res, "pdf_page_source", None),
                used_in=report_anchor,
                case_dir=case_dir,
            )
            p.add_run().add_picture(str(mat), width=width)
        else:
            p.add_run().add_picture(str(resolved_path), width=width)
    except Exception:
        p.add_run("【첨부 오류】이미지 삽입 실패")


def _add_assets_section(doc: Document, case: Case, search_dirs: list[Path], *, derived_dir: Path) -> None:
    if not case.assets:
        doc.add_paragraph("【첨부 필요】도면/사진/조사표(assets) 미첨부")
        return
    for a in case.assets:
        p = doc.add_paragraph()
        resolved = _resolve_existing_path(a.file_path, search_dirs)
        if resolved is not None:
            _insert_asset_picture(
                p,
                case=case,
                asset=a,
                resolved_path=resolved,
                derived_dir=derived_dir,
                report_anchor=str(getattr(a, "asset_id", "") or "").strip() or "ASSET",
            )
        else:
            p.add_run("【첨부 필요】")
        doc.add_paragraph(f"{a.caption.text_or_placeholder()} {format_citations(a.source_ids)}")


def build_docx(
    case: Case,
    sources: SourceRegistry,
    draft: ReportDraft,
    out_path: str | Path,
    template_path: str | Path | None = None,
    spec_dir: str | Path | None = None,
    use_template_map: bool = False,
    asset_base_dir: str | Path | None = None,
) -> None:
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    search_dirs: list[Path] = []
    if asset_base_dir:
        search_dirs.append(Path(asset_base_dir).expanduser().resolve())
    # Also search relative to where we're writing the output.
    search_dirs.append(out_path.parent.resolve())
    # And cwd as a last resort.
    search_dirs.append(Path.cwd().resolve())

    spec = None
    spec_path = Path(spec_dir).expanduser() if spec_dir else None
    if spec_path and spec_path.exists():
        from eia_gen.spec.load import load_spec_bundle

        spec = load_spec_bundle(spec_path)

    derived_dir = out_path.parent.resolve()
    if asset_base_dir:
        cand = Path(asset_base_dir).expanduser().resolve() / "attachments" / "derived"
        derived_dir = cand
    derived_dir.mkdir(parents=True, exist_ok=True)

    resolved_template: Path | None = None
    if template_path:
        cand = Path(template_path).expanduser()
        resolved_template = cand if cand.exists() else None
    elif use_template_map and spec is not None:
        # Use spec-provided default template when not explicitly passed.
        tf = (spec.template_map.template_file or "").strip()
        if tf:
            cand = Path(tf).expanduser()
            if not cand.is_absolute() and spec_path is not None:
                cand = (spec_path.parent / cand).expanduser()
            resolved_template = cand if cand.exists() else None

    if resolved_template is not None:
        if use_template_map and spec is not None:
            from eia_gen.services.docx.spec_renderer import render_template_docx

            render_template_docx(
                case=case,
                sources=sources,
                draft=draft,
                spec=spec,
                template_path=resolved_template,
                out_path=out_path,
                asset_search_dirs=search_dirs,
                derived_dir=derived_dir,
            )
            return

        doc = Document(str(resolved_template))
        _fill_template(doc, case, sources, draft, search_dirs=search_dirs, derived_dir=derived_dir)
        doc.save(str(out_path))
        return

    if spec is not None:
        from eia_gen.services.docx.spec_renderer import render_builtin_docx

        render_builtin_docx(
            case=case,
            sources=sources,
            draft=draft,
            spec=spec,
            out_path=out_path,
            asset_search_dirs=search_dirs,
            derived_dir=derived_dir,
        )
        return

    doc = Document()
    _build_builtin(doc, case, sources, draft)
    doc.save(str(out_path))


def _build_builtin(doc: Document, case: Case, sources: SourceRegistry, draft: ReportDraft) -> None:
    # Cover (very minimal)
    doc.add_paragraph(case.cover.project_name.text_or_placeholder("소규모환경영향평가서"))
    doc.add_paragraph(case.project_overview.location.address.text_or_placeholder(""))
    doc.add_paragraph(case.cover.submit_date.text_or_placeholder(""))
    doc.add_page_break()

    section_by_id: dict[str, SectionDraft] = {s.section_id: s for s in draft.sections}

    def add_section(section_id: str, heading_level: int = 1) -> None:
        s = section_by_id.get(section_id)
        if not s:
            return
        doc.add_heading(s.title, level=heading_level)
        for p in s.paragraphs:
            doc.add_paragraph(p)
        if s.todos:
            doc.add_paragraph("TODO: " + "; ".join(s.todos))

    add_section("CH0_SUMMARY", heading_level=1)
    add_section("CH1_OVERVIEW", heading_level=1)

    # Tables & key figures for CH1
    td = _table_parcels(case)
    doc.add_paragraph(f"표 1-1. {td.caption} {format_citations(td.source_ids)}")
    _add_table_end(doc, td)
    td = _table_facilities(case)
    doc.add_paragraph(f"표 1-2. {td.caption} {format_citations(td.source_ids)}")
    _add_table_end(doc, td)

    add_section("CH1_PERMITS", heading_level=1)
    add_section("CH2_METHOD", heading_level=1)

    # Chapter 2 subsections
    add_section("CH2_NAT_TG", heading_level=2)
    add_section("CH2_NAT_ECO", heading_level=2)
    add_section("CH2_NAT_WATER", heading_level=2)
    add_section("CH2_LIFE_AIR", heading_level=2)
    add_section("CH2_LIFE_NOISE", heading_level=2)
    add_section("CH2_LIFE_ODOR", heading_level=2)
    add_section("CH2_SOC_LANDUSE", heading_level=2)
    add_section("CH2_SOC_LANDSCAPE", heading_level=2)
    add_section("CH2_SOC_POP", heading_level=2)

    # Chapter 3
    doc.add_heading("3.1 평가항목 선정(스코핑)", level=2)
    td = _table_scoping(case)
    doc.add_paragraph(f"표 3-1. {td.caption} {format_citations(td.source_ids)}")
    _add_table_end(doc, td)
    add_section("CH3_CONSTRUCTION", heading_level=1)
    add_section("CH3_OPERATION", heading_level=1)

    # Chapter 4
    add_section("CH4_TEXT", heading_level=1)
    td = _table_mitigation(case)
    doc.add_paragraph(f"표 4-1. {td.caption} {format_citations(td.source_ids)}")
    _add_table_end(doc, td)

    # Chapter 5
    add_section("CH5_TEXT", heading_level=1)
    td = _table_condition_tracker(case)
    doc.add_paragraph(f"표 5-1. {td.caption} {format_citations(td.source_ids)}")
    _add_table_end(doc, td)

    # Appendices
    doc.add_page_break()
    doc.add_heading("부록", level=1)
    sr = _table_source_register(sources)
    doc.add_heading(sr.caption, level=2)
    _add_table_end(doc, sr)

    doc.add_heading("부록-2. 첨부자료(도면/사진) 목록", level=2)
    _add_assets_section(doc, case, search_dirs=[Path.cwd().resolve()], derived_dir=Path.cwd().resolve())


def _add_table_end(doc: Document, td: TableData) -> None:
    table = doc.add_table(rows=1, cols=len(td.headers))
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    for i, h in enumerate(td.headers):
        hdr[i].text = h
    for r in td.rows:
        row = table.add_row().cells
        for i, v in enumerate(r):
            row[i].text = v


def _fill_template(
    doc: Document,
    case: Case,
    sources: SourceRegistry,
    draft: ReportDraft,
    *,
    search_dirs: list[Path],
    derived_dir: Path,
) -> None:
    # Anchor-based replacement.
    section_by_id: dict[str, SectionDraft] = {s.section_id: s for s in draft.sections}

    def replace_block(p: Paragraph, section_id: str) -> None:
        s = section_by_id.get(section_id)
        if not s:
            _set_paragraph_text(p, f"【작성자 기입 필요】(섹션 누락: {section_id})")
            return
        if not s.paragraphs:
            _set_paragraph_text(p, "【작성자 기입 필요】")
            return
        _set_paragraph_text(p, s.paragraphs[0])
        cur = p
        for extra in s.paragraphs[1:]:
            cur = _insert_paragraph_after(cur, extra)

    def replace_table(p: Paragraph, table_id: str) -> None:
        if table_id == "PARCELS":
            td = _table_parcels(case)
        elif table_id == "FACILITIES":
            td = _table_facilities(case)
        elif table_id == "SCOPING":
            td = _table_scoping(case)
        elif table_id == "MITIGATION_PLAN":
            td = _table_mitigation(case)
        elif table_id == "CONDITION_TRACKER":
            td = _table_condition_tracker(case)
        elif table_id in {"SOURCE_REGISTER", "SOURCE"}:
            td = _table_source_register(sources)
        else:
            _set_paragraph_text(p, f"【작성자 기입 필요】(알 수 없는 표 앵커: {table_id})")
            return

        _add_caption(p, f"{td.caption} {format_citations(td.source_ids)}" if td.source_ids else td.caption)
        _insert_table_after(p, td.headers, td.rows)

    def replace_fig(p: Paragraph, asset_id: str) -> None:
        file_path, caption, src_ids = _find_asset(case, asset_id)
        asset = _find_asset_obj(case, asset_id)
        p.clear()
        resolved = _resolve_existing_path(file_path, search_dirs)
        if resolved is not None:
            if asset is not None:
                _insert_asset_picture(
                    p,
                    case=case,
                    asset=asset,
                    resolved_path=resolved,
                    derived_dir=derived_dir,
                    report_anchor=asset_id,
                )
            else:
                try:
                    p.add_run().add_picture(str(resolved), width=Inches(6.0))
                except Exception:
                    p.add_run("【첨부 오류】이미지 삽입 실패")
        else:
            p.add_run("【첨부 필요】")
        _insert_paragraph_after(p, f"{caption} {format_citations(src_ids)}")

    for p in list(doc.paragraphs):
        txt = p.text.strip()
        if not (txt.startswith("[[") and txt.endswith("]]")):
            continue
        inner = txt[2:-2]
        if ":" not in inner:
            continue
        kind, ident = inner.split(":", 1)
        kind = kind.strip().upper()
        ident = ident.strip()
        if kind == "BLOCK":
            replace_block(p, ident)
        elif kind == "TABLE":
            replace_table(p, ident)
        elif kind == "FIG":
            replace_fig(p, ident)
