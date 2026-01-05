from __future__ import annotations

from collections import defaultdict
from copy import deepcopy
from dataclasses import replace
from datetime import datetime
import json
import os
from pathlib import Path
import re

from docx import Document
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.table import CT_Tbl
from docx.shared import Inches, Mm
from docx.table import Table
from docx.text.paragraph import Paragraph

from eia_gen.config import settings
from eia_gen.models.case import Case
from eia_gen.models.sources import SourceRegistry
from eia_gen.services.citations import extract_citation_ids, ensure_citation, format_citations, strip_citations
from eia_gen.services.conditions import eval_condition
from eia_gen.services.draft import ReportDraft, SectionDraft
from eia_gen.services.figures.derived_evidence import (
    guess_case_dir_from_derived_dir,
    record_derived_evidence,
)
from eia_gen.services.figures.spec_figures import is_required as is_required_figure, resolve_figure
from eia_gen.services.figures.materialize import (
    MaterializeOptions,
    materialize_figure_image_result,
    resolve_source_path,
)
from eia_gen.services.tables.spec_tables import build_table
from eia_gen.spec.models import SpecBundle


class _Numbering:
    def __init__(self, table_style: str, figure_style: str) -> None:
        self._table_style = table_style
        self._figure_style = figure_style
        self._table_counter: dict[int, int] = defaultdict(int)
        self._figure_counter: dict[int, int] = defaultdict(int)

    def next_table(self, chapter: int, fixed_label: str | None = None) -> str:
        if fixed_label:
            return self._table_style.format(ch=chapter, n=0, label=fixed_label)
        self._table_counter[chapter] += 1
        auto_label = f"{chapter}-{self._table_counter[chapter]}"
        return self._table_style.format(ch=chapter, n=self._table_counter[chapter], label=auto_label)

    def next_figure(self, chapter: int, fixed_label: str | None = None) -> str:
        if fixed_label:
            return self._figure_style.format(ch=chapter, n=0, label=fixed_label)
        self._figure_counter[chapter] += 1
        auto_label = f"{chapter}-{self._figure_counter[chapter]}"
        return self._figure_style.format(ch=chapter, n=self._figure_counter[chapter], label=auto_label)


def _docx_text(text: str) -> str:
    return text if settings.docx_render_citations else strip_citations(text)


def _write_placeholder_png(out_path: Path, *, title: str, lines: list[str]) -> None:
    from PIL import Image, ImageDraw, ImageFont

    out_path.parent.mkdir(parents=True, exist_ok=True)

    w, h = 1600, 1200
    img = Image.new("RGB", (w, h), (245, 245, 245))
    d = ImageDraw.Draw(img)
    font = ImageFont.load_default()

    y = 32
    d.text((32, y), str(title), font=font, fill=(0, 0, 0))
    y += 28
    for line in lines:
        d.text((32, y), str(line), font=font, fill=(0, 0, 0))
        y += 20

    img.save(out_path, format="PNG", optimize=True)


def _clamp_picture_width(doc: Document, width):
    try:
        if not getattr(doc, "sections", None):
            return width
        sec = doc.sections[0]
        max_width = sec.page_width - sec.left_margin - sec.right_margin
        if width and max_width and width > max_width:
            return max_width
    except Exception:
        return width
    return width


def _length_to_mm(width) -> float | None:
    if width is None:
        return None
    # python-docx Length supports .inches/.mm; prefer inches for compatibility.
    try:
        return float(width.inches) * 25.4
    except Exception:
        try:
            return float(width.mm)
        except Exception:
            return None


def _iter_paragraphs_in_table(table: Table) -> list[Paragraph]:
    out: list[Paragraph] = []
    for row in table.rows:
        for cell in row.cells:
            out.extend(_iter_paragraphs_in_cell(cell))
    return out


def _iter_paragraphs_in_cell(cell) -> list[Paragraph]:
    out: list[Paragraph] = []
    out.extend(list(getattr(cell, "paragraphs", []) or []))
    for t in list(getattr(cell, "tables", []) or []):
        out.extend(_iter_paragraphs_in_table(t))
    return out


def _iter_all_paragraphs(doc: Document) -> list[Paragraph]:
    out: list[Paragraph] = []
    out.extend(list(getattr(doc, "paragraphs", []) or []))
    for t in list(getattr(doc, "tables", []) or []):
        out.extend(_iter_paragraphs_in_table(t))
    return out


def _set_paragraph_text(paragraph: Paragraph, text: str, style: str | None = None) -> None:
    if style:
        try:
            paragraph.style = style
        except Exception:
            pass
    # Preserve run-level formatting when possible (template fidelity).
    # Clearing a paragraph recreates runs with default formatting, which
    # can cause visible style drift vs. the template/PDF sample.
    if paragraph.runs:
        first = paragraph.runs[0]
        for r in list(paragraph.runs)[1:]:
            try:
                r._r.getparent().remove(r._r)
            except Exception:
                pass
        first.text = text
        return

    paragraph.clear()
    paragraph.add_run(text)


def _insert_paragraph_after_element(
    paragraph: Paragraph,
    element,
    text: str,
    *,
    style: str | None = None,
    clone_from: Paragraph | None = None,
) -> Paragraph:
    from docx.text.paragraph import Paragraph as ParagraphCls

    new_p = OxmlElement("w:p")
    if clone_from is not None and getattr(clone_from._p, "pPr", None) is not None:
        try:
            new_p.append(deepcopy(clone_from._p.pPr))
        except Exception:
            pass
    element.addnext(new_p)
    p = ParagraphCls(new_p, paragraph._parent)
    if style:
        try:
            p.style = style
        except Exception:
            pass
    run = p.add_run(text)
    # Try to preserve run-level formatting when cloning paragraphs.
    if clone_from is not None and clone_from.runs:
        try:
            src_rpr = clone_from.runs[0]._r.rPr
            if src_rpr is not None:
                if run._r.rPr is not None:
                    run._r.remove(run._r.rPr)
                run._r.insert(0, deepcopy(src_rpr))
        except Exception:
            pass
    return p


def _insert_paragraph_after(
    paragraph: Paragraph,
    text: str,
    style: str | None = None,
    *,
    clone_from: Paragraph | None = None,
) -> Paragraph:
    return _insert_paragraph_after_element(paragraph, paragraph._p, text, style=style, clone_from=clone_from)


def _insert_empty_paragraph_after(paragraph: Paragraph, *, clone_from: Paragraph | None = None) -> Paragraph:
    from docx.text.paragraph import Paragraph as ParagraphCls

    new_p = OxmlElement("w:p")
    if clone_from is not None and getattr(clone_from._p, "pPr", None) is not None:
        try:
            new_p.append(deepcopy(clone_from._p.pPr))
        except Exception:
            pass
    paragraph._p.addnext(new_p)
    return ParagraphCls(new_p, paragraph._parent)


def _parse_pdf_page_directive(text: str) -> dict[str, object] | None:
    """Parse section-embedded PDF_PAGE directive.

    Syntax:
      [[PDF_PAGE:<file_path>|page=3|width_mm=170|crop=AUTO|dpi=200|break=1|auth=SSOT_SAMPLE]]
    """
    raw = (text or "").strip()
    src_ids = extract_citation_ids(raw)

    # Writer may append inline citations (e.g. `〔SRC:...〕`). Strip them so the directive remains
    # parseable while keeping normal paragraphs traceable.
    s = strip_citations(raw).strip()
    if not s.startswith("[[PDF_PAGE:") or not s.endswith("]]"):
        return None
    inner = s[len("[[PDF_PAGE:") : -2]
    parts = [p.strip() for p in inner.split("|") if p.strip()]
    if not parts:
        return None

    file_path = parts[0]
    params: dict[str, str] = {}
    for token in parts[1:]:
        if "=" not in token:
            continue
        k, v = token.split("=", 1)
        params[k.strip().lower()] = v.strip()

    ins_id = (params.get("id") or "").strip() or None
    raw_auth = (params.get("auth") or params.get("authenticity") or "").strip()
    auth = raw_auth.upper().replace("-", "_").replace(" ", "_") if raw_auth else ""
    auth = re.sub(r"[^A-Z0-9_]+", "_", auth).strip("_")
    auth = auth or None

    def _as_int(key: str, default: int) -> int:
        try:
            return int(params.get(key, default))
        except Exception:
            return default

    def _as_float(key: str, default: float) -> float:
        try:
            return float(params.get(key, default))
        except Exception:
            return default

    def _as_bool(key: str, default: bool) -> bool:
        v = (params.get(key) or "").strip().lower()
        if not v:
            return default
        return v in {"1", "true", "t", "y", "yes", "on"}

    return {
        "id": ins_id,
        "src_ids": src_ids,
        "file_path": file_path,
        "page": max(1, _as_int("page", 1)),
        "width_mm": _as_float("width_mm", 170.0),
        "crop": (params.get("crop") or "").strip() or None,
        "dpi": max(72, _as_int("dpi", 200)),
        "break": _as_bool("break", True),
        "auth": auth,
    }


def _insert_pdf_page_image(
    paragraph: Paragraph,
    *,
    file_path: str,
    page: int,
    width_mm: float,
    crop: str | None,
    dpi: int,
    auth: str | None = None,
    src_ids: list[str] | None,
    note_extra: dict[str, object] | None = None,
    case: Case,
    case_dir: Path | None,
    derived_dir: Path,
    asset_search_dirs: list[Path] | None,
    cache_key: str,
) -> tuple[Path, Path, int | None, str | None] | None:
    paragraph.clear()

    resolved = resolve_source_path(file_path, asset_search_dirs=asset_search_dirs)
    if resolved is None:
        paragraph.add_run("【첨부 필요】")
        return None

    try:
        width = Mm(float(width_mm)) if (width_mm and width_mm > 0) else Inches(6.0)
        width = _clamp_picture_width(paragraph.part.document, width)
        width_mm_eff = _length_to_mm(width) or float(width_mm)
        mat_dir = derived_dir / "figures" / "_materialized"
        gen_method = f"PDF_PAGE:{int(page)}"
        if auth:
            gen_method = f"{gen_method} AUTHENTICITY:{str(auth).strip()}"
        mat_res = materialize_figure_image_result(
            resolved,
            MaterializeOptions(
                out_dir=mat_dir,
                fig_id=cache_key,
                gen_method=gen_method,
                crop=crop,
                width_mm=width_mm_eff,
                target_dpi=dpi,
            ),
            include_meta=True,
        )
        mat = mat_res.path
        evidence_type = "derived_png" if mat.suffix.lower() == ".png" else "derived_jpg"
        note_obj = mat_res.meta if isinstance(getattr(mat_res, "meta", None), dict) else {"kind": "MATERIALIZE"}
        if isinstance(note_obj, dict) and case_dir is not None:
            note_obj = dict(note_obj)
            try:
                note_obj["src_rel_path"] = str(resolved.resolve().relative_to(case_dir.resolve())).replace("\\", "/")
            except Exception:
                note_obj["src_rel_path"] = resolved.name
        if isinstance(note_obj, dict) and isinstance(note_extra, dict) and note_extra:
            note_obj = dict(note_obj)
            for k, v in note_extra.items():
                if k not in note_obj:
                    note_obj[k] = v
        note = json.dumps(note_obj, ensure_ascii=False, sort_keys=True)
        record_derived_evidence(
            case,
            derived_path=mat,
            related_fig_id=cache_key,
            report_anchor=cache_key,
            src_ids=src_ids or [],
            evidence_type=evidence_type,
            title=f"{resolved.name} p{int(page)}",
            note=note,
            pdf_page=mat_res.pdf_page,
            pdf_page_source=mat_res.pdf_page_source,
            used_in=cache_key,
            case_dir=case_dir,
        )
        paragraph.add_run().add_picture(str(mat), width=width)
        return (resolved, mat, mat_res.pdf_page, mat_res.pdf_page_source)
    except ImportError as e:
        # Fallback: create a clearly marked placeholder PNG so doc generation and traceability can proceed.
        placeholder = mat_dir / f"{cache_key}__PLACEHOLDER__NO_PYMUPDF.png"
        _write_placeholder_png(
            placeholder,
            title="PDF rasterize unavailable (placeholder)",
            lines=[
                "missing dependency: PyMuPDF (fitz)",
                f"src: {resolved.name}",
                f"page: {int(page)}",
                f"error: {type(e).__name__}",
            ],
        )

        note_obj: dict[str, object] = {
            "kind": "PDF_PAGE_PLACEHOLDER",
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "missing_dependency": "PyMuPDF(fitz)",
            "src_name": resolved.name,
            "page_1based": int(page),
            "width_mm": float(width_mm_eff),
            "dpi": int(dpi),
            "crop": crop or "",
            "error": f"{type(e).__name__}: {e}",
        }
        if isinstance(note_extra, dict) and note_extra:
            for k, v in note_extra.items():
                if k not in note_obj:
                    note_obj[k] = v
        note = json.dumps(note_obj, ensure_ascii=False, sort_keys=True)

        record_derived_evidence(
            case,
            derived_path=placeholder,
            related_fig_id=cache_key,
            report_anchor=cache_key,
            src_ids=src_ids or [],
            evidence_type="derived_png",
            title=f"{resolved.name} p{int(page)} (placeholder)",
            note=note,
            pdf_page=int(page),
            pdf_page_source="explicit",
            used_in=cache_key,
            case_dir=case_dir,
        )
        paragraph.add_run().add_picture(str(placeholder), width=width)
        return (resolved, placeholder, int(page), "explicit")
    except Exception:
        paragraph.add_run("【첨부 오류】이미지 삽입 실패")
        return None


def _insert_page_break(paragraph: Paragraph) -> None:
    try:
        paragraph.add_run().add_break(WD_BREAK.PAGE)
    except Exception:
        _insert_paragraph_after(paragraph, "")


def _insert_table_after(paragraph: Paragraph, headers: list[str], rows: list[list[str]]):
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
    return table


_GEN_METHOD_REFERENCE_RE = re.compile(r"(?i)(?:^|\b)AUTHENTICITY\s*[:=]\s*REFERENCE\b")


def _needs_materialize(src_path: Path, fdata: FigureData) -> bool:
    suf = src_path.suffix.lower()
    if suf == ".pdf":
        return True
    if suf not in {".png", ".jpg", ".jpeg"}:
        return False
    if fdata.crop:
        return True
    if fdata.width_mm and fdata.width_mm > 0:
        return True
    if fdata.gen_method and _GEN_METHOD_REFERENCE_RE.search(fdata.gen_method):
        return True
    # Efficiency: apply SSOT materialize policy (format/size) even when no explicit crop/width is set.
    return True


def _insert_section_block(
    paragraph: Paragraph,
    draft: SectionDraft | None,
    *,
    case: Case,
    case_dir: Path | None,
    derived_dir: Path,
    asset_search_dirs: list[Path] | None,
    block_id: str,
) -> None:
    if block_id == "APPENDIX_INSERTS":
        raw_rows = (getattr(case, "model_extra", None) or {}).get("appendix_inserts")
        rows: list[dict[str, object]] = []
        if isinstance(raw_rows, list):
            rows = [r for r in raw_rows if isinstance(r, dict)]

        def _sort_key(r: dict[str, object]) -> tuple[int, str]:
            try:
                order = int(r.get("order") or 0)
            except Exception:
                order = 0
            return (order, str(r.get("ins_id") or ""))

        rows.sort(key=_sort_key)

        if not rows:
            _set_paragraph_text(paragraph, "")
            return

        width_mm_default = 170.0
        try:
            dpi_default = int(os.getenv("EIA_GEN_SSOT_PDF_DPI", "120"))
        except Exception:
            dpi_default = 120
        dpi_default = max(72, dpi_default)

        cur = paragraph
        for idx, r in enumerate(rows):
            br = idx < len(rows) - 1

            ins_id = str(r.get("ins_id") or "").strip() or f"INS-{idx + 1:04d}"
            safe_id = re.sub(r"[^A-Za-z0-9_-]+", "_", ins_id)[:64]

            ins_pdf = str(r.get("file_path") or "").strip()
            try:
                page = int(r.get("page") or 0)
            except Exception:
                page = 0
            page = max(1, page)

            try:
                width_mm = float(r.get("width_mm") or width_mm_default)
            except Exception:
                width_mm = float(width_mm_default)
            width_mm = float(width_mm_default) if width_mm <= 0 else width_mm

            try:
                dpi = int(r.get("dpi") or dpi_default)
            except Exception:
                dpi = int(dpi_default)
            dpi = max(72, dpi)

            crop = str(r.get("crop") or "").strip() or None
            caption = str(r.get("caption") or "").strip()
            note_text = str(r.get("note") or "").strip()

            src_ids = r.get("src_ids")
            if not isinstance(src_ids, list) or not all(isinstance(x, str) for x in src_ids):
                src_ids = ["S-TBD"]

            note_extra: dict[str, object] = {"insert_id": ins_id}
            if note_text:
                note_extra["insert_note"] = note_text

            cache_key = f"{block_id}_{safe_id}_p{page:03d}"

            if caption:
                _set_paragraph_text(cur, _docx_text(ensure_citation(caption, src_ids)))
                try:
                    cur.paragraph_format.keep_with_next = True
                except Exception:
                    pass
                cur = _insert_empty_paragraph_after(cur, clone_from=cur)

            _insert_pdf_page_image(
                cur,
                file_path=ins_pdf,
                page=page,
                width_mm=width_mm,
                crop=crop,
                dpi=dpi,
                src_ids=src_ids,
                note_extra=note_extra,
                case=case,
                case_dir=case_dir,
                derived_dir=derived_dir,
                asset_search_dirs=asset_search_dirs,
                cache_key=cache_key,
            )
            if br:
                _insert_page_break(cur)
                cur = _insert_empty_paragraph_after(cur, clone_from=cur)
            else:
                try:
                    if cur._p.getnext() is not None:
                        _insert_page_break(cur)
                except Exception:
                    pass

        return

    if draft is None:
        _set_paragraph_text(paragraph, "【작성자 기입 필요】(섹션 초안 없음)")
        return
    if not draft.paragraphs:
        _set_paragraph_text(paragraph, "【작성자 기입 필요】")
        return

    cur = paragraph
    for idx, raw in enumerate(draft.paragraphs):
        directive = _parse_pdf_page_directive(raw)
        if directive is not None:
            if idx != 0:
                cur = _insert_empty_paragraph_after(cur, clone_from=cur)

            page = int(directive["page"])
            raw_id = str(directive.get("id") or "").strip()
            safe_id = re.sub(r"[^A-Za-z0-9_-]+", "_", raw_id)[:64] if raw_id else ""
            cache_key = f"{block_id}_{safe_id}_p{page:03d}" if safe_id else f"{block_id}_p{page:03d}"
            src_ids = directive.get("src_ids")
            if not isinstance(src_ids, list) or not all(isinstance(x, str) for x in src_ids):
                src_ids = []
            auth = directive.get("auth")
            if not isinstance(auth, str) or not auth.strip():
                auth = None
            _insert_pdf_page_image(
                cur,
                file_path=str(directive["file_path"]),
                page=page,
                width_mm=float(directive["width_mm"]),
                crop=directive["crop"] if isinstance(directive["crop"], str) else None,
                dpi=int(directive["dpi"]),
                auth=auth,
                src_ids=src_ids,
                case=case,
                case_dir=case_dir,
                derived_dir=derived_dir,
                asset_search_dirs=asset_search_dirs,
                cache_key=cache_key,
            )
            if bool(directive["break"]):
                _insert_page_break(cur)
            continue

        text = _docx_text(raw)
        next_raw = draft.paragraphs[idx + 1] if idx + 1 < len(draft.paragraphs) else None
        next_is_pdf_page = bool(next_raw and _parse_pdf_page_directive(str(next_raw)) is not None)
        if idx == 0:
            _set_paragraph_text(cur, text)
            if next_is_pdf_page:
                try:
                    cur.paragraph_format.keep_with_next = True
                except Exception:
                    pass
            continue
        cur = _insert_paragraph_after(cur, text, clone_from=cur)
        if next_is_pdf_page:
            try:
                cur.paragraph_format.keep_with_next = True
            except Exception:
                pass


def _find_table_after_paragraph(paragraph: Paragraph) -> Table | None:
    nxt = paragraph._p.getnext()
    if isinstance(nxt, CT_Tbl):
        return Table(nxt, paragraph._parent)
    return None


def _set_cell_text(cell, text: str) -> None:
    # Preserve cell/paragraph formatting as much as possible.
    if not cell.paragraphs:
        cell.text = text
        return
    _set_paragraph_text(cell.paragraphs[0], text)


def _clone_table_row_in_place(table: Table, template_row_idx: int) -> None:
    # Deep-copy the row XML to preserve widths/borders/shading.
    template_row = table.rows[template_row_idx]
    table._tbl.append(deepcopy(template_row._tr))


def _delete_table_row_in_place(table: Table, row_idx: int) -> None:
    tr = table.rows[row_idx]._tr
    tr.getparent().remove(tr)


def _fill_table_in_place(table: Table, headers: list[str], rows: list[list[str]]) -> bool:
    if not table.rows:
        return False

    # Basic column count sanity: template must have enough columns.
    if len(table.rows[0].cells) < len(headers):
        return False

    # Fill header row.
    for j, h in enumerate(headers):
        _set_cell_text(table.rows[0].cells[j], h)

    desired_data_rows = len(rows)
    current_data_rows = max(0, len(table.rows) - 1)

    # Add rows (clone) or remove rows to match.
    if desired_data_rows > current_data_rows:
        template_row_idx = 1 if len(table.rows) > 1 else 0
        for _ in range(desired_data_rows - current_data_rows):
            _clone_table_row_in_place(table, template_row_idx=template_row_idx)
    elif desired_data_rows < current_data_rows:
        # Remove from bottom to preserve template header & early rows.
        for _ in range(current_data_rows - desired_data_rows):
            _delete_table_row_in_place(table, row_idx=len(table.rows) - 1)

    # Fill data rows.
    for i, r in enumerate(rows, start=1):
        # Template may have more columns; only fill what we have.
        for j in range(min(len(headers), len(table.rows[i].cells))):
            _set_cell_text(table.rows[i].cells[j], r[j] if j < len(r) else "")
        # Clear remaining template cells (if any) to avoid placeholder bleed.
        for j in range(len(headers), len(table.rows[i].cells)):
            _set_cell_text(table.rows[i].cells[j], "")

    return True


def render_template_docx(
    case: Case,
    sources: SourceRegistry,
    draft: ReportDraft,
    spec: SpecBundle,
    template_path: str | Path,
    out_path: str | Path,
    asset_search_dirs: list[Path] | None = None,
    derived_dir: Path | None = None,
) -> None:
    template_path = Path(template_path)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    derived_dir = (derived_dir or out_path.parent).resolve()
    derived_dir.mkdir(parents=True, exist_ok=True)
    case_dir = guess_case_dir_from_derived_dir(derived_dir)

    doc = Document(str(template_path))

    sections_by_id = {s.id: s for s in spec.sections.sections}
    draft_by_id = {s.section_id: s for s in draft.sections}
    tables_by_id = {t.id: t for t in spec.tables.tables}
    figures_by_id = {f.id: f for f in spec.figures.figures}

    numbering = _Numbering(
        table_style=spec.sections.doc_profile.numbering.table_style,
        figure_style=spec.sections.doc_profile.numbering.figure_style,
    )

    anchor_entries: list[tuple[str, object]] = []
    for a in spec.template_map.anchors:
        anchor = str(getattr(a, "anchor", "") or "").strip()
        if not anchor:
            continue
        anchor_entries.append((anchor, a.insert))

    dup_in_spec: list[str] = []
    seen: set[str] = set()
    for anchor, _ in anchor_entries:
        if anchor in seen:
            dup_in_spec.append(anchor)
        else:
            seen.add(anchor)
    if dup_in_spec:
        d = sorted(set(dup_in_spec))
        raise ValueError(f"Spec template_map has duplicate anchors ({len(d)}): {d[:10]}")

    anchor_map = {anchor: insert for anchor, insert in anchor_entries}
    all_paras = list(_iter_all_paragraphs(doc))

    strict = bool(getattr(settings, "docx_strict_template", False))
    if strict:
        found = {p.text.strip() for p in all_paras}
        missing = [a for a in anchor_map.keys() if a.strip() and a.strip() not in found]
        if missing:
            raise ValueError(f"Template missing anchors ({len(missing)}): {missing[:10]}")

        dup_counts: dict[str, int] = defaultdict(int)
        for p in all_paras:
            t = p.text.strip()
            if t in anchor_map:
                dup_counts[t] += 1
        dup = sorted([a for a, c in dup_counts.items() if c > 1])
        if dup:
            examples = [f"{a}({dup_counts[a]})" for a in dup[:10]]
            raise ValueError(f"Template has duplicate anchors ({len(dup)}): {examples}")

    for p in all_paras:
        anchor_text = p.text.strip()
        if anchor_text not in anchor_map:
            continue

        insert = anchor_map[anchor_text]

        # conditional section insert
        if insert.conditional:
            sec_spec = sections_by_id.get(insert.id)
            if sec_spec and sec_spec.condition and not eval_condition(case, sec_spec.condition):
                _set_paragraph_text(p, "")
                continue

        if insert.type == "section":
            _insert_section_block(
                p,
                draft_by_id.get(insert.id),
                case=case,
                case_dir=case_dir,
                derived_dir=derived_dir,
                asset_search_dirs=asset_search_dirs,
                block_id=insert.id,
            )
            continue

        if insert.type == "table":
            t_spec = tables_by_id.get(insert.id)
            if not t_spec:
                if strict:
                    raise ValueError(f"Missing table spec: {insert.id}")
                _set_paragraph_text(p, f"【작성자 기입 필요】(표 스펙 없음: {insert.id})")
                continue
            td = build_table(case, sources, t_spec, spec.tables.defaults)
            # UX: avoid flooding the generated DOCX with "[작성자 기입 필요]" in table cells.
            # Missing inputs should be tracked via QA reports rather than cluttering the report body.
            empty_cell = str(getattr(spec.tables.defaults, "empty_cell", "") or "").strip()
            if empty_cell:
                cleaned_rows: list[list[str]] = []
                for row in td.rows:
                    cleaned_row: list[str] = []
                    for cell in row:
                        cell_text = "" if cell is None else str(cell)
                        if cell_text.strip() == empty_cell:
                            cell_text = ""
                        cleaned_row.append(cell_text)
                    cleaned_rows.append(cleaned_row)
                td = replace(td, rows=cleaned_rows)
            label = numbering.next_table(t_spec.chapter, fixed_label=getattr(t_spec, "label", None))
            caption = _docx_text(f"<{label}> {td.caption} {format_citations(td.source_ids)}")
            caption_style = "Caption" if (getattr(getattr(p, "style", None), "name", "") == "Normal") else None
            _set_paragraph_text(p, caption, style=caption_style)
            table = _find_table_after_paragraph(p)
            if table is None:
                if strict:
                    raise ValueError(f"Missing table under anchor: {anchor_text} (table_id={insert.id})")
                table = _insert_table_after(p, td.headers, td.rows)
            else:
                ok = _fill_table_in_place(table, td.headers, td.rows)
                if not ok:
                    if strict:
                        raise ValueError(
                            f"Table column mismatch under anchor: {anchor_text} (table_id={insert.id})"
                        )
                    table = _insert_table_after(p, td.headers, td.rows)
            if settings.docx_render_citations and td.source_ids:
                _insert_paragraph_after_element(
                    p, table._tbl, f"출처: {', '.join(td.source_ids)}", style=None
                )
            continue

        if insert.type == "figure":
            f_spec = figures_by_id.get(insert.id)
            if not f_spec:
                if strict:
                    raise ValueError(f"Missing figure spec: {insert.id}")
                _set_paragraph_text(p, f"【작성자 기입 필요】(그림 스펙 없음: {insert.id})")
                continue
            required_fig = is_required_figure(case, f_spec)
            fdata = resolve_figure(case, f_spec, asset_search_dirs=asset_search_dirs)
            p.clear()
            if fdata.file_path:
                try:
                    width = Mm(float(fdata.width_mm)) if (fdata.width_mm and fdata.width_mm > 0) else Inches(6.0)
                    width = _clamp_picture_width(doc, width)
                    width_mm_eff = _length_to_mm(width)
                    src_path = Path(fdata.file_path)
                    # Apply v2 figure controls (PDF rasterize / crop / resize) when requested.
                    if _needs_materialize(src_path, fdata):
                        out_dir = derived_dir / "figures" / "_materialized"
                        mat_res = materialize_figure_image_result(
                            src_path,
                            MaterializeOptions(
                                out_dir=out_dir,
                                fig_id=f_spec.id,
                                gen_method=fdata.gen_method,
                                crop=fdata.crop,
                                width_mm=width_mm_eff,
                                asset_type=f_spec.asset_type,
                            ),
                            include_meta=True,
                        )
                        mat = mat_res.path
                        note_obj = mat_res.meta if isinstance(getattr(mat_res, "meta", None), dict) else {"kind": "MATERIALIZE"}
                        if isinstance(note_obj, dict) and case_dir is not None:
                            note_obj = dict(note_obj)
                            try:
                                note_obj["src_rel_path"] = str(src_path.resolve().relative_to(case_dir.resolve())).replace("\\", "/")
                            except Exception:
                                note_obj["src_rel_path"] = src_path.name
                        note = json.dumps(note_obj, ensure_ascii=False, sort_keys=True)
                        evidence_type = "derived_png" if mat.suffix.lower() == ".png" else "derived_jpg"
                        record_derived_evidence(
                            case,
                            derived_path=mat,
                            related_fig_id=f_spec.id,
                            report_anchor=f_spec.id,
                            src_ids=fdata.source_ids,
                            evidence_type=evidence_type,
                            title=fdata.caption,
                            note=note,
                            pdf_page=mat_res.pdf_page,
                            pdf_page_source=mat_res.pdf_page_source,
                            used_in=f_spec.id,
                            case_dir=case_dir,
                        )
                        p.add_run().add_picture(str(mat), width=width)
                    else:
                        p.add_run().add_picture(str(src_path), width=width)
                except Exception as e:
                    if strict and required_fig:
                        raise ValueError(f"Failed to insert figure {f_spec.id}: {e}") from e
                    p.add_run("【첨부 오류】이미지 삽입 실패")
            else:
                if strict and required_fig:
                    raise ValueError(f"Missing figure asset for {f_spec.id}")
                p.add_run("【첨부 필요】")
            label = numbering.next_figure(f_spec.chapter, fixed_label=getattr(f_spec, "label", None))
            caption = _docx_text(f"<{label}> {fdata.caption} {format_citations(fdata.source_ids)}")
            _insert_paragraph_after(p, caption, style="Caption")
            continue

    doc.save(str(out_path))


def render_builtin_docx(
    case: Case,
    sources: SourceRegistry,
    draft: ReportDraft,
    spec: SpecBundle,
    out_path: str | Path,
    asset_search_dirs: list[Path] | None = None,
    derived_dir: Path | None = None,
) -> None:
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    derived_dir = (derived_dir or out_path.parent).resolve()
    derived_dir.mkdir(parents=True, exist_ok=True)
    case_dir = guess_case_dir_from_derived_dir(derived_dir)

    doc = Document()

    # Strict template mode is only meaningful when rendering into an existing DOCX template.
    strict = False

    sections_by_id = {s.id: s for s in spec.sections.sections}
    draft_by_id = {s.section_id: s for s in draft.sections}
    tables_by_id = {t.id: t for t in spec.tables.tables}
    figures_by_id = {f.id: f for f in spec.figures.figures}

    numbering = _Numbering(
        table_style=spec.sections.doc_profile.numbering.table_style,
        figure_style=spec.sections.doc_profile.numbering.figure_style,
    )

    for sec in spec.sections.sections:
        if sec.condition and not eval_condition(case, sec.condition):
            continue

        doc.add_heading(sec.heading, level=1)

        # section body
        sd = draft_by_id.get(sec.id)
        if sd and sd.paragraphs:
            for p in sd.paragraphs:
                doc.add_paragraph(p)
        else:
            doc.add_paragraph("【작성자 기입 필요】(섹션 초안 없음)")

        # tables
        for table_id in sec.outputs.tables:
            t_spec = tables_by_id.get(table_id)
            if not t_spec:
                doc.add_paragraph(f"【작성자 기입 필요】(표 스펙 없음: {table_id})")
                continue
            td = build_table(case, sources, t_spec, spec.tables.defaults)
            label = numbering.next_table(t_spec.chapter, fixed_label=getattr(t_spec, "label", None))
            doc.add_paragraph(f"<{label}> {td.caption} {format_citations(td.source_ids)}")
            table = doc.add_table(rows=1, cols=len(td.headers))
            table.style = "Table Grid"
            hdr = table.rows[0].cells
            for i, h in enumerate(td.headers):
                hdr[i].text = h
            for r in td.rows:
                cells = table.add_row().cells
                for i, v in enumerate(r):
                    cells[i].text = v
            if td.source_ids:
                doc.add_paragraph(f"출처: {', '.join(td.source_ids)}")

        # figures
        for fig_id in sec.outputs.figures:
            f_spec = figures_by_id.get(fig_id)
            if not f_spec:
                doc.add_paragraph(f"【작성자 기입 필요】(그림 스펙 없음: {fig_id})")
                continue
            required_fig = is_required_figure(case, f_spec)
            fdata = resolve_figure(case, f_spec, asset_search_dirs=asset_search_dirs)
            p = doc.add_paragraph()
            if fdata.file_path:
                try:
                    width = Mm(float(fdata.width_mm)) if (fdata.width_mm and fdata.width_mm > 0) else Inches(6.0)
                    width = _clamp_picture_width(doc, width)
                    width_mm_eff = _length_to_mm(width)
                    src_path = Path(fdata.file_path)
                    if _needs_materialize(src_path, fdata):
                        out_dir = derived_dir / "figures" / "_materialized"
                        mat_res = materialize_figure_image_result(
                            src_path,
                            MaterializeOptions(
                                out_dir=out_dir,
                                fig_id=f_spec.id,
                                gen_method=fdata.gen_method,
                                crop=fdata.crop,
                                width_mm=width_mm_eff,
                                asset_type=f_spec.asset_type,
                            ),
                            include_meta=True,
                        )
                        mat = mat_res.path
                        evidence_type = "derived_png" if mat.suffix.lower() == ".png" else "derived_jpg"
                        note_obj = mat_res.meta if isinstance(getattr(mat_res, "meta", None), dict) else {"kind": "MATERIALIZE"}
                        if isinstance(note_obj, dict) and case_dir is not None:
                            note_obj = dict(note_obj)
                            try:
                                note_obj["src_rel_path"] = str(src_path.resolve().relative_to(case_dir.resolve())).replace("\\", "/")
                            except Exception:
                                note_obj["src_rel_path"] = src_path.name
                        note = json.dumps(note_obj, ensure_ascii=False, sort_keys=True)
                        record_derived_evidence(
                            case,
                            derived_path=mat,
                            related_fig_id=f_spec.id,
                            report_anchor=f_spec.id,
                            src_ids=fdata.source_ids,
                            evidence_type=evidence_type,
                            title=fdata.caption,
                            note=note,
                            pdf_page=mat_res.pdf_page,
                            pdf_page_source=mat_res.pdf_page_source,
                            used_in=f_spec.id,
                            case_dir=case_dir,
                        )
                        p.add_run().add_picture(str(mat), width=width)
                    else:
                        p.add_run().add_picture(str(src_path), width=width)
                except Exception as e:
                    if strict and required_fig:
                        raise ValueError(f"Failed to insert figure {f_spec.id}: {e}") from e
                    p.add_run("【첨부 오류】이미지 삽입 실패")
            else:
                if strict and required_fig:
                    raise ValueError(f"Missing figure asset for {f_spec.id}")
                p.add_run("【첨부 필요】")
            label = numbering.next_figure(f_spec.chapter, fixed_label=getattr(f_spec, "label", None))
            doc.add_paragraph(f"<{label}> {fdata.caption} {format_citations(fdata.source_ids)}")

    doc.save(str(out_path))
