from pathlib import Path
from enum import Enum
import json
import re
import subprocess
import sys

import typer
from rich.console import Console

from eia_gen.config import settings
from eia_gen.services.loaders import load_case, load_sources

app = typer.Typer(add_completion=False)
console = Console()
_REPO_ROOT = Path(__file__).resolve().parents[2]
_BLOCK_ANCHOR_RE = re.compile(r"\[\[BLOCK:([^\]]+)\]\]")


def _resolve_repo_path(path: Path) -> Path:
    if path.is_absolute():
        return path
    # Users sometimes pass repo-prefixed relative paths like `eia-gen/output/...`
    # from the parent directory. Avoid creating `eia-gen/eia-gen/...` accidentally.
    parts = path.parts
    if parts and parts[0] == _REPO_ROOT.name:
        return (_REPO_ROOT.parent / path).resolve()
    return (_REPO_ROOT / path).resolve()


def _resolve_spec_dir(path: Path) -> Path:
    if path.is_absolute():
        return path
    cand = _REPO_ROOT / path
    return cand if cand.exists() else path


def _resolve_effective_template_path(
    *,
    template: Path | None,
    use_template_map: bool,
    spec_dir: Path | None,
    spec_bundle: object | None,
) -> Path | None:
    if template is not None and template.exists():
        return template
    if not use_template_map or spec_dir is None or spec_bundle is None:
        return None
    tf = str(getattr(getattr(spec_bundle, "template_map", None), "template_file", "") or "").strip()
    if not tf:
        return None
    cand = Path(tf).expanduser()
    if not cand.is_absolute():
        cand = (spec_dir.parent / cand).expanduser()
    return cand if cand.exists() else None


def _read_docx_all_text(docx_path: Path) -> str:
    import zipfile
    import xml.etree.ElementTree as ET

    texts: list[str] = []
    try:
        with zipfile.ZipFile(docx_path) as zf:
            for name in zf.namelist():
                if not (name.startswith("word/") and name.endswith(".xml")):
                    continue
                try:
                    data = zf.read(name)
                except Exception:
                    continue
                try:
                    root = ET.fromstring(data)
                except Exception:
                    continue
                for el in root.iter():
                    if el.tag.endswith("}t") and el.text:
                        texts.append(el.text)
    except Exception:
        return ""
    return "".join(texts)


def _compute_draft_section_id_allowlist(
    *,
    template_path: Path | None,
    spec_bundle: object | None,
    section_id_prefix: str | None = None,
) -> set[str] | None:
    if template_path is None or not template_path.exists():
        return None

    template_text = _read_docx_all_text(template_path)
    if not template_text.strip():
        return None

    allow: set[str] = set()

    # 1) Generic `[[BLOCK:<section_id>]]` placeholders (works without spec).
    for sid in _BLOCK_ANCHOR_RE.findall(template_text):
        sid = str(sid).strip()
        if sid:
            allow.add(sid)

    # 2) Spec anchors (supports non-placeholder anchor styles).
    sections = getattr(getattr(spec_bundle, "sections", None), "sections", None)
    if isinstance(sections, list):
        for sec in sections:
            sid = str(getattr(sec, "id", "") or "").strip()
            anchor = str(getattr(sec, "anchor", "") or "").strip()
            if sid and anchor and anchor in template_text:
                allow.add(sid)

    if not allow:
        return None
    if section_id_prefix:
        return {f"{section_id_prefix}:{sid}" for sid in allow}
    return allow


def _write_qa_next_actions_md(*, validation_report: Path, out_md: Path, kind: str) -> None:
    """Write QA next-actions markdown from a validation report (best-effort)."""
    script = (_REPO_ROOT / "scripts" / "qa_next_actions.py").resolve()
    if not script.exists():
        console.print(f"[yellow]WARN[/yellow] missing script: {script}")
        return

    out_md.parent.mkdir(parents=True, exist_ok=True)
    cmd = [
        sys.executable,
        str(script),
        "--validation-report",
        str(validation_report),
        "--kind",
        str(kind),
        "--out-md",
        str(out_md),
    ]
    try:
        proc = subprocess.run(cmd, check=False)
    except Exception as e:
        console.print(f"[yellow]WARN[/yellow] failed to run qa_next_actions: {e}")
        return
    if proc.returncode != 0:
        console.print(f"[yellow]WARN[/yellow] qa_next_actions failed (exit={proc.returncode})")


def _enrich_case_xlsx(*, xlsx: Path, overwrite_plan: bool) -> None:
    """Run DATA_REQUESTS planner+runner in-place and write _data_requests_run.json."""
    import json

    from eia_gen.services.data_requests.planner import plan_data_requests_for_workbook
    from eia_gen.services.data_requests.models import DataRequest
    from eia_gen.services.data_requests.runner import run_data_requests
    from eia_gen.services.data_requests.xlsx_io import (
        load_workbook,
        read_data_requests,
        save_workbook,
        write_data_requests,
    )

    wb = load_workbook(xlsx)
    existing_reqs = read_data_requests(wb)
    if overwrite_plan or not existing_reqs:
        plan = plan_data_requests_for_workbook(
            wb=wb,
            wms_layers_config=_REPO_ROOT / "config/wms_layers.yaml",
        )
        write_data_requests(wb, plan)
        console.print(f"[green]OK[/green] planned DATA_REQUESTS rows={len(plan)}")

    result = run_data_requests(
        wb=wb,
        case_dir=xlsx.parent.resolve(),
        wms_layers_config=_REPO_ROOT / "config/wms_layers.yaml",
        cache_config=_REPO_ROOT / "config/cache.yaml",
    )
    save_workbook(wb, xlsx)

    report_path = xlsx.parent / "_data_requests_run.json"
    report_path.write_text(
        json.dumps(
            {
                "executed": result.executed,
                "skipped": result.skipped,
                "warnings": result.warnings,
                "evidences": [e.__dict__ for e in result.evidences],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    console.print(
        f"[green]OK[/green] ran DATA_REQUESTS executed={result.executed} skipped={result.skipped} "
        f"(wrote {report_path})"
    )
    if result.warnings:
        console.print(f"[yellow]WARN[/yellow] DATA_REQUESTS had {len(result.warnings)} warnings (see report)")


@app.command()
def generate(
    case: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    sources: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    out: Path = typer.Option(Path("output/report.docx")),
    template: Path | None = typer.Option(None),
    spec_dir: Path = typer.Option(Path("spec")),
    use_template_map: bool = typer.Option(False),
    use_llm: bool = typer.Option(True),
) -> None:
    """Generate report.docx from case.yaml + sources.yaml."""
    case = case.resolve()
    sources = sources.resolve()
    out = _resolve_repo_path(out)
    spec_dir = _resolve_spec_dir(spec_dir)
    template = _resolve_repo_path(template) if template else None

    case_obj = load_case(case)
    sources_obj = load_sources(sources)

    llm = None
    if use_llm and settings.openai_api_key:
        from eia_gen.services.llm.openai_client import OpenAIChatClient

        llm = OpenAIChatClient(api_key=settings.openai_api_key, model=settings.openai_model)
    elif use_llm:
        console.print("[yellow]LLM disabled[/yellow] (missing OPENAI_API_KEY)")

    from eia_gen.services.writer import ReportWriter, SpecReportWriter, WriterOptions

    spec_bundle = None
    if spec_dir.exists():
        from eia_gen.spec.load import load_spec_bundle

        spec_bundle = load_spec_bundle(spec_dir)

    if spec_bundle is not None:
        writer = SpecReportWriter(
            spec=spec_bundle,
            sources=sources_obj,
            llm=llm,
            options=WriterOptions(use_llm=use_llm),
        )
    else:
        writer = ReportWriter(sources=sources_obj, llm=llm, options=WriterOptions(use_llm=use_llm))

    draft = writer.generate(case_obj)

    from eia_gen.services.docx.builder import build_docx

    build_docx(
        case=case_obj,
        sources=sources_obj,
        draft=draft,
        out_path=out,
        template_path=template,
        spec_dir=spec_dir,
        use_template_map=use_template_map,
        asset_base_dir=case.parent,
    )

    from eia_gen.services.qa.run import run_qa

    qa = run_qa(
        case_obj,
        sources_obj,
        draft,
        spec=spec_bundle,
        asset_search_dirs=[case.parent, out.parent, Path.cwd()],
        template_path=template,
    )
    qa_path = out.parent / "validation_report.json"
    qa_path.write_text(json.dumps(qa.model_dump(), indent=2, ensure_ascii=False), encoding="utf-8")

    from eia_gen.services.export.source_register_xlsx import build_source_register_xlsx_bytes

    xlsx_path = out.parent / "source_register.xlsx"
    report_tag = "DIA" if spec_dir.name == "spec_dia" else "EIA"
    effective_template = _resolve_effective_template_path(
        template=template, use_template_map=use_template_map, spec_dir=spec_dir, spec_bundle=spec_bundle
    )
    section_allowlist = _compute_draft_section_id_allowlist(template_path=effective_template, spec_bundle=spec_bundle)
    xlsx_path.write_bytes(
        build_source_register_xlsx_bytes(
            case_obj,
            sources_obj,
            draft,
            validation_reports=[(report_tag, qa)],
            report_tag=report_tag,
            draft_section_id_allowlist=section_allowlist,
        )
    )

    console.print(f"[green]OK[/green] wrote {out}")
    console.print(f"[green]OK[/green] wrote {qa_path}")
    console.print(f"[green]OK[/green] wrote {xlsx_path}")


@app.command()
def serve(
    host: str = typer.Option("127.0.0.1"),
    port: int = typer.Option(8000),
    reload: bool = typer.Option(False),
) -> None:
    """Run API server."""
    import uvicorn

    uvicorn.run("eia_gen.api:app", host=host, port=port, reload=reload)


@app.command()
def validate(
    case: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    sources: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
) -> None:
    """Validate case.yaml and sources.yaml against v1 schema."""
    _ = load_case(case)
    _ = load_sources(sources)
    console.print("[green]OK[/green] schema validated")


@app.command("xlsx-template")
def xlsx_template(
    out: Path = typer.Option(Path("templates/case_template.xlsx")),
) -> None:
    """Create a blank case.xlsx input template."""
    from eia_gen.services.xlsx.case_template import write_case_template_xlsx

    out_path = write_case_template_xlsx(out)
    console.print(f"[green]OK[/green] wrote {out_path}")


@app.command("xlsx-template-v2")
def xlsx_template_v2(
    out: Path = typer.Option(Path("templates/case_template.v2.xlsx")),
) -> None:
    """Create a blank v2(case.xlsx) template (snake_case + ENV/DRR sheets)."""
    from eia_gen.services.xlsx.case_template_v2 import write_case_template_v2_xlsx

    out_path = write_case_template_v2_xlsx(out)
    console.print(f"[green]OK[/green] wrote {out_path}")


@app.command("xlsx-upgrade-v2")
def xlsx_upgrade_v2(
    xlsx: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    out: Path | None = typer.Option(None, help="Write to a new file (default: in-place upgrade)."),
    backup: bool = typer.Option(True, help="When upgrading in-place, create a timestamped .bak copy first."),
    preserve_unknown_sheets: bool = typer.Option(True, help="Preserve non-template sheets (best-effort)."),
    preserve_unknown_columns: bool = typer.Option(True, help="Preserve non-template columns inside template sheets."),
) -> None:
    """Upgrade an existing case.xlsx(v2) to the latest template schema (adds missing sheets/columns without losing data)."""
    import shutil
    from datetime import datetime

    from eia_gen.services.xlsx.upgrade_v2 import upgrade_case_xlsx_v2

    xlsx = xlsx.resolve()
    out_path = _resolve_repo_path(out).resolve() if out else xlsx

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    if out_path == xlsx and backup:
        backup_path = xlsx.with_suffix(f".bak.{ts}.xlsx")
        shutil.copy2(xlsx, backup_path)
        console.print(f"[green]OK[/green] backup {backup_path}")

    tmp_out = out_path.with_suffix(f".tmp.{ts}.xlsx")
    rep = upgrade_case_xlsx_v2(
        xlsx_in=xlsx,
        xlsx_out=tmp_out,
        preserve_unknown_sheets=preserve_unknown_sheets,
        preserve_unknown_columns=preserve_unknown_columns,
    )
    tmp_out.replace(out_path)

    console.print(f"[green]OK[/green] upgraded {rep.input_path} -> {out_path}")
    if rep.added_sheets:
        console.print(f"- added sheets: {len(rep.added_sheets)} (e.g., {rep.added_sheets[:5]})")
    if rep.preserved_extra_sheets:
        console.print(f"- preserved extra sheets: {len(rep.preserved_extra_sheets)}")
    if rep.added_columns_by_sheet:
        console.print(f"- added columns (sheets): {len(rep.added_columns_by_sheet)}")


@app.command("xlsx-to-yaml")
def xlsx_to_yaml(
    xlsx: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    out: Path = typer.Option(Path("output/case.from_xlsx.yaml")),
) -> None:
    """Convert case.xlsx → case.yaml (for debugging / pipeline compatibility)."""
    from eia_gen.services.xlsx.case_reader import load_case_from_xlsx

    case_obj = load_case_from_xlsx(xlsx)
    out.parent.mkdir(parents=True, exist_ok=True)
    import yaml

    out.write_text(
        yaml.safe_dump(case_obj.model_dump(mode="python"), allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )
    console.print(f"[green]OK[/green] wrote {out}")


@app.command("xlsx-status")
def xlsx_status(
    xlsx: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    out: Path | None = typer.Option(None, help="Write JSON report to this path."),
    json_only: bool = typer.Option(False, help="Print JSON to stdout."),
) -> None:
    """Summarize how much of case.xlsx is filled (v1/v2)."""
    import json

    from rich.table import Table

    from eia_gen.services.xlsx.status import compute_xlsx_status

    st = compute_xlsx_status(xlsx)

    if out is not None:
        out_path = _resolve_repo_path(out)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(st.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
        console.print(f"[green]OK[/green] wrote {out_path}")

    if json_only:
        print(json.dumps(st.to_dict(), ensure_ascii=False, indent=2))
        return

    table = Table(title=f"case.xlsx status ({st.version})")
    table.add_column("sheet")
    table.add_column("rows", justify="right")
    table.add_column("data_fill", justify="right")
    table.add_column("src_empty", justify="right")
    table.add_column("src_tbd", justify="right")

    for s in st.sheet_stats:
        table.add_row(
            s.sheet,
            str(s.row_count),
            f"{s.data_fill_ratio:.2%}",
            str(s.src_id_empty_rows),
            str(s.src_id_tbd_rows),
        )

    console.print(table)
    console.print(f"Total data fill: {st.total_data_fill_ratio:.2%} (rows={st.total_rows})")
    if st.missing_files:
        console.print(f"[yellow]WARN[/yellow] missing referenced files: {len(st.missing_files)}")
        for m in st.missing_files[:10]:
            console.print(f"- {m}")


@app.command("generate-xlsx")
def generate_xlsx(
    xlsx: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    sources: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    out: Path = typer.Option(Path("output/report.docx")),
    template: Path | None = typer.Option(None),
    spec_dir: Path = typer.Option(Path("spec")),
    use_template_map: bool = typer.Option(False),
    use_llm: bool = typer.Option(True),
    debug: bool = typer.Option(False, help="Print debug info (e.g., figure auto-gen failures)."),
    enrich: bool = typer.Option(
        False,
        help="Run DATA_REQUESTS planner+runner before generating reports (updates the XLSX in-place).",
    ),
    enrich_overwrite_plan: bool = typer.Option(
        False,
        help="Overwrite existing DATA_REQUESTS rows before running (use with --enrich).",
    ),
    write_next_actions: bool = typer.Option(
        False,
        help="Write QA next-actions markdown next to validation_report.json (best-effort).",
    ),
) -> None:
    """Generate report.docx from case.xlsx + sources.yaml."""
    from eia_gen.services.xlsx.case_reader import load_case_from_xlsx

    xlsx = xlsx.resolve()
    sources = sources.resolve()
    out = _resolve_repo_path(out)
    spec_dir = _resolve_spec_dir(spec_dir)
    template = _resolve_repo_path(template) if template else None

    if enrich:
        _enrich_case_xlsx(xlsx=xlsx, overwrite_plan=enrich_overwrite_plan)

    case_obj = load_case_from_xlsx(xlsx)
    sources_obj = load_sources(sources)

    # DIA standard-form auto generation (best-effort): keep DIA checks from failing on empty forms.
    try:
        from eia_gen.services.dia.auto_generate import ensure_dia_standard_forms

        ensure_dia_standard_forms(case_obj)
    except Exception:
        # QA-only command should not crash on helper failures.
        pass

    llm = None
    if use_llm and settings.openai_api_key:
        from eia_gen.services.llm.openai_client import OpenAIChatClient

        llm = OpenAIChatClient(api_key=settings.openai_api_key, model=settings.openai_model)
    elif use_llm:
        console.print("[yellow]LLM disabled[/yellow] (missing OPENAI_API_KEY)")

    from eia_gen.services.writer import ReportWriter, SpecReportWriter, WriterOptions

    spec_bundle = None
    if spec_dir.exists():
        from eia_gen.spec.load import load_spec_bundle

        spec_bundle = load_spec_bundle(spec_dir)

    # DIA standard-form auto generation (best-effort): ensures at least one appendix-like form
    # (e.g., maintenance ledger) exists even when DRR_MAINTENANCE is empty.
    if spec_bundle is not None and spec_dir.name == "spec_dia":
        try:
            from eia_gen.services.dia.auto_generate import ensure_dia_standard_forms

            res_forms = ensure_dia_standard_forms(case_obj)
            if res_forms.get("errors"):
                console.print(
                    f"[yellow]WARN[/yellow] DIA standard-form auto-generation had errors: {len(res_forms['errors'])}"
                )
                if debug:
                    console.print(res_forms)
        except Exception as e:
            console.print(f"[yellow]WARN[/yellow] DIA standard-form auto-generation failed: {e}")
            if debug:
                import traceback

                console.print(traceback.format_exc())

    # Best-effort figure auto-generation. Runs before draft generation so
    # generated PNGs + source_ids are reflected in draft/source_register exports.
    if spec_bundle is not None:
        try:
            from eia_gen.services.figures.auto_generate import ensure_figures_from_geojson
            from eia_gen.services.figures.map_generate import ensure_figures_from_map_methods
            from eia_gen.services.figures.photo_sheet_generate import ensure_photo_sheets_from_attachments

            res_maps = ensure_figures_from_map_methods(
                case=case_obj,
                figure_specs=spec_bundle.figures.figures,
                case_xlsx=xlsx,
                out_dir=out.parent,
            )
            if res_maps.get("errors"):
                console.print(
                    f"[yellow]WARN[/yellow] map auto-generation had errors: {len(res_maps['errors'])} figures"
                )
                if debug:
                    console.print(res_maps)

            ensure_figures_from_geojson(
                case=case_obj,
                figure_specs=spec_bundle.figures.figures,
                case_xlsx=xlsx,
                out_dir=out.parent,
            )
            res = ensure_photo_sheets_from_attachments(
                case=case_obj,
                figure_specs=spec_bundle.figures.figures,
                case_xlsx=xlsx,
                style_path=_REPO_ROOT / "config/figure_style.yaml",
            )
            if res.get("errors"):
                console.print(
                    f"[yellow]WARN[/yellow] photo_sheet auto-generation had errors: {len(res['errors'])} figures"
                )
                if debug:
                    console.print(res)
        except Exception as e:
            console.print(f"[yellow]WARN[/yellow] figure auto-generation failed: {e}")
            if debug:
                import traceback

                console.print(traceback.format_exc())

    writer = (
        SpecReportWriter(spec=spec_bundle, sources=sources_obj, llm=llm, options=WriterOptions(use_llm=use_llm))
        if spec_bundle is not None
        else ReportWriter(sources=sources_obj, llm=llm, options=WriterOptions(use_llm=use_llm))
    )
    draft = writer.generate(case_obj)

    from eia_gen.services.docx.builder import build_docx

    build_docx(
        case=case_obj,
        sources=sources_obj,
        draft=draft,
        out_path=out,
        template_path=template,
        spec_dir=spec_dir,
        use_template_map=use_template_map,
        asset_base_dir=xlsx.parent,
    )

    from eia_gen.services.qa.run import run_qa

    qa = run_qa(
        case_obj,
        sources_obj,
        draft,
        spec=spec_bundle,
        asset_search_dirs=[xlsx.parent, out.parent, Path.cwd()],
        template_path=template,
        case_xlsx_path=xlsx,
    )
    qa_path = out.parent / "validation_report.json"
    qa_path.write_text(json.dumps(qa.model_dump(), indent=2, ensure_ascii=False), encoding="utf-8")

    from eia_gen.services.export.source_register_xlsx import build_source_register_xlsx_bytes

    xlsx_path = out.parent / "source_register.xlsx"
    report_tag = "DIA" if spec_dir.name == "spec_dia" else "EIA"
    effective_template = _resolve_effective_template_path(
        template=template, use_template_map=use_template_map, spec_dir=spec_dir, spec_bundle=spec_bundle
    )
    section_allowlist = _compute_draft_section_id_allowlist(template_path=effective_template, spec_bundle=spec_bundle)
    xlsx_path.write_bytes(
        build_source_register_xlsx_bytes(
            case_obj,
            sources_obj,
            draft,
            validation_reports=[(report_tag, qa)],
            report_tag=report_tag,
            draft_section_id_allowlist=section_allowlist,
        )
    )

    console.print(f"[green]OK[/green] wrote {out}")
    console.print(f"[green]OK[/green] wrote {qa_path}")
    console.print(f"[green]OK[/green] wrote {xlsx_path}")

    if settings.deliverables_dir:
        from eia_gen.services.export.deliverables import copy_docx_deliverable

        res = copy_docx_deliverable(
            report_path=out,
            kind=report_tag,
            case=case_obj,
            case_xlsx_path=xlsx,
            report_out_dir=out.parent,
            deliverables_dir=settings.deliverables_dir,
            deliverables_tag=settings.deliverables_tag,
        )
        for p in res.copied:
            console.print(f"[green]OK[/green] copied deliverable {p}")
        for reason in res.skipped:
            console.print(f"[yellow]WARN[/yellow] deliverables: {reason}")

    if write_next_actions:
        kind = "DIA" if spec_dir.name == "spec_dia" else "EIA"
        _write_qa_next_actions_md(validation_report=qa_path, out_md=(out.parent / "next_actions.md"), kind=kind)


@app.command("generate-xlsx-both")
def generate_xlsx_both(
    xlsx: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    sources: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    out_dir: Path = typer.Option(Path("output")),
    template_eia: Path | None = typer.Option(None),
    template_dia: Path | None = typer.Option(None),
    spec_dir_eia: Path = typer.Option(Path("spec")),
    spec_dir_dia: Path = typer.Option(Path("spec_dia")),
    use_template_map: bool = typer.Option(True),
    use_llm: bool = typer.Option(True),
    debug: bool = typer.Option(False, help="Print debug info (e.g., figure auto-gen failures)."),
    enrich: bool = typer.Option(
        False,
        help="Run DATA_REQUESTS planner+runner before generating reports (updates the XLSX in-place).",
    ),
    enrich_overwrite_plan: bool = typer.Option(
        False,
        help="Overwrite existing DATA_REQUESTS rows before running (use with --enrich).",
    ),
    submission: bool = typer.Option(
        False,
        help="Use strict 'submission mode' rules (treat missing core sheets/rows as ERROR).",
    ),
    write_next_actions: bool = typer.Option(
        False,
        help="Write QA next-actions markdown next to validation_report_*.json (best-effort).",
    ),
) -> None:
    """Generate BOTH EIA + DIA reports from one case.xlsx.

    Outputs (default):
    - output/report_eia.docx
    - output/report_dia.docx
    - output/validation_report_eia.json
    - output/validation_report_dia.json
    - output/source_register.xlsx (merged usage)
    """

    from eia_gen.services.xlsx.case_reader import load_case_from_xlsx

    xlsx = xlsx.resolve()
    sources = sources.resolve()
    out_dir = _resolve_repo_path(out_dir)
    spec_dir_eia = _resolve_spec_dir(spec_dir_eia)
    spec_dir_dia = _resolve_spec_dir(spec_dir_dia)

    template_eia = _resolve_repo_path(template_eia) if template_eia else None
    template_dia = _resolve_repo_path(template_dia) if template_dia else None

    if enrich:
        _enrich_case_xlsx(xlsx=xlsx, overwrite_plan=enrich_overwrite_plan)

    case_obj = load_case_from_xlsx(xlsx)
    sources_obj = load_sources(sources)

    llm = None
    if use_llm and settings.openai_api_key:
        from eia_gen.services.llm.openai_client import OpenAIChatClient

        llm = OpenAIChatClient(api_key=settings.openai_api_key, model=settings.openai_model)
    elif use_llm:
        console.print("[yellow]LLM disabled[/yellow] (missing OPENAI_API_KEY)")

    from eia_gen.services.writer import SpecReportWriter, WriterOptions
    from eia_gen.spec.load import load_spec_bundle

    if not spec_dir_eia.exists():
        raise typer.BadParameter(f"spec_dir_eia not found: {spec_dir_eia}")
    if not spec_dir_dia.exists():
        raise typer.BadParameter(f"spec_dir_dia not found: {spec_dir_dia}")

    spec_eia = load_spec_bundle(spec_dir_eia)
    spec_dia = load_spec_bundle(spec_dir_dia)

    writer_eia = SpecReportWriter(spec=spec_eia, sources=sources_obj, llm=llm, options=WriterOptions(use_llm=use_llm))
    writer_dia = SpecReportWriter(spec=spec_dia, sources=sources_obj, llm=llm, options=WriterOptions(use_llm=use_llm))

    # Default templates (if not provided)
    if template_eia is None:
        # Prefer SSOT sample-reuse template when available (recommended workflow),
        # otherwise fall back to the normal scaffolded template.
        for cand in [
            _REPO_ROOT / "templates/report_template.sample_changwon_2025.ssot_full.scaffolded.docx",
            _REPO_ROOT / "templates/report_template.sample_changwon_2025.scaffolded.docx",
            _REPO_ROOT / "templates/report_template.sample_changwon_2025.docx",
            _REPO_ROOT / "templates/report_template.docx",
        ]:
            if cand.exists():
                template_eia = cand
                break
    if template_dia is None:
        for cand in [
            _REPO_ROOT / "templates/dia_template.scaffolded.docx",
            _REPO_ROOT / "templates/dia_template.docx",
        ]:
            if cand.exists():
                template_dia = cand
                break

    out_dir.mkdir(parents=True, exist_ok=True)

    # Best-effort figure auto-generation shared across EIA/DIA.
    try:
        from eia_gen.services.figures.auto_generate import ensure_figures_from_geojson
        from eia_gen.services.figures.map_generate import ensure_figures_from_map_methods
        from eia_gen.services.figures.photo_sheet_generate import ensure_photo_sheets_from_attachments

        res_maps = ensure_figures_from_map_methods(
            case=case_obj,
            figure_specs=[*spec_eia.figures.figures, *spec_dia.figures.figures],
            case_xlsx=xlsx,
            out_dir=out_dir,
        )
        if res_maps.get("errors"):
            console.print(f"[yellow]WARN[/yellow] map auto-generation had errors: {len(res_maps['errors'])} figures")
            if debug:
                console.print(res_maps)

        ensure_figures_from_geojson(
            case=case_obj,
            figure_specs=[*spec_eia.figures.figures, *spec_dia.figures.figures],
            case_xlsx=xlsx,
            out_dir=out_dir,
        )
        res = ensure_photo_sheets_from_attachments(
            case=case_obj,
            figure_specs=[*spec_eia.figures.figures, *spec_dia.figures.figures],
            case_xlsx=xlsx,
            style_path=_REPO_ROOT / "config/figure_style.yaml",
        )
        if res.get("errors"):
            console.print(f"[yellow]WARN[/yellow] photo_sheet auto-generation had errors: {len(res['errors'])} figures")
            if debug:
                console.print(res)
    except Exception as e:
        console.print(f"[yellow]WARN[/yellow] figure auto-generation failed: {e}")
        if debug:
            import traceback

            console.print(traceback.format_exc())

    # DIA standard-form auto generation (best-effort): keep DIA tables from failing on empty forms.
    try:
        from eia_gen.services.dia.auto_generate import ensure_dia_standard_forms

        res_forms = ensure_dia_standard_forms(case_obj)
        if res_forms.get("errors"):
            console.print(f"[yellow]WARN[/yellow] DIA standard-form auto-generation had errors: {len(res_forms['errors'])}")
            if debug:
                console.print(res_forms)
    except Exception as e:
        console.print(f"[yellow]WARN[/yellow] DIA standard-form auto-generation failed: {e}")
        if debug:
            import traceback

            console.print(traceback.format_exc())

    draft_eia = writer_eia.generate(case_obj)
    draft_dia = writer_dia.generate(case_obj)

    from eia_gen.services.docx.builder import build_docx

    out_eia = out_dir / "report_eia.docx"
    out_dia = out_dir / "report_dia.docx"

    build_docx(
        case=case_obj,
        sources=sources_obj,
        draft=draft_eia,
        out_path=out_eia,
        template_path=template_eia,
        spec_dir=spec_dir_eia,
        use_template_map=use_template_map,
        asset_base_dir=xlsx.parent,
    )
    build_docx(
        case=case_obj,
        sources=sources_obj,
        draft=draft_dia,
        out_path=out_dia,
        template_path=template_dia,
        spec_dir=spec_dir_dia,
        use_template_map=use_template_map,
        asset_base_dir=xlsx.parent,
    )

    from eia_gen.services.qa.run import run_qa

    asset_search_dirs = [xlsx.parent, out_dir, Path.cwd()]
    data_rules_eia = (_REPO_ROOT / "config/data_acquisition_rules_submission.yaml") if submission else None
    data_rules_dia = (_REPO_ROOT / "config/data_acquisition_rules_submission_dia.yaml") if submission else None
    qa_eia = run_qa(
        case_obj,
        sources_obj,
        draft_eia,
        spec=spec_eia,
        asset_search_dirs=asset_search_dirs,
        template_path=template_eia,
        case_xlsx_path=xlsx,
        data_acquisition_rules_path=data_rules_eia,
        submission_mode=submission,
    )
    qa_dia = run_qa(
        case_obj,
        sources_obj,
        draft_dia,
        spec=spec_dia,
        asset_search_dirs=asset_search_dirs,
        template_path=template_dia,
        case_xlsx_path=xlsx,
        data_acquisition_rules_path=data_rules_dia,
        submission_mode=submission,
    )
    qa_eia_path = out_dir / "validation_report_eia.json"
    qa_dia_path = out_dir / "validation_report_dia.json"
    qa_eia_path.write_text(json.dumps(qa_eia.model_dump(), indent=2, ensure_ascii=False), encoding="utf-8")
    qa_dia_path.write_text(json.dumps(qa_dia.model_dump(), indent=2, ensure_ascii=False), encoding="utf-8")

    # Merge drafts for a single source_register.xlsx (EIA/DIA usage mixed with prefixes)
    from eia_gen.services.draft import ReportDraft, SectionDraft
    from eia_gen.services.export.source_register_xlsx import build_source_register_xlsx_bytes

    merged_sections: list[SectionDraft] = []
    for tag, d in [("EIA", draft_eia), ("DIA", draft_dia)]:
        for s in d.sections:
            merged_sections.append(
                SectionDraft(
                    section_id=f"{tag}:{s.section_id}",
                    title=s.title,
                    paragraphs=s.paragraphs,
                    tables=s.tables,
                    figures=s.figures,
                    todos=s.todos,
                    meta=s.meta,
                )
            )
    merged_draft = ReportDraft(sections=merged_sections)

    xlsx_path = out_dir / "source_register.xlsx"
    effective_template_eia = _resolve_effective_template_path(
        template=template_eia, use_template_map=use_template_map, spec_dir=spec_dir_eia, spec_bundle=spec_eia
    )
    effective_template_dia = _resolve_effective_template_path(
        template=template_dia, use_template_map=use_template_map, spec_dir=spec_dir_dia, spec_bundle=spec_dia
    )
    allow_eia = _compute_draft_section_id_allowlist(
        template_path=effective_template_eia, spec_bundle=spec_eia, section_id_prefix="EIA"
    )
    allow_dia = _compute_draft_section_id_allowlist(
        template_path=effective_template_dia, spec_bundle=spec_dia, section_id_prefix="DIA"
    )
    merged_allowlist: set[str] | None = None
    if allow_eia or allow_dia:
        merged_allowlist = set()
        if allow_eia:
            merged_allowlist |= allow_eia
        if allow_dia:
            merged_allowlist |= allow_dia
    xlsx_path.write_bytes(
        build_source_register_xlsx_bytes(
            case_obj,
            sources_obj,
            merged_draft,
            validation_reports=[("EIA", qa_eia), ("DIA", qa_dia)],
            draft_section_id_allowlist=merged_allowlist,
        )
    )

    console.print(f"[green]OK[/green] wrote {out_eia}")
    console.print(f"[green]OK[/green] wrote {out_dia}")
    console.print(f"[green]OK[/green] wrote {qa_eia_path}")
    console.print(f"[green]OK[/green] wrote {qa_dia_path}")
    console.print(f"[green]OK[/green] wrote {xlsx_path}")

    if settings.deliverables_dir:
        from eia_gen.services.export.deliverables import copy_docx_deliverable

        res = copy_docx_deliverable(
            report_path=out_eia,
            kind="EIA",
            case=case_obj,
            case_xlsx_path=xlsx,
            report_out_dir=out_dir,
            deliverables_dir=settings.deliverables_dir,
            deliverables_tag=settings.deliverables_tag,
        )
        for p in res.copied:
            console.print(f"[green]OK[/green] copied deliverable {p}")
        for reason in res.skipped:
            console.print(f"[yellow]WARN[/yellow] deliverables: {reason}")

    if write_next_actions:
        _write_qa_next_actions_md(
            validation_report=qa_eia_path,
            out_md=(out_dir / "next_actions_eia.md"),
            kind="EIA",
        )
        _write_qa_next_actions_md(
            validation_report=qa_dia_path,
            out_md=(out_dir / "next_actions_dia.md"),
            kind="DIA",
        )


@app.command("pdf-image-docx")
def pdf_image_docx(
    pdf: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    out: Path = typer.Option(...),
    dpi: int = typer.Option(200, help="Rasterization DPI (higher = sharper, bigger DOCX)."),
) -> None:
    """Make a DOCX that is visually identical to a PDF (each page as an image).

    This is a strict "layout identical" option. The resulting DOCX is not meaningfully editable.
    """
    from eia_gen.services.pdf.pdf_image_docx import pdf_to_image_docx

    pdf_to_image_docx(pdf_path=pdf, out_docx_path=out, dpi=dpi)
    console.print(f"[green]OK[/green] wrote {out}")


@app.command("template-check")
def template_check(
    template: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    spec_dir: Path = typer.Option(Path("spec"), help="Spec directory (spec or spec_dia)."),
    out: Path | None = typer.Option(None, help="Write JSON report to a file (optional)."),
    allow_missing_anchors: bool = typer.Option(
        False,
        help="Allow missing anchors (do not fail); still reports missing anchors as WARN.",
    ),
    ignore_missing_anchor_prefix: list[str] = typer.Option(
        [],
        help="Ignore missing anchors that start with this prefix (repeatable).",
    ),
) -> None:
    """Check a DOCX template readiness for 'template mode' rendering.

    - By default, verifies all anchors exist in the template (fail-fast).
      Use `--allow-missing-anchors` or `--ignore-missing-anchor-prefix` for
      workflows where missing anchors are intentional (e.g., SSOT templates).
    - Verifies each [[TABLE:...]] anchor has a table 바로 아래에 존재(권장: in-place fill).
    """
    import json

    from eia_gen.services.docx.template_tools import check_template_docx
    from eia_gen.spec.load import load_spec_bundle

    template = template.resolve()
    spec_dir = _resolve_spec_dir(spec_dir)
    spec = load_spec_bundle(spec_dir)

    rep = check_template_docx(
        template_path=template,
        spec=spec,
        spec_dir=spec_dir,
        ignore_missing_anchor_prefixes=ignore_missing_anchor_prefix,
    )
    if out:
        out = _resolve_repo_path(out)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(rep.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
        console.print(f"[green]OK[/green] wrote {out}")

    missing_anchor_errors = 0 if allow_missing_anchors else len(rep.missing_anchors)
    has_errors = bool(
        rep.duplicate_anchors_in_spec
        or rep.duplicate_anchors_in_template
        or missing_anchor_errors
        or rep.table_anchors_missing_table
        or rep.table_anchors_column_mismatch
    )

    if (
        rep.duplicate_anchors_in_spec
        or rep.duplicate_anchors_in_template
        or rep.missing_anchors
        or rep.ignored_missing_anchors
        or rep.table_anchors_missing_table
        or rep.table_anchors_column_mismatch
    ):
        console.print(
            f"[yellow]WARN[/yellow] dup_spec={len(rep.duplicate_anchors_in_spec)} "
            f"dup_template={len(rep.duplicate_anchors_in_template)} "
            f"missing_anchors={len(rep.missing_anchors)} "
            f"ignored_missing_anchors={len(rep.ignored_missing_anchors)} "
            f"missing_tables={len(rep.table_anchors_missing_table)} "
            f"col_mismatch={len(rep.table_anchors_column_mismatch)}"
        )

    if has_errors:
        raise typer.Exit(code=2)

    console.print("[green]OK[/green] template looks ready")


@app.command("template-scaffold")
def template_scaffold(
    template: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    out: Path = typer.Option(..., help="Output scaffolded DOCX path."),
    spec_dir: Path = typer.Option(Path("spec"), help="Spec directory (spec or spec_dia)."),
    min_data_rows: int = typer.Option(1, help="Placeholder data rows per table (>=1)."),
    table_style: str = typer.Option("Table Grid", help="Table style name to apply when creating placeholders."),
) -> None:
    """Create a scaffolded template with placeholder tables under every [[TABLE:...]] anchor.

    This improves layout fidelity by enabling 'in-place fill' instead of creating new tables at render time.
    """
    import json

    from eia_gen.services.docx.template_tools import scaffold_template_docx
    from eia_gen.spec.load import load_spec_bundle

    template = template.resolve()
    out = _resolve_repo_path(out)
    spec_dir = _resolve_spec_dir(spec_dir)
    spec = load_spec_bundle(spec_dir)

    rep = scaffold_template_docx(
        template_path=template,
        out_path=out,
        spec=spec,
        spec_dir=spec_dir,
        min_data_rows=max(1, int(min_data_rows)),
        table_style=table_style,
    )

    console.print(f"[green]OK[/green] wrote {out}")
    # Also dump a small report next to the output for auditability.
    report_path = out.with_suffix(".template_check.json")
    report_path.write_text(json.dumps(rep.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
    console.print(f"[green]OK[/green] wrote {report_path}")


@app.command("init-case")
def init_case(
    out_dir: Path = typer.Option(..., help="Output case folder (will be created)."),
    project_id: str = typer.Option("PRJ-YYYY-0001"),
    project_name: str = typer.Option("OO 관광농원 조성사업"),
    copy_templates: bool = typer.Option(True, help="Copy editable DOCX templates into the case folder."),
    plan_data_requests: bool = typer.Option(True, help="Pre-populate DATA_REQUESTS with a default plan."),
    reference_pack_id: str = typer.Option("", help="Reference pack id (folder name) to apply during init (optional)."),
    reference_packs_dir: Path = typer.Option(Path("reference_packs"), help="Parent dir containing reference packs (repo-relative by default)."),
) -> None:
    """Create a new case folder starter kit (v2 case.xlsx + sources.yaml + attachments/)."""
    import shutil
    from datetime import datetime

    import yaml

    from eia_gen.services.xlsx.case_template_v2 import write_case_template_v2_xlsx

    out_dir = _resolve_repo_path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # NOTE: These CLI functions are also called directly in unit tests.
    # Typer uses OptionInfo objects as default values; unwrap them here.
    ref_pack_id = str(getattr(reference_pack_id, "default", reference_pack_id) or "").strip()
    ref_packs_dir = getattr(reference_packs_dir, "default", reference_packs_dir)
    if not isinstance(ref_packs_dir, Path):
        ref_packs_dir = Path(str(ref_packs_dir))

    # attachments skeleton
    attachments = out_dir / "attachments"
    for sub in ["figures", "photos", "gis", "evidence", "inbox", "normalized", "derived"]:
        (attachments / sub).mkdir(parents=True, exist_ok=True)

    # reports (ingest/extract logs)
    (out_dir / "reports").mkdir(parents=True, exist_ok=True)

    # case.xlsx (v2)
    case_xlsx = out_dir / "case.xlsx"
    write_case_template_v2_xlsx(case_xlsx)

    # Seed a few guide rows so QA tables don't start empty (users can edit/remove freely).
    try:
        from eia_gen.services.data_requests.xlsx_io import load_workbook, save_workbook

        wb = load_workbook(case_xlsx)

        def _sheet_has_data(sheet_name: str) -> bool:
            if sheet_name not in wb.sheetnames:
                return False
            ws = wb[sheet_name]
            for r in ws.iter_rows(min_row=2, values_only=True):
                if any(v is not None and (not isinstance(v, str) or v.strip()) for v in r):
                    return True
            return False

        if not _sheet_has_data("ZONING_OVERLAY"):
            ws = wb["ZONING_OVERLAY"]
            ws.append(["ZO-001", "NATURE", "자연공원", "UNKNOWN", "", "", "【자료 확인 필요】(발급본/WMS)", "CLIENT_PROVIDED", "S-CLIENT-001"])
            ws.append([
                "ZO-002",
                "NATURE",
                "생태자연도(환경공간정보서비스)",
                "UNKNOWN",
                "",
                "",
                "【자료 확인 필요】(WMS: ECO_NATURE_MCEE_2015)",
                "OFFICIAL_DB",
                "SRC_MCEE_EZMAP_ECO_2015",
            ])
            ws.append([
                "ZO-003",
                "DISASTER",
                "침수흔적도(생활안전지도)",
                "UNKNOWN",
                "",
                "",
                "【자료 확인 필요】(WMS: FLOOD_TRACE)",
                "OFFICIAL_DB",
                "SRC_SAFEMAP_FLOOD_TRACE",
            ])
            ws.append([
                "ZO-004",
                "DISASTER",
                "산사태위험지도(생활안전지도)",
                "UNKNOWN",
                "",
                "",
                "【자료 확인 필요】(WMS: LANDSLIDE_RISK)",
                "OFFICIAL_DB",
                "SRC_SAFEMAP_LANDSLIDE",
            ])

        if not _sheet_has_data("DRR_BASE_HAZARD"):
            ws = wb["DRR_BASE_HAZARD"]
            ws.append(["HZ-01", "FLOOD", "UNKNOWN", "N", "", "", "", "OFFICIAL_DB", "S-03"])
            ws.append(["HZ-02", "INLAND", "UNKNOWN", "N", "", "", "", "OFFICIAL_DB", "S-03"])
            ws.append(["HZ-03", "SEDIMENT", "UNKNOWN", "N", "", "", "", "OFFICIAL_DB", "S-03"])
            ws.append(["HZ-04", "SLOPE", "UNKNOWN", "N", "", "", "", "OFFICIAL_DB", "S-03"])

        save_workbook(wb, case_xlsx)

    except Exception as e:
        console.print(f"[yellow]WARN[/yellow] failed to seed/init case.xlsx extras: {e}")

    # sources.yaml (v2-style wrapper; SourceRegistry ignores extra keys)
    sources_yaml = out_dir / "sources.yaml"
    src_doc = {
        "version": "2.0",
        "project": {"id": project_id, "name": project_name},
        "sources": [
            {
                "source_id": "S-CLIENT-001",
                "kind": "client_provided",
                "title": "의뢰인 제공 기본자료",
                "publisher": "의뢰인",
                "issued_date": datetime.now().date().isoformat(),
                "local_file": "attachments/",
                "note": "토지이용계획확인서/설계도면/사진/발급본 등",
            }
            ,
            {
                "source_id": "S-01",
                "kind": "law",
                "title": "관련 법령/지침(대상판정/작성기준)",
                "publisher": "국가법령정보센터/관계기관",
                "issued_date": "최신본",
                "note": "대상판정 근거 및 작성기준 확인",
            },
            {
                "source_id": "S-03",
                "kind": "api",
                "title": "공공 DB/OpenAPI/WMS 기반 기초현황 자료",
                "publisher": "공공데이터포털/관계기관",
                "issued_date": "",
                "accessed_date": datetime.now().date().isoformat(),
                "note": "AIRKOREA/KMA_ASOS/WMS 등 DATA_REQUESTS 수집 근거",
            },
            {
                "source_id": "SRC_VWORLD_WMTS",
                "kind": "wmts",
                "title": "VWorld 배경지도(WMTS)",
                "publisher": "VWorld",
                "issued_date": "",
                "accessed_date": datetime.now().date().isoformat(),
                "url": "https://api.vworld.kr/",
                "note": "WMTS 타일 기반 배경지도(환경설정: config/basemap.yaml, env: VWORLD_API_KEY 필요)",
            },
            {
                "source_id": "SRC_OSM_TILE",
                "kind": "tile",
                "title": "OpenStreetMap 기본 타일(배경지도)",
                "publisher": "OpenStreetMap contributors",
                "issued_date": "",
                "accessed_date": datetime.now().date().isoformat(),
                "url": "https://tile.openstreetmap.org/",
                "note": "XYZ 타일 기반 배경지도(키 불필요; VWorld 키 미설정 시 best-effort 폴백)",
            },
            {
                "source_id": "SRC_MCEE_EZMAP_ECO_2015",
                "kind": "wms",
                "title": "환경공간정보서비스 생태자연도(2015)",
                "publisher": "환경공간정보서비스",
                "accessed_date": datetime.now().date().isoformat(),
                "url": "https://api.mcee.go.kr/geoserver/wms",
            },
            {
                "source_id": "SRC_SAFEMAP_FLOOD_TRACE",
                "kind": "wms",
                "title": "생활안전지도 침수흔적도",
                "publisher": "생활안전지도",
                "accessed_date": datetime.now().date().isoformat(),
                "url": "https://www.safemap.go.kr/",
            },
            {
                "source_id": "SRC_SAFEMAP_LANDSLIDE",
                "kind": "wms",
                "title": "생활안전지도 산사태위험지도",
                "publisher": "생활안전지도",
                "accessed_date": datetime.now().date().isoformat(),
                "url": "https://www.safemap.go.kr/",
            },
            {
                "source_id": "SRC_DATAGO_ECOLOGYZMP",
                "kind": "api",
                "title": "국립생태원 생태자연도 서비스(OpenAPI)",
                "publisher": "공공데이터포털(국립생태원)",
                "accessed_date": datetime.now().date().isoformat(),
                "url": "http://apis.data.go.kr/B553084/ecoapi/EcologyzmpService/wms/getEcologyzmpWMS",
            },
        ],
    }
    sources_yaml.write_text(
        yaml.safe_dump(src_doc, allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )

    # Optionally apply a reference pack (baseline reuse for same region).
    if ref_pack_id:
        try:
            from eia_gen.services.reference_packs import apply_reference_pack

            pack_base = _resolve_repo_path(ref_packs_dir)
            pack_dir = (pack_base / ref_pack_id).resolve()
            if not pack_dir.exists():
                raise FileNotFoundError(f"reference pack not found: {pack_dir}")
            res = apply_reference_pack(xlsx=case_xlsx, sources=sources_yaml, pack_dir=pack_dir)
            console.print(f"[green]OK[/green] applied reference pack={reference_pack_id} (sheets_applied={len(res.get("applied_sheets", []))})")
        except Exception as e:
            console.print(f"[yellow]WARN[/yellow] failed to apply reference pack: {e}")

    # DATA_REQUESTS plan (after applying reference pack, so the plan can be minimal).
    if plan_data_requests:
        try:
            from eia_gen.services.data_requests.planner import plan_data_requests_for_workbook
            from eia_gen.services.data_requests.xlsx_io import load_workbook, save_workbook, write_data_requests

            wb = load_workbook(case_xlsx)
            plan = plan_data_requests_for_workbook(wb=wb, wms_layers_config=_REPO_ROOT / "config/wms_layers.yaml")
            write_data_requests(wb, plan)
            save_workbook(wb, case_xlsx)
            console.print(f"[green]OK[/green] populated DATA_REQUESTS rows={len(plan)}")
        except Exception as e:
            console.print(f"[yellow]WARN[/yellow] failed to plan DATA_REQUESTS: {e}")

    # templates (optional copy; user can edit per project)
    if copy_templates:
        tpl_out = out_dir / "templates"
        tpl_out.mkdir(parents=True, exist_ok=True)
        eia_src = None
        for cand in [
            _REPO_ROOT / "templates/report_template.sample_changwon_2025.scaffolded.docx",
            _REPO_ROOT / "templates/report_template.sample_changwon_2025.docx",
            _REPO_ROOT / "templates/report_template.docx",
        ]:
            if cand.exists():
                eia_src = cand
                break
        if eia_src is not None:
            shutil.copyfile(eia_src, tpl_out / "report_template_eia.docx")
        dia_src = None
        for cand in [
            _REPO_ROOT / "templates/dia_template.scaffolded.docx",
            _REPO_ROOT / "templates/dia_template.docx",
        ]:
            if cand.exists():
                dia_src = cand
                break
        if dia_src is not None:
            shutil.copyfile(dia_src, tpl_out / "report_template_dia.docx")

    console.print(f"[green]OK[/green] wrote {case_xlsx}")
    console.print(f"[green]OK[/green] wrote {sources_yaml}")
    console.print(f"[green]OK[/green] created {attachments}/")
    if copy_templates:
        console.print(f"[green]OK[/green] created {out_dir/'templates'}/")


@app.command("purge-derived")
def purge_derived_cmd(
    case_dir: Path = typer.Option(..., exists=True, file_okay=False, dir_okay=True, help="Case folder."),
    days: int = typer.Option(30, help="Delete derived files older than N days (0 = delete all)."),
    apply: bool = typer.Option(False, help="Actually delete files (default: dry-run)."),
) -> None:
    """Purge case-local derived artifacts under attachments/derived (best-effort).

    This helps keep large `_materialized` outputs from growing without bound.
    """
    import os
    import time

    case_dir = case_dir.expanduser().resolve()
    target = (case_dir / "attachments" / "derived").resolve()
    if not target.exists():
        console.print(f"[yellow]SKIP[/yellow] not found: {target}")
        return

    days_i = int(days)
    cutoff = time.time() - max(0, days_i) * 86400

    candidates: list[Path] = []
    total_bytes = 0

    for root, dirs, files in os.walk(target, topdown=True, followlinks=False):
        root_path = Path(root)
        dirs[:] = [d for d in dirs if not (root_path / d).is_symlink()]
        for fn in files:
            p = root_path / fn
            try:
                st = p.lstat()
            except FileNotFoundError:
                continue
            if days_i <= 0 or float(st.st_mtime) < cutoff:
                candidates.append(p)
                total_bytes += int(st.st_size)

    def _fmt_bytes(n: int) -> str:
        units = ["B", "KB", "MB", "GB", "TB"]
        v = float(n)
        for u in units:
            if v < 1024.0 or u == units[-1]:
                return f"{v:.1f}{u}" if u != "B" else f"{int(v)}B"
            v /= 1024.0
        return f"{v:.1f}TB"

    console.print(f"target: {target}")
    console.print(f"candidates: {len(candidates)} files ({_fmt_bytes(total_bytes)})")
    for p in candidates[:20]:
        try:
            rel = p.relative_to(case_dir)
        except Exception:
            rel = p
        console.print(f"- {rel}")
    if len(candidates) > 20:
        console.print(f"... (+{len(candidates) - 20} more)")

    if not apply:
        console.print("[yellow]DRY-RUN[/yellow] no files were deleted (use --apply to delete)")
        return

    deleted = 0
    for p in candidates:
        try:
            p.unlink()
            deleted += 1
        except Exception:
            continue

    removed_dirs = 0
    for root, dirs, files in os.walk(target, topdown=False, followlinks=False):
        rp = Path(root)
        if rp == target:
            continue
        try:
            if not any(rp.iterdir()):
                rp.rmdir()
                removed_dirs += 1
        except Exception:
            continue

    console.print(f"[green]OK[/green] deleted files={deleted} removed_dirs={removed_dirs}")


@app.command("ingest-attachments")
def ingest_attachments_cmd(
    xlsx: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    inbox: Path = typer.Option(Path("attachments/inbox"), help="Inbox folder (relative to case dir by default)."),
    normalized: Path = typer.Option(
        Path("attachments/normalized"), help="Normalized folder (relative to case dir by default)."
    ),
    src_id: str = typer.Option("S-CLIENT-001", help="Source id for ingested files."),
    data_origin: str = typer.Option("CLIENT_PROVIDED", help="DATA_ORIGIN for ingested evidence."),
    evidence_type: str = typer.Option("", help="Override EVIDENCE_TYPE (default: infer by extension)."),
    copy_only: bool = typer.Option(False, help="Copy files instead of moving them."),
) -> None:
    """Ingest files from `attachments/inbox/` into `attachments/normalized/` and register in ATTACHMENTS sheet."""
    from eia_gen.services.ingest_attachments import ingest_inbox

    xlsx = xlsx.resolve()
    case_dir = xlsx.parent.resolve()

    inbox_dir = inbox
    if not inbox_dir.is_absolute():
        inbox_dir = (case_dir / inbox_dir).resolve()

    normalized_dir = normalized
    if not normalized_dir.is_absolute():
        normalized_dir = (case_dir / normalized_dir).resolve()

    res = ingest_inbox(
        xlsx=xlsx,
        inbox_dir=inbox_dir,
        normalized_dir=normalized_dir,
        src_id=src_id,
        data_origin=data_origin,
        evidence_type=(evidence_type.strip() or None),
        move=(not copy_only),
    )

    console.print(f"[green]OK[/green] ingested={res.get('ingested')} (manifest={res.get('manifest','')})")


@app.command("export-reference-pack")
def export_reference_pack_cmd(
    case_dir: Path = typer.Option(..., exists=True, file_okay=False, dir_okay=True, help="Case folder (contains case.xlsx)."),
    pack_id: str = typer.Option(..., help="Reference pack id (folder name)."),
    out_dir: Path = typer.Option(Path("reference_packs"), help="Parent dir for packs."),
    title: str = typer.Option("", help="Human-friendly title."),
    sheets: str = typer.Option(
        "ENV_BASE_WATER,ENV_BASE_AIR,ENV_BASE_SOCIO",
        help="Comma-separated sheet names to inherit (applied only when target sheet is empty).",
    ),
    copy_attachments: bool = typer.Option(True, help="Copy referenced files from ATTACHMENTS into the pack."),
) -> None:
    """Export a reusable reference pack from an existing case folder."""
    from eia_gen.services.reference_packs import export_reference_pack

    case_dir = case_dir.resolve()
    out_dir = _resolve_repo_path(out_dir)

    sheet_list = [s.strip() for s in (sheets or "").split(",") if s.strip()]
    pack_dir = export_reference_pack(
        case_dir=case_dir,
        out_dir=out_dir,
        pack_id=pack_id,
        title=title or pack_id,
        apply_sheets=sheet_list,
        copy_attachments=copy_attachments,
    )
    console.print(f"[green]OK[/green] wrote {pack_dir}")


@app.command("apply-reference-pack")
def apply_reference_pack_cmd(
    xlsx: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    sources: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    pack_dir: Path = typer.Option(..., exists=True, file_okay=False, dir_okay=True),
) -> None:
    """Apply a reference pack to a case.

    - Copies pack sheets only when the target sheet is empty.
    - Merges sources.yaml (append-only).
    - Copies referenced evidence files when missing.
    """
    from eia_gen.services.reference_packs import apply_reference_pack

    xlsx = xlsx.resolve()
    sources = sources.resolve()
    pack_dir = pack_dir.resolve()

    res = apply_reference_pack(xlsx=xlsx, sources=sources, pack_dir=pack_dir)

    console.print(f"[green]OK[/green] applied pack={res.get('pack_id')}")
    console.print(f"- sheets: applied={len(res.get('applied_sheets', []))} skipped={len(res.get('skipped_sheets', []))}")
    console.print(f"- sources: added={len(res.get('added_sources', []))}")
    console.print(f"- files: copied={len(res.get('copied_files', []))}")



@app.command("check-xlsx-both")
def check_xlsx_both(
    xlsx: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    sources: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    out_dir: Path = typer.Option(Path("output/check")),
    spec_dir_eia: Path = typer.Option(Path("spec")),
    spec_dir_dia: Path = typer.Option(Path("spec_dia")),
    template_eia: Path | None = typer.Option(
        None,
        help="Optional EIA DOCX template to scope QA to anchors present in the template.",
    ),
    template_dia: Path | None = typer.Option(
        None,
        help="Optional DIA DOCX template to scope QA to anchors present in the template.",
    ),
    scope_to_default_templates: bool = typer.Option(
        False,
        help=(
            "When set, picks default templates (same selection as generate-xlsx-both) and scopes QA to template anchors. "
            "Useful to match 'what will actually render' in template mode."
        ),
    ),
    use_llm: bool = typer.Option(False, help="Run writers with LLM (requires OPENAI_API_KEY)."),
    fail_on_warn: bool = typer.Option(False, help="Exit with non-zero code when WARN exists."),
    submission: bool = typer.Option(
        False,
        help="Use strict 'submission mode' rules (treat missing core sheets/rows as ERROR).",
    ),
    write_next_actions: bool = typer.Option(
        False,
        help="Write QA next-actions markdown next to validation_report_*.json (best-effort).",
    ),
) -> None:
    """Run QA only (no DOCX render) for EIA+DIA and write validation reports.

    Notes:
    - Default behavior runs QA across the full spec (all sections).
    - When `--template-eia/--template-dia` is provided (or `--scope-to-default-templates` is set),
      QA is scoped to the anchors actually present in the template (closer to what will render).
    """
    from eia_gen.services.qa.run import run_qa
    from eia_gen.services.writer import SpecReportWriter, WriterOptions
    from eia_gen.services.xlsx.case_reader import load_case_from_xlsx
    from eia_gen.spec.load import load_spec_bundle

    # When called as a plain Python function (e.g., unit tests), Typer leaves defaults
    # as OptionInfo objects. Normalize them to real defaults.
    from typer.models import OptionInfo

    if isinstance(template_eia, OptionInfo):
        template_eia = None
    if isinstance(template_dia, OptionInfo):
        template_dia = None
    if isinstance(scope_to_default_templates, OptionInfo):
        scope_to_default_templates = False
    if isinstance(write_next_actions, OptionInfo):
        write_next_actions = False
    if isinstance(submission, OptionInfo):
        submission = False
    if isinstance(fail_on_warn, OptionInfo):
        fail_on_warn = False
    if isinstance(use_llm, OptionInfo):
        use_llm = False

    xlsx = xlsx.resolve()
    sources = sources.resolve()
    out_dir = _resolve_repo_path(out_dir)
    spec_dir_eia = _resolve_spec_dir(spec_dir_eia)
    spec_dir_dia = _resolve_spec_dir(spec_dir_dia)
    template_eia = _resolve_repo_path(template_eia) if template_eia else None
    template_dia = _resolve_repo_path(template_dia) if template_dia else None

    case_obj = load_case_from_xlsx(xlsx)
    sources_obj = load_sources(sources)

    llm = None
    if use_llm and settings.openai_api_key:
        from eia_gen.services.llm.openai_client import OpenAIChatClient

        llm = OpenAIChatClient(api_key=settings.openai_api_key, model=settings.openai_model)
    elif use_llm:
        console.print("[yellow]LLM disabled[/yellow] (missing OPENAI_API_KEY)")

    spec_eia = load_spec_bundle(spec_dir_eia)
    spec_dia = load_spec_bundle(spec_dir_dia)

    writer_eia = SpecReportWriter(spec=spec_eia, sources=sources_obj, llm=llm, options=WriterOptions(use_llm=use_llm))
    writer_dia = SpecReportWriter(spec=spec_dia, sources=sources_obj, llm=llm, options=WriterOptions(use_llm=use_llm))
    draft_eia = writer_eia.generate(case_obj)
    draft_dia = writer_dia.generate(case_obj)

    data_rules_eia = (_REPO_ROOT / "config/data_acquisition_rules_submission.yaml") if submission else None
    data_rules_dia = (_REPO_ROOT / "config/data_acquisition_rules_submission_dia.yaml") if submission else None

    asset_search_dirs = [xlsx.parent, out_dir, Path.cwd()]
    # Optional: match template-mode rendering by scoping QA to anchors present in templates.
    if scope_to_default_templates:
        if template_eia is None:
            for cand in [
                _REPO_ROOT / "templates/report_template.sample_changwon_2025.ssot_full.scaffolded.docx",
                _REPO_ROOT / "templates/report_template.sample_changwon_2025.scaffolded.docx",
                _REPO_ROOT / "templates/report_template.sample_changwon_2025.docx",
                _REPO_ROOT / "templates/report_template.docx",
            ]:
                if cand.exists():
                    template_eia = cand
                    break
        if template_dia is None:
            for cand in [
                _REPO_ROOT / "templates/dia_template.scaffolded.docx",
                _REPO_ROOT / "templates/dia_template.docx",
            ]:
                if cand.exists():
                    template_dia = cand
                    break

    qa_eia = run_qa(
        case_obj,
        sources_obj,
        draft_eia,
        spec=spec_eia,
        asset_search_dirs=asset_search_dirs,
        template_path=template_eia,
        case_xlsx_path=xlsx,
        data_acquisition_rules_path=data_rules_eia,
        submission_mode=submission,
    )
    qa_dia = run_qa(
        case_obj,
        sources_obj,
        draft_dia,
        spec=spec_dia,
        asset_search_dirs=asset_search_dirs,
        template_path=template_dia,
        case_xlsx_path=xlsx,
        data_acquisition_rules_path=data_rules_dia,
        submission_mode=submission,
    )

    out_dir.mkdir(parents=True, exist_ok=True)
    p_eia = out_dir / "validation_report_eia.json"
    p_dia = out_dir / "validation_report_dia.json"
    p_eia.write_text(json.dumps(qa_eia.model_dump(), indent=2, ensure_ascii=False), encoding="utf-8")
    p_dia.write_text(json.dumps(qa_dia.model_dump(), indent=2, ensure_ascii=False), encoding="utf-8")

    e_err = int((qa_eia.stats or {}).get("error_count", 0))
    d_err = int((qa_dia.stats or {}).get("error_count", 0))
    e_warn = int((qa_eia.stats or {}).get("warn_count", 0))
    d_warn = int((qa_dia.stats or {}).get("warn_count", 0))

    console.print(f"[green]OK[/green] wrote {p_eia} (EIA errors={e_err} warns={e_warn})")
    console.print(f"[green]OK[/green] wrote {p_dia} (DIA errors={d_err} warns={d_warn})")

    if write_next_actions:
        _write_qa_next_actions_md(validation_report=p_eia, out_md=(out_dir / "next_actions_eia.md"), kind="EIA")
        _write_qa_next_actions_md(validation_report=p_dia, out_md=(out_dir / "next_actions_dia.md"), kind="DIA")

    if e_err > 0 or d_err > 0:
        raise typer.Exit(code=1)
    if fail_on_warn and (e_warn > 0 or d_warn > 0):
        raise typer.Exit(code=2)


@app.command("plan-data-requests")
def plan_data_requests(
    xlsx: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    out: Path | None = typer.Option(None, help="Write to a new file (default: in-place update)."),
    overwrite: bool = typer.Option(False, help="Overwrite existing DATA_REQUESTS rows."),
) -> None:
    """Populate DATA_REQUESTS sheet (v2) with a minimal default plan (WMS + simple AUTO_GIS)."""
    from eia_gen.services.data_requests.planner import plan_data_requests_for_workbook
    from eia_gen.services.data_requests.models import DataRequest
    from eia_gen.services.data_requests.xlsx_io import read_data_requests, write_data_requests, load_workbook, save_workbook

    xlsx = xlsx.resolve()
    out_path = out.resolve() if out else xlsx

    wb = load_workbook(xlsx)
    existing = read_data_requests(wb)

    plan = plan_data_requests_for_workbook(wb=wb, wms_layers_config=_REPO_ROOT / "config/wms_layers.yaml")

    if existing and not overwrite:
        # Merge: keep user/PDF_PAGE rows, add missing planner rows.
        # Also, when a planner-managed row was previously disabled due to missing env
        # and hasn't run yet, update enabled flag.
        by_id = {r.req_id: r for r in existing}
        added = 0
        updated = 0
        migrated = 0
        for p in plan:
            e = by_id.get(p.req_id)
            if e is None:
                by_id[p.req_id] = p
                added += 1
                continue

            # Preserve most user edits; only do safe refreshes.
            enabled = e.enabled
            if (not e.last_run_at) and (not e.enabled) and p.enabled:
                enabled = True
                updated += 1

            note = e.note
            if (not note) or ("disabled: missing env" in note):
                note = p.note

            params_json = e.params_json or p.params_json
            params = e.params or p.params
            # Migrate legacy KOSIS request payloads (query_params+mappings placeholders) to dataset_keys mode.
            if e.req_id == "REQ-KOSIS-ENV_BASE_SOCIO":
                try:
                    legacy = dict(e.params or {})
                    has_dataset_keys = bool(legacy.get("dataset_keys") or legacy.get("dataset_key"))
                    legacy_query = legacy.get("query_params")
                    legacy_mappings = legacy.get("mappings")
                    looks_placeholder = False
                    if (not has_dataset_keys) and isinstance(legacy_query, dict) and not legacy_query and isinstance(legacy_mappings, list):
                        # Treat as placeholder when all mapping rules are empty.
                        nonempty = False
                        for m in legacy_mappings:
                            if not isinstance(m, dict):
                                continue
                            if str(m.get("match_itm_id") or m.get("itm_id") or "").strip():
                                nonempty = True
                                break
                            if str(m.get("match_itm_nm_contains") or m.get("itm_nm_contains") or m.get("field") or "").strip():
                                nonempty = True
                                break
                        looks_placeholder = not nonempty

                    if looks_placeholder:
                        params_json = p.params_json
                        params = p.params
                        migrated += 1
                except Exception:
                    pass

            merged = DataRequest(
                req_id=e.req_id,
                enabled=enabled,
                priority=e.priority or p.priority,
                connector=e.connector or p.connector,
                purpose=e.purpose or p.purpose,
                src_id=e.src_id or p.src_id,
                params_json=params_json,
                params=params,
                output_sheet=e.output_sheet or p.output_sheet,
                merge_strategy=e.merge_strategy or p.merge_strategy,
                upsert_keys=e.upsert_keys or p.upsert_keys,
                run_mode=e.run_mode or p.run_mode,
                last_run_at=e.last_run_at,
                last_evidence_ids=e.last_evidence_ids,
                note=note,
            )
            by_id[e.req_id] = merged

        merged_plan = list(by_id.values())
        merged_plan.sort(key=lambda x: (x.priority, x.req_id))
        write_data_requests(wb, merged_plan)
        save_workbook(wb, out_path)
        console.print(
            f"[green]OK[/green] wrote {out_path} (DATA_REQUESTS merged: total={len(merged_plan)} added={added} updated={updated} migrated={migrated})"
        )
        return

    if existing and overwrite:
        console.print("[yellow]OVERWRITE[/yellow] replacing existing DATA_REQUESTS rows")

    plan = plan_data_requests_for_workbook(wb=wb, wms_layers_config=_REPO_ROOT / "config/wms_layers.yaml")
    write_data_requests(wb, plan)
    save_workbook(wb, out_path)
    console.print(f"[green]OK[/green] wrote {out_path} (DATA_REQUESTS rows={len(plan)})")




@app.command("plan-pdf-evidence")
def plan_pdf_evidence(
    xlsx: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    index: Path = typer.Option(
        ...,
        exists=True,
        file_okay=True,
        dir_okay=True,
        help="OCR index file (combined_index.json/pass2_hits.json) or a directory containing them.",
    ),
    src_id: str = typer.Option("S-CHANGWON-SAMPLE", help="Source id for extracted PDF evidence."),
    pdf: Path | None = typer.Option(
        None,
        help="Override PDF path used in requests (default: index.json's pdf_path).",
    ),
    kinds: str = typer.Option("figure,table", help="Comma-separated kinds: figure,table"),
    include_labels: str = typer.Option(
        "",
        help="Optional: semicolon-separated labels to include (exact match). Example: 7.6-3;7.3.1-3",
    ),
    include_text: str = typer.Option(
        "",
        help="Optional: semicolon-separated keywords. Keep hits whose caption contains any keyword.",
    ),
    exclude_text: str = typer.Option(
        "",
        help="Optional: semicolon-separated keywords to exclude.",
    ),
    dpi: int = typer.Option(250, help="Rasterization DPI for PDF_PAGE."),
    enabled: bool = typer.Option(True, help="Set planned requests enabled=true."),
    run_mode: str = typer.Option("ONCE", help="AUTO/ONCE/NEVER"),
    req_prefix: str = typer.Option("REQ-PDF", help="req_id prefix for generated rows."),
    max_items: int = typer.Option(80, help="Safety limit for planned rows."),
    out: Path | None = typer.Option(None, help="Write updated case.xlsx to a new file (default: in-place)."),
) -> None:
    """Create DATA_REQUESTS(PDF_PAGE) rows from OCR index JSON.

    Typical workflow:
    1) Build an OCR index (scanned PDFs):
       `python scripts/extract_pdf_index_twopass.py --pdf <pdf> --out-dir <dir> --page-start ... --page-end ...`
    2) Plan evidence requests:
       `eia-gen plan-pdf-evidence --xlsx case.xlsx --index <dir> --src-id S-CHANGWON-SAMPLE`
    3) Execute once:
       `eia-gen run-data-requests --xlsx case.xlsx`

    The runner will materialize PNGs under `attachments/evidence/pdf/` and register them in `ATTACHMENTS`.
    """
    from eia_gen.services.data_requests.pdf_index import (
        build_pdf_page_data_requests,
        filter_hits,
        load_pdf_index_hits,
    )
    from eia_gen.services.data_requests.xlsx_io import apply_rows_to_sheet, load_workbook, save_workbook

    xlsx = xlsx.resolve()
    out_path = out.resolve() if out else xlsx

    hits, pdf_from_index = load_pdf_index_hits(index)

    pdf_path = ""
    if pdf is not None:
        pdf_path = str(pdf.expanduser().resolve())
    elif pdf_from_index:
        raw = str(pdf_from_index).strip()
        if raw:
            p = Path(raw).expanduser()
            # For relative paths inside index files, prefer resolving from repo root first.
            # This makes reference packs portable across machines/cwd.
            candidates: list[Path] = []
            if p.is_absolute():
                candidates.append(p)
            else:
                candidates.append((_REPO_ROOT / p).expanduser())
                base = index.expanduser().resolve()
                base_dir = base.parent if base.is_file() else base
                candidates.append((base_dir / p).expanduser())
                candidates.append(p)

            chosen: Path | None = None
            for c in candidates:
                try:
                    if c.exists():
                        chosen = c
                        break
                except Exception:
                    continue
            chosen = chosen or candidates[-1]
            try:
                pdf_path = str(chosen.resolve())
            except Exception:
                pdf_path = str(chosen)

    if not pdf_path:
        raise ValueError("PDF path not found. Provide --pdf or use index.json created with --pdf.")

    kind_set = {k.strip().lower() for k in (kinds or "").split(",") if k.strip()}
    labels = {s.strip() for s in (include_labels or "").split(";") if s.strip()} or None
    include_any = [s.strip() for s in (include_text or "").split(";") if s.strip()]
    exclude_any = [s.strip() for s in (exclude_text or "").split(";") if s.strip()]

    selected = filter_hits(
        hits,
        kinds=kind_set or {"figure", "table"},
        include_labels=labels,
        include_text_any=include_any,
        exclude_text_any=exclude_any,
    )
    if not selected:
        console.print("[yellow]SKIP[/yellow] no matching hits")
        return

    selected = selected[: max(0, int(max_items))]
    rows = build_pdf_page_data_requests(
        selected,
        pdf_path=pdf_path,
        src_id=src_id,
        req_prefix=req_prefix,
        enabled=enabled,
        run_mode=run_mode,
        dpi=dpi,
    )

    wb = load_workbook(xlsx)
    # UPSERT by req_id to keep sheet validations/styles.
    sheet_warn = apply_rows_to_sheet(
        wb,
        sheet_name="DATA_REQUESTS",
        rows=rows,
        merge_strategy="UPSERT_KEYS",
        upsert_keys=["req_id"],
    )
    save_workbook(wb, out_path)

    console.print(f"[green]OK[/green] updated {out_path} (planned={len(rows)})")
    if sheet_warn:
        console.print(f"[yellow]WARN[/yellow] {len(sheet_warn)} warnings")

@app.command("run-data-requests")
def run_data_requests_cmd(
    xlsx: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    out: Path | None = typer.Option(None, help="Write updated case.xlsx to a new file (default: in-place)."),
) -> None:
    """Execute enabled DATA_REQUESTS and apply results (WMS evidence + AUTO_GIS calculations)."""
    import json

    from eia_gen.services.data_requests.runner import run_data_requests
    from eia_gen.services.data_requests.xlsx_io import load_workbook, save_workbook

    xlsx = xlsx.resolve()
    out_path = out.resolve() if out else xlsx
    case_dir = xlsx.parent.resolve()

    wb = load_workbook(xlsx)
    result = run_data_requests(
        wb=wb,
        case_dir=case_dir,
        wms_layers_config=_REPO_ROOT / "config/wms_layers.yaml",
        cache_config=_REPO_ROOT / "config/cache.yaml",
    )
    save_workbook(wb, out_path)

    report_path = case_dir / "_data_requests_run.json"
    report_path.write_text(
        json.dumps(
            {
                "executed": result.executed,
                "skipped": result.skipped,
                "warnings": result.warnings,
                "evidences": [e.__dict__ for e in result.evidences],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    console.print(f"[green]OK[/green] updated {out_path}")
    console.print(f"[green]OK[/green] wrote {report_path}")
    if result.warnings:
        console.print(f"[yellow]WARN[/yellow] {len(result.warnings)} warnings (see report)")


@app.command("wire-wms-fallback")
def wire_wms_fallback_cmd(
    xlsx: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    out: Path | None = typer.Option(None, help="Write updated case.xlsx to a new file (default: in-place)."),
    enable_when_wired: bool = typer.Option(
        False,
        help="When wired, also set enabled=true for previously disabled WMS rows (useful for offline fallback).",
    ),
    force: bool = typer.Option(
        False,
        help="Unsafe: allow wiring even when bbox/srs mismatch (can cause wrong map evidence usage).",
    ),
) -> None:
    """Autofill WMS fallback_file_path from local evidences in ATTACHMENTS.

    This is a best-effort helper for offline/on-prem runs when WMS keys/approvals are blocked.
    It only wires a fallback when it can match bbox/srs between:
    - the WMS request (computed bbox from LOCATION + request params), and
    - an existing `attachments/evidence/wms/*.png` evidence registered in ATTACHMENTS.
    """
    from eia_gen.services.data_requests.wms_fallback import wire_wms_fallbacks_for_workbook
    from eia_gen.services.data_requests.xlsx_io import load_workbook, save_workbook

    xlsx = xlsx.resolve()
    out_path = out.resolve() if out else xlsx
    case_dir = xlsx.parent.resolve()

    wb = load_workbook(xlsx)
    rep = wire_wms_fallbacks_for_workbook(
        wb=wb,
        case_dir=case_dir,
        enable_when_wired=enable_when_wired,
        force=force,
    )
    save_workbook(wb, out_path)

    console.print(f"[green]OK[/green] updated {out_path} (wired={len(rep.updated_req_ids)})")
    if rep.enabled_req_ids:
        console.print(f"[green]OK[/green] enabled {len(rep.enabled_req_ids)} WMS rows")
    if rep.skipped:
        console.print(f"[yellow]WARN[/yellow] skipped {len(rep.skipped)} WMS rows (no safe match)")


class VerifyKeysMode(str, Enum):
    network = "network"
    presence = "presence"


@app.command("verify-keys")
def verify_keys(
    mode: VerifyKeysMode = typer.Option(
        VerifyKeysMode.network,
        help="network=call endpoints, presence=env-only (no network calls)",
    ),
    center_lon: float = typer.Option(126.9780, help="WGS84 lon for checks (used by AirKorea/WMS)."),
    center_lat: float = typer.Option(37.5665, help="WGS84 lat for checks (used by AirKorea/WMS)."),
    wms_radius_m: int = typer.Option(1000, help="Radius used to build a WMS bbox for checks."),
    timeout_sec: int = typer.Option(12, help="HTTP timeout in seconds."),
    wms_use_cache: bool = typer.Option(
        False,
        help="Use cached WMS images when available (default: force network call for key validation).",
    ),
    out: Path | None = typer.Option(None, help="Write a JSON report (never includes key values)."),
    strict: bool = typer.Option(False, help="Exit non-zero if any required check fails."),
    strict_ignore_category: list[str] = typer.Option(
        [],
        help="When --strict, ignore failures with these categories (repeatable).",
    ),
) -> None:
    """Verify API keys/endpoints with minimal network calls (never prints key values).

    Checks (best-effort):
    - data.go.kr key: KMA ASOS daily + AirKorea nearby station
    - SAFEMAP key: WMS GetMap for FLOOD_TRACE/LANDSLIDE_RISK
    - KOSIS key + endpoint smoke (Param/statisticsParameterData.do, 1-cell)
    """
    from datetime import datetime, timedelta
    import json

    import httpx
    from pyproj import CRS, Transformer

    from eia_gen.services.data_requests.data_go_kr import build_url
    from eia_gen.services.data_requests.kma_asos import KMA_ASOS_DAILY_URL
    from eia_gen.services.data_requests.sanitize import redact_text
    from eia_gen.services.data_requests.wms import fetch_wms

    def _present(env: str) -> bool:
        import os

        return bool(os.environ.get(env, "").strip())

    def _kma_key() -> str:
        import os

        return (os.environ.get("KMA_API_KEY") or os.environ.get("DATA_GO_KR_SERVICE_KEY") or "").strip()

    def _airkorea_key() -> str:
        import os

        return (os.environ.get("AIRKOREA_API_KEY") or os.environ.get("DATA_GO_KR_SERVICE_KEY") or "").strip()

    def _wms_bbox_3857(lon: float, lat: float, radius_m: int) -> tuple[float, float, float, float]:
        t = Transformer.from_crs(CRS.from_epsg(4326), CRS.from_epsg(3857), always_xy=True)
        x, y = t.transform(lon, lat)
        r = float(max(0, int(radius_m)))
        return (float(x - r), float(y - r), float(x + r), float(y + r))

    results: list[dict[str, object]] = []

    def _record(name: str, status: str, msg: str, *, category: str | None = None) -> None:
        rec: dict[str, object] = {"name": name, "status": status, "message": msg}
        if category:
            rec["category"] = category
        results.append(rec)

    def _print_status(name: str, ok: bool, msg: str, *, category: str | None = None) -> None:
        tag = "[green]OK[/green]" if ok else "[red]FAIL[/red]"
        console.print(f"- {name}: {tag} {msg}")
        _record(name, "ok" if ok else "fail", msg, category=category)

    def _print_skip(name: str, msg: str, *, category: str | None = None) -> None:
        console.print(f"- {name}: [yellow]SKIP[/yellow] {msg}")
        _record(name, "skip", msg, category=category)

    def _strict_failures() -> list[dict[str, object]]:
        ignore_cats = {str(c or "").strip() for c in strict_ignore_category if str(c or "").strip()}
        nonfatal_fail_names = {
            # AirKorea station lookup can be blocked even when measurement endpoint works.
            # Keep this check visible, but don't fail strict mode on it.
            "AirKorea STATION (data.go.kr)",
            # EcoZmp (ecology map/WMS조회) is optional and frequently needs per-API approval on data.go.kr.
            # Keep visible, but don't block strict mode.
            "EcoZmp WMS (data.go.kr)",
        }
        failed = []
        for r in results:
            if r.get("status") != "fail":
                continue
            if r.get("name") in nonfatal_fail_names:
                continue
            if ignore_cats and (r.get("category") in ignore_cats):
                continue
            failed.append(r)
        return failed

    console.print("[bold]Key/Endpoint Checks[/bold]")

    if mode == VerifyKeysMode.presence:
        key_datago = _kma_key()
        if not key_datago:
            _print_status(
                "KMA ASOS (data.go.kr)",
                False,
                "missing `DATA_GO_KR_SERVICE_KEY` (or `KMA_API_KEY`)",
                category="missing_key",
            )
        else:
            _print_status("KMA ASOS (data.go.kr)", True, "key present (network skipped)", category="present")

        key_air = _airkorea_key()
        if not key_air:
            _print_status(
                "AirKorea (data.go.kr)",
                False,
                "missing `DATA_GO_KR_SERVICE_KEY` (or `AIRKOREA_API_KEY`)",
                category="missing_key",
            )
        else:
            _print_status("AirKorea (data.go.kr)", True, "key present (network skipped)", category="present")

        if not _present("SAFEMAP_API_KEY"):
            _print_status("SAFEMAP WMS", False, "missing `SAFEMAP_API_KEY`", category="missing_key")
        else:
            _print_status("SAFEMAP WMS", True, "key present (network skipped)", category="present")

        # --- KOSIS ---
        if _present("KOSIS_API_KEY"):
            msg = "key present (note: DATA_REQUESTS에 orgId/tblId/itmId 등 query_params가 필요)"
            console.print(f"- KOSIS: [green]OK[/green] {msg}")
            _record("KOSIS", "ok", msg, category="present")
        else:
            _print_skip("KOSIS", "missing `KOSIS_API_KEY` (통계 자동수집은 선택)", category="missing_key_optional")

        # --- VWORLD (optional) ---
        if _present("VWORLD_API_KEY"):
            msg = "key present (geocode quality improves)"
            console.print(f"- VWORLD: [green]OK[/green] {msg}")
            _record("VWORLD", "ok", msg, category="present")
        else:
            _print_skip("VWORLD", "missing `VWORLD_API_KEY` (optional)", category="missing_key_optional")

        if out:
            out_path = _resolve_repo_path(out)
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_text(
                json.dumps(
                    {
                        "mode": mode.value,
                        "center_lon": float(center_lon),
                        "center_lat": float(center_lat),
                        "wms_radius_m": int(wms_radius_m),
                        "timeout_sec": int(timeout_sec),
                        "wms_use_cache": bool(wms_use_cache),
                        "results": results,
                    },
                    ensure_ascii=False,
                    indent=2,
                ),
                encoding="utf-8",
            )
            console.print(f"[green]OK[/green] wrote {out_path}")

        if strict:
            if _strict_failures():
                raise typer.Exit(code=1)
        return

    # --- data.go.kr / KMA ASOS ---
    key_datago = _kma_key()
    if not key_datago:
        _print_status(
            "KMA ASOS (data.go.kr)",
            False,
            "missing `DATA_GO_KR_SERVICE_KEY` (or `KMA_API_KEY`)",
            category="missing_key",
        )
    else:
        # KMA ASOS daily API typically serves data up to *yesterday*.
        # If we ask for today, it may return resultCode=99 ("전날 자료까지 제공...").
        end_dt = (datetime.now() - timedelta(days=1)).strftime("%Y%m%d")
        start_dt = (datetime.now() - timedelta(days=4)).strftime("%Y%m%d")
        params = {
            "pageNo": 1,
            "numOfRows": 1,
            "dataType": "JSON",
            "dataCd": "ASOS",
            "dateCd": "DAY",
            "startDt": start_dt,
            "endDt": end_dt,
            "stnIds": "108",  # 서울(예시)
        }
        url = build_url(base_url=KMA_ASOS_DAILY_URL, service_key=key_datago, params=params, key_param="serviceKey")
        try:
            with httpx.Client(timeout=timeout_sec, follow_redirects=True) as client:
                r = client.get(url)
                r.raise_for_status()
                content_type = (r.headers.get("content-type") or "").split(";")[0].strip().lower()
                if content_type != "application/json":
                    text = (r.text or "").strip()
                    is_html_503 = ("service unavailable" in text.lower()) and ("503" in text)
                    if is_html_503:
                        _print_status(
                            "KMA ASOS (data.go.kr)",
                            False,
                            "HTTP 200 but HTML '503 Service Unavailable' (data.go.kr/KMA outage suspected)",
                            category="outage",
                        )
                        console.print("  hint: 키/승인 문제가 아니라 서비스 장애일 수 있어, 시간 두고 재시도하세요.")
                    else:
                        snippet = " ".join(text.split())[:160]
                        _print_status(
                            "KMA ASOS (data.go.kr)",
                            False,
                            f"HTTP {r.status_code} unexpected content-type={content_type or 'unknown'} (snippet={snippet})",
                            category="invalid_response",
                        )
                        console.print("  hint: data.go.kr에서 해당 API 활용신청/승인 여부 및 'Decoding(일반)' 키 사용 여부를 확인하세요.")
                else:
                    data = r.json()
                    header = ((data or {}).get("response") or {}).get("header") or {}
                    code = str(header.get("resultCode") or "").strip()
                    msg = str(header.get("resultMsg") or "").strip()
                    if code and code != "00":
                        _print_status(
                            "KMA ASOS (data.go.kr)", False, f"resultCode={code} {msg}", category="api_error"
                        )
                    else:
                        _print_status(
                            "KMA ASOS (data.go.kr)",
                            True,
                            f"HTTP 200 (station=108, {start_dt}..{end_dt})",
                            category="ok",
                        )
        except Exception as e:
            _print_status("KMA ASOS (data.go.kr)", False, redact_text(str(e)), category="network_error")
            console.print("  hint: data.go.kr에서 해당 API 활용신청/승인 여부 및 'Decoding(일반)' 키 사용 여부를 확인하세요.")

    # --- AirKorea (data.go.kr) ---
    key_air = _airkorea_key()
    if not key_air:
        _print_status(
            "AirKorea (data.go.kr)",
            False,
            "missing `DATA_GO_KR_SERVICE_KEY` (or `AIRKOREA_API_KEY`)",
            category="missing_key",
        )
    else:
        station_name: str | None = None
        try:
            # AirKorea `getNearbyMsrstnList` expects TM coordinates; use EPSG:5181 (Kakao/Daum TM).
            t = Transformer.from_crs(CRS.from_epsg(4326), CRS.from_epsg(5181), always_xy=True)
            tm_x, tm_y = t.transform(float(center_lon), float(center_lat))
            base_url = "http://apis.data.go.kr/B552584/MsrstnInfoInqireSvc/getNearbyMsrstnList"
            params = {"returnType": "json", "tmX": f"{tm_x:.3f}", "tmY": f"{tm_y:.3f}"}
            url = build_url(base_url=base_url, service_key=key_air, params=params, key_param="serviceKey")
            with httpx.Client(timeout=timeout_sec, follow_redirects=True) as client:
                r = client.get(url)
                r.raise_for_status()
                data = r.json()
            items = (((data or {}).get("response") or {}).get("body") or {}).get("items") or []
            if not isinstance(items, list) or not items:
                _print_status(
                    "AirKorea STATION (data.go.kr)",
                    False,
                    "no stations returned (unexpected response shape)",
                    category="unexpected_response",
                )
            else:
                st = items[0] or {}
                station_name = str(st.get("stationName") or "").strip() or None
                station = station_name or "(unknown)"
                _print_status(
                    "AirKorea STATION (data.go.kr)",
                    True,
                    f"HTTP 200 (nearest station={station})",
                    category="ok",
                )
        except Exception as e:
            _print_status("AirKorea STATION (data.go.kr)", False, redact_text(str(e)), category="network_error")
            console.print(
                "  hint: AirKorea는 '측정소정보(MsrstnInfoInqireSvc)'와 '대기오염정보(ArpltnInforInqireSvc)'가 "
                "별도 승인/권한일 수 있습니다."
            )

        # Measurement endpoint check (can be OK even when station lookup is forbidden).
        try:
            station_for_check = station_name or "종로구"
            fallback = "" if station_name else " (fallback station=종로구)"
            base_url = "http://apis.data.go.kr/B552584/ArpltnInforInqireSvc/getMsrstnAcctoRltmMesureDnsty"
            params = {
                "returnType": "json",
                "numOfRows": 1,
                "pageNo": 1,
                "stationName": station_for_check,
                "dataTerm": "MONTH",
                "ver": "1.3",
            }
            url = build_url(base_url=base_url, service_key=key_air, params=params, key_param="serviceKey")
            with httpx.Client(timeout=timeout_sec, follow_redirects=True) as client:
                r = client.get(url)
                r.raise_for_status()
                data = r.json()

            header = ((data or {}).get("response") or {}).get("header") or {}
            code = str(header.get("resultCode") or "").strip()
            msg = str(header.get("resultMsg") or "").strip()
            if code and code != "00":
                _print_status(
                    "AirKorea MEASURE (data.go.kr)", False, f"resultCode={code} {msg}", category="api_error"
                )
            else:
                _print_status(
                    "AirKorea MEASURE (data.go.kr)",
                    True,
                    f"HTTP 200 (station={station_for_check}){fallback}",
                    category="ok",
                )
        except Exception as e:
            _print_status("AirKorea MEASURE (data.go.kr)", False, redact_text(str(e)), category="network_error")
            console.print("  hint: data.go.kr에서 AirKorea(ArpltnInforInqireSvc) 활용신청/승인 여부를 확인하세요.")

    # --- EcoZmp WMS (data.go.kr, optional) ---
    # NOTE: Public Data Portal (data.go.kr) often requires separate "utilization approval" per API.
    # This check is informative and does not fail strict mode (see nonfatal_fail_names above).
    if key_datago:
        try:
            t = Transformer.from_crs(CRS.from_epsg(4326), CRS.from_epsg(5186), always_xy=True)
            x, y = t.transform(float(center_lon), float(center_lat))
            r = float(max(0, int(wms_radius_m)))
            bbox_5186 = (float(x - r), float(y - r), float(x + r), float(y + r))
            fetched = fetch_wms(
                layer_key="ECO_NATURE_DATAGO_OPEN",
                bbox=bbox_5186,
                width=512,
                height=512,
                out_srs="EPSG:5186",
                wms_layers_config=_REPO_ROOT / "config/wms_layers.yaml",
                cache_config=_REPO_ROOT / "config/cache.yaml",
                force_refresh=not wms_use_cache,
            )
            _print_status(
                "EcoZmp WMS (data.go.kr)",
                True,
                f"image/png (cache_hit={fetched.cache_hit})",
                category="ok",
            )
        except Exception as e:
            raw = str(e)
            cat = "network_error"
            if "HTTP 401" in raw or "HTTP 403" in raw or "Forbidden" in raw:
                cat = "unauthorized"
            elif "HTTP 404" in raw:
                cat = "endpoint_error"
            elif "non-image content" in raw or "invalid image bytes" in raw:
                cat = "invalid_response"
            _print_status("EcoZmp WMS (data.go.kr)", False, redact_text(raw), category=cat)
            console.print("  hint: 공공데이터포털 '생태자연도 서비스'는 별도 활용신청/승인이 필요할 수 있습니다(403=승인/권한).")

    # --- SAFEMAP WMS ---
    if not _present("SAFEMAP_API_KEY"):
        _print_status("SAFEMAP WMS", False, "missing `SAFEMAP_API_KEY`", category="missing_key")
    else:
        bbox = _wms_bbox_3857(float(center_lon), float(center_lat), int(wms_radius_m))
        for layer_key in ("FLOOD_TRACE", "LANDSLIDE_RISK"):
            try:
                fetched = fetch_wms(
                    layer_key=layer_key,
                    bbox=bbox,
                    width=512,
                    height=512,
                    out_srs="EPSG:3857",
                    wms_layers_config=_REPO_ROOT / "config/wms_layers.yaml",
                    cache_config=_REPO_ROOT / "config/cache.yaml",
                    force_refresh=not wms_use_cache,
                )
                _print_status(
                    f"SAFEMAP WMS:{layer_key}",
                    True,
                    f"image/png (cache_hit={fetched.cache_hit})",
                    category="ok",
                )
            except Exception as e:
                raw = str(e)
                cat = "network_error"
                if "키 값이 맞지 않습니다" in raw or "키값이 맞지 않습니다" in raw:
                    cat = "unauthorized"
                elif "에러 : 생활안전지도" in raw:
                    cat = "unauthorized"
                elif "HTTP 401" in raw or "HTTP 403" in raw:
                    cat = "unauthorized"
                elif "HTTP 404" in raw:
                    cat = "endpoint_error"
                elif "non-image content" in raw or "invalid image bytes" in raw:
                    cat = "invalid_response"
                _print_status(f"SAFEMAP WMS:{layer_key}", False, redact_text(raw), category=cat)
                console.print("  hint: 생활안전지도 OpenAPI 키는 서비스별 승인/제한이 있을 수 있습니다.")

    # --- KOSIS ---
    if not _present("KOSIS_API_KEY"):
        _print_skip("KOSIS", "missing `KOSIS_API_KEY` (통계 자동수집은 선택)", category="missing_key_optional")
    elif mode == VerifyKeysMode.presence:
        msg = "key present"
        console.print(f"- KOSIS: [green]OK[/green] {msg}")
        _record("KOSIS", "ok", msg, category="present")
    else:
        # Quick smoke check (no user inputs): 전국(00) + 총인구(T100) 1개 셀만 조회해
        # 키/엔드포인트/응답 형태를 확인한다.
        try:
            import os

            key_kosis = os.environ.get("KOSIS_API_KEY", "").strip()
            base_url = "https://kosis.kr/openapi/Param/statisticsParameterData.do"
            params = {
                "method": "getList",
                "apiKey": key_kosis,
                "format": "json",
                "jsonVD": "Y",
                "orgId": "101",
                "tblId": "DT_1IN1502",
                "prdSe": "Y",
                "startPrdDe": str(datetime.now().year - 2),
                "endPrdDe": str(datetime.now().year - 2),
                "objL1": "00",
                "itmId": "T100",
            }
            with httpx.Client(timeout=timeout_sec, follow_redirects=True) as client:
                r = client.get(base_url, params=params)
                r.raise_for_status()
                data = r.json()
            if isinstance(data, dict) and str(data.get("err") or "").strip():
                raise ValueError(f"err={data.get('err')} {data.get('errMsg')}")
            if not isinstance(data, list) or not data:
                raise ValueError("empty/invalid JSON list")
            _print_status("KOSIS", True, "HTTP 200 (Param/statisticsParameterData.do)", category="ok")
        except Exception as e:
            _print_status("KOSIS", False, redact_text(str(e)), category="network_error")
            console.print("  hint: KOSIS는 `openapi/Param/statisticsParameterData.do` 엔드포인트가 필요합니다(SSOT: config/kosis_datasets.yaml).")

    # --- VWORLD (optional) ---
    if _present("VWORLD_API_KEY"):
        msg = "key present (geocode quality improves)"
        console.print(f"- VWORLD: [green]OK[/green] {msg}")
        _record("VWORLD", "ok", msg, category="present")
    else:
        _print_skip("VWORLD", "missing `VWORLD_API_KEY` (optional)", category="missing_key_optional")

    if out:
        out_path = _resolve_repo_path(out)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(
            json.dumps(
                {
                    "mode": mode.value,
                    "center_lon": float(center_lon),
                    "center_lat": float(center_lat),
                    "wms_radius_m": int(wms_radius_m),
                    "timeout_sec": int(timeout_sec),
                    "wms_use_cache": bool(wms_use_cache),
                    "results": results,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        console.print(f"[green]OK[/green] wrote {out_path}")

    if strict:
        if _strict_failures():
            raise typer.Exit(code=1)


@app.command("fetch-kma-asos-stations")
def fetch_kma_asos_stations(
    out: Path = typer.Option(Path("config/stations/kma_asos_stations.csv"), help="Output CSV path."),
    overwrite: bool = typer.Option(False, help="Overwrite the existing file."),
) -> None:
    """Fetch and cache KMA ASOS station catalog (id/name/lat/lon).

    This enables the planner to auto-pick the nearest ASOS station for `KMA_ASOS` requests.
    Requires `KMA_API_KEY` or `DATA_GO_KR_SERVICE_KEY`.
    """
    from eia_gen.services.data_requests.kma_stations import (
        evidence_bytes,
        fetch_asos_station_catalog,
        write_asos_station_catalog_csv,
    )

    out = _resolve_repo_path(out)
    if out.exists() and not overwrite:
        console.print(f"[yellow]SKIP[/yellow] {out} already exists (use --overwrite to replace)")
        return

    try:
        cat = fetch_asos_station_catalog()
    except Exception as e:
        from eia_gen.services.data_requests.sanitize import redact_text

        console.print(f"[red]ERROR[/red] fetch-kma-asos-stations failed: {redact_text(str(e))}")
        raise typer.Exit(code=1)
    write_asos_station_catalog_csv(out, cat.stations)

    ev_path = out.with_suffix(".evidence.json")
    ev_path.write_bytes(evidence_bytes(cat.evidence_json))

    console.print(f"[green]OK[/green] wrote {out} (stations={len(cat.stations)})")
    console.print(f"[green]OK[/green] wrote {ev_path}")


@app.command("generate-draft")
def generate_draft(
    case: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    sources: Path = typer.Option(..., exists=True, file_okay=True, dir_okay=False),
    out: Path = typer.Option(Path("output/draft.json")),
    use_llm: bool = typer.Option(True),
    spec_dir: Path = typer.Option(Path("spec")),
) -> None:
    """Generate section drafts (JSON) without DOCX rendering."""
    case_obj = load_case(case)
    sources_obj = load_sources(sources)

    llm = None
    if use_llm and settings.openai_api_key:
        from eia_gen.services.llm.openai_client import OpenAIChatClient

        llm = OpenAIChatClient(api_key=settings.openai_api_key, model=settings.openai_model)
    elif use_llm:
        console.print("[yellow]LLM disabled[/yellow] (missing OPENAI_API_KEY)")

    from eia_gen.services.writer import ReportWriter, SpecReportWriter, WriterOptions

    spec_bundle = None
    if spec_dir.exists():
        from eia_gen.spec.load import load_spec_bundle

        spec_bundle = load_spec_bundle(spec_dir)

    if spec_bundle is not None:
        writer = SpecReportWriter(
            spec=spec_bundle,
            sources=sources_obj,
            llm=llm,
            options=WriterOptions(use_llm=use_llm),
        )
    else:
        writer = ReportWriter(sources=sources_obj, llm=llm, options=WriterOptions(use_llm=use_llm))

    draft = writer.generate(case_obj)

    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(draft.model_dump(), indent=2, ensure_ascii=False), encoding="utf-8")
    from eia_gen.services.qa.run import run_qa

    qa = run_qa(
        case_obj,
        sources_obj,
        draft,
        spec=spec_bundle,
        asset_search_dirs=[case.parent, out.parent, Path.cwd()],
    )
    qa_path = out.parent / "validation_report.json"
    qa_path.write_text(json.dumps(qa.model_dump(), indent=2, ensure_ascii=False), encoding="utf-8")

    from eia_gen.services.export.source_register_xlsx import build_source_register_xlsx_bytes

    xlsx_path = out.parent / "source_register.xlsx"
    report_tag = "DIA" if spec_dir.name == "spec_dia" else "EIA"
    xlsx_path.write_bytes(
        build_source_register_xlsx_bytes(
            case_obj,
            sources_obj,
            draft,
            validation_reports=[(report_tag, qa)],
            report_tag=report_tag,
        )
    )
    console.print(f"[green]OK[/green] wrote {out}")
    console.print(f"[green]OK[/green] wrote {qa_path}")
    console.print(f"[green]OK[/green] wrote {xlsx_path}")
