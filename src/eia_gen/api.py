from __future__ import annotations

import io
import json
import tempfile
import zipfile
from pathlib import Path
from typing import Any

import yaml
from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.responses import StreamingResponse

from eia_gen.config import settings
from eia_gen.models.case import Case
from eia_gen.models.sources import SourceRegistry
from eia_gen.services.canonicalize import canonicalize_case
from eia_gen.services.docx.builder import build_docx
from eia_gen.services.qa.run import run_qa
from eia_gen.services.writer import ReportWriter, SpecReportWriter, WriterOptions
from eia_gen.spec.load import load_spec_bundle

app = FastAPI(title="eia-gen", version="0.1.0")


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


def _safe_extract_zip(zf: zipfile.ZipFile, dest: Path) -> None:
    dest_resolved = dest.resolve()
    for info in zf.infolist():
        name = info.filename
        if name.endswith("/"):
            continue
        target = (dest / name).resolve()
        if not str(target).startswith(str(dest_resolved)):
            raise HTTPException(status_code=400, detail="Invalid zip entry path")
        target.parent.mkdir(parents=True, exist_ok=True)
        with zf.open(info) as src, open(target, "wb") as dst:
            dst.write(src.read())


def _load_yaml_bytes(b: bytes) -> Any:
    try:
        return yaml.safe_load(b.decode("utf-8"))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid YAML: {type(e).__name__}") from e


def _looks_like_xlsx(filename: str | None, content_type: str | None, data: bytes) -> bool:
    name = (filename or "").lower().strip()
    ctype = (content_type or "").lower().strip()
    if name.endswith(".xlsx"):
        return True
    if ctype in {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/octet-stream",
    } and data[:2] == b"PK":
        return True
    return False


@app.post("/v1/reports/small-eia:generate")
async def generate_small_eia(
    case_file: UploadFile = File(...),
    sources_file: UploadFile = File(...),
    assets_zip: UploadFile | None = File(None),
    template_file: UploadFile | None = File(None),
    use_llm: bool = Query(True),
    use_template_map: bool = Query(False),
) -> StreamingResponse:
    case_bytes = await case_file.read()
    sources_raw = _load_yaml_bytes(await sources_file.read())

    llm = None
    if use_llm and settings.openai_api_key:
        from eia_gen.services.llm.openai_client import OpenAIChatClient

        llm = OpenAIChatClient(api_key=settings.openai_api_key, model=settings.openai_model)

    spec_dir = Path("spec")
    spec_bundle = load_spec_bundle(spec_dir) if spec_dir.exists() else None

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)

        try:
            if _looks_like_xlsx(case_file.filename, case_file.content_type, case_bytes):
                from eia_gen.services.xlsx.case_reader import load_case_from_xlsx

                case_path = tmp / "case.xlsx"
                case_path.write_bytes(case_bytes)
                case = load_case_from_xlsx(case_path)
            else:
                case_raw = _load_yaml_bytes(case_bytes)
                case = Case.model_validate(case_raw or {})
                case = canonicalize_case(case)

            sources = SourceRegistry.model_validate(sources_raw or {})
        except HTTPException:
            raise
        except Exception as e:
            raise HTTPException(status_code=422, detail=str(e)) from e

        # Assets zip (optional)
        if assets_zip is not None:
            data = await assets_zip.read()
            try:
                with zipfile.ZipFile(io.BytesIO(data)) as zf:
                    _safe_extract_zip(zf, tmp / "assets")
            except HTTPException:
                raise
            except Exception as e:
                raise HTTPException(status_code=400, detail=f"Invalid assets zip: {type(e).__name__}") from e

            for a in case.assets:
                p = Path(a.file_path)
                if not p.is_absolute():
                    a.file_path = str((tmp / "assets" / p).resolve())

        # Template (optional)
        template_path: Path | None = None
        if template_file is not None:
            template_path = tmp / "template.docx"
            template_path.write_bytes(await template_file.read())

        writer = (
            SpecReportWriter(spec=spec_bundle, sources=sources, llm=llm, options=WriterOptions(use_llm=use_llm))
            if spec_bundle is not None
            else ReportWriter(sources=sources, llm=llm, options=WriterOptions(use_llm=use_llm))
        )
        draft = writer.generate(case)
        qa = run_qa(
            case,
            sources,
            draft,
            spec=spec_bundle,
            asset_search_dirs=[tmp / "assets", tmp, Path.cwd()],
            template_path=template_path,
        )

        out_docx = tmp / "report.docx"
        build_docx(
            case=case,
            sources=sources,
            draft=draft,
            out_path=out_docx,
            template_path=template_path,
            spec_dir=spec_dir if spec_dir.exists() else None,
            use_template_map=use_template_map,
        )

        from eia_gen.services.export.source_register_xlsx import build_source_register_xlsx_bytes

        report_tag = "DIA" if spec_dir.name == "spec_dia" else "EIA"
        out_xlsx = build_source_register_xlsx_bytes(
            case,
            sources,
            draft,
            validation_reports=[(report_tag, qa)],
            report_tag=report_tag,
        )

        bundle = io.BytesIO()
        with zipfile.ZipFile(bundle, mode="w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr("report.docx", out_docx.read_bytes())
            z.writestr("validation_report.json", json.dumps(qa.model_dump(), indent=2, ensure_ascii=False))
            z.writestr("draft.json", json.dumps(draft.model_dump(), indent=2, ensure_ascii=False))
            z.writestr("source_register.xlsx", out_xlsx)

        bundle.seek(0)
        headers = {"Content-Disposition": 'attachment; filename="small_eia_bundle.zip"'}
        return StreamingResponse(bundle, media_type="application/zip", headers=headers)
