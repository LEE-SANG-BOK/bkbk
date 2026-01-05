# eia-gen

관광농원 **소규모환경영향평가서(DOCX) 자동작성** MVP.

## Quickstart

### 1) 설치
- Python `3.11+`
- (권장) 가상환경

```bash
python -m venv .venv
source .venv/bin/activate
pip install -U pip
pip install -e .
```

### 1-1) (권장) 공공 API 키 설정(.env.local)
`DATA_REQUESTS`로 공공DB/API/WMS를 자동 수집하려면 키가 필요합니다.

```bash
cp .env.example .env.local
# .env.local에 키 입력(값 자체는 절대 커밋/공유 금지)
```

키/권한(활용신청/승인) 상태를 빠르게 점검:
```bash
python -m eia_gen verify-keys
```

주의:
- `DATA_GO_KR_SERVICE_KEY`: data.go.kr에서 **해당 API(기상청/에어코리아 등) 활용신청/승인**이 되어야 403이 사라집니다.
- data.go.kr는 보통 **Decoding(일반) 키**를 쓰는 편이 이중 인코딩 이슈가 적습니다.
- `SAFEMAP_API_KEY`: 레이어/서비스별 승인/제한이 있어 WMS가 일부만 동작할 수 있습니다.

### 2) 샘플로 DOCX 생성
```bash
eia-gen generate \
  --case examples/case.sample.yaml \
  --sources examples/sources.sample.yaml \
  --out output/report.docx
```
산출물:
- `output/report.docx`
- `output/validation_report.json`
- `output/source_register.xlsx`
  - 시트: `Source Register`(출처 원장/사용처), `Claims`(문장/표/그림 단위 연결)

### 2-1) (선택) 템플릿(DOCX) 앵커 기반 생성
```bash
# 템플릿 재생성(필요 시)
python scripts/make_template.py --out templates/report_template.docx

eia-gen generate \
  --case examples/case.sample.yaml \
  --sources examples/sources.sample.yaml \
  --template templates/report_template.docx \
  --use-template-map \
  --out output/report_from_template.docx
```

### 2-2) (권장) case.xlsx 폼 생성 → DOCX 생성
```bash
# 빈 입력 폼 생성
eia-gen xlsx-template --out output/case_template.xlsx

# (작성 후) xlsx → docx
eia-gen generate-xlsx \
  --xlsx output/case_template.xlsx \
  --sources examples/sources.sample.yaml \
  --out output/report_from_xlsx.docx \
  --no-use-llm
```

### 2-2a) (추가) v2(case.xlsx) 템플릿(통합 snake_case) 생성
v2는 `LOOKUPS` 시트를 포함하며, 로더가 자동으로 v2로 감지해 처리합니다.
```bash
eia-gen xlsx-template-v2 --out output/case_template.v2.xlsx
```

### 2-2b) (권장) 신규 케이스 스타터킷 생성
```bash
eia-gen init-case --out-dir output/case_new --project-id PRJ-2026-0001 --project-name 'OO 관광농원 조성사업'
```
기본값으로 `case.xlsx`의 `DATA_REQUESTS`가 플래너로 채워집니다(키/좌표가 없으면 disabled). 필요 없으면 `--no-plan-data-requests`를 사용하세요.

(선택) 받은 첨부를 `attachments/inbox/`에 넣은 뒤 정규화+엑셀등록:
```bash
eia-gen ingest-attachments --xlsx output/case_new/case.xlsx --src-id S-CLIENT-001
```


### 2-3) (신규) 하나의 case.xlsx로 EIA + DIA 병렬 생성
`case.xlsx`에 DIA 시트(`DIA_SCOPE`, `DIA_MAINTENANCE` 등)를 채우면, 한 번에 두 보고서를 생성할 수 있습니다.
```bash
eia-gen generate-xlsx-both \
  --xlsx output/case_template.xlsx \
  --sources examples/sources.sample.yaml \
  --out-dir output \
  --no-use-llm
```
(옵션) 생성 전에 `DATA_REQUESTS`까지 자동 실행하려면 `--enrich`(필요시 `--enrich-overwrite-plan`)를 추가합니다.

산출물:
- `output/report_eia.docx`
- `output/report_dia.docx`
- `output/validation_report_eia.json`
- `output/validation_report_dia.json`
- `output/source_register.xlsx`

### 2-3a) DOCX 생성 없이 QA만 빠르게 점검
```bash
eia-gen check-xlsx-both --xlsx output/case_template.v2.xlsx --sources examples/sources.sample.yaml --out-dir output/_check
```

### 2-3b) (옵션) DATA_REQUESTS로 WMS/AUTO_GIS 보강(v0)
`DATA_REQUESTS`는 “빈칸 탐지 → 수집/계산 → evidence 저장 → 시트 반영”을 위한 레지스트리입니다.

```bash
eia-gen plan-data-requests --xlsx output/case_template.v2.xlsx
eia-gen run-data-requests --xlsx output/case_template.v2.xlsx
```

또는 생성 커맨드에서 한 번에 실행할 수 있습니다:
- `eia-gen generate-xlsx --enrich ...`
- `eia-gen generate-xlsx-both --enrich ...`

현재(v0) 지원:
- `GEOCODE`: 주소(`LOCATION.address_*`) → 좌표(`LOCATION.center_lat/center_lon`) best-effort 채움 + `attachments/evidence/api/*_geocode.json` 저장
  - `VWORLD_API_KEY`가 있으면 VWorld를 우선 사용(권장), 없으면 Nominatim(OpenStreetMap)으로 best-effort
  - 목적: “좌표가 없어서 WMS/API 수집이 안 됨” 진입장벽 제거(1-pass enrich 체인)
- `WMS`: 레이어 이미지를 `attachments/evidence/wms/*.png`로 저장 + `ATTACHMENTS` 등록
- `PDF_PAGE`: 스캔본/PDF에서 특정 페이지를 PNG evidence로 고정(`attachments/evidence/pdf/*.png`) + `ATTACHMENTS` 등록
  - OCR 인덱스 기반 대량 계획: `eia-gen plan-pdf-evidence --xlsx ... --index output/pdf_index/... --src-id S-...`
- `AIRKOREA`: 대기현황(`ENV_BASE_AIR`) 자동 채움 + `attachments/evidence/api/*_airkorea.json` 저장
- `KMA_ASOS`: 강우 요약(`DRR_HYDRO_RAIN`) best-effort 채움 + `attachments/evidence/api/*_kma_asos.json` 저장
  - 관측소 ID 자동 선택을 위해 `eia-gen fetch-kma-asos-stations` 사용(또는 플래너가 키가 있으면 자동 갱신)
- `AUTO_GIS`: 일부 시트 자동 채움(예: `PARCELS.zoning` 기반 `ZONING_BREAKDOWN` 집계) + `attachments/evidence/calc/*.csv` 저장
- `AUTO_GIS`: (추가) WMS evidence(투명 PNG) 기반 `ZONING_OVERLAY` O/X·거리 best-effort 산정 + `attachments/evidence/gis/*.csv` 저장
### 2-3c) (옵션) Reference Pack(샘플 재사용)
같은 권역에서 샘플 기반 파트를 재사용하려면, QA-clean 케이스를 pack으로 내보내고 새 케이스에 적용합니다.

```bash
# pack 내보내기
eia-gen export-reference-pack --case-dir output/case_changwon_2025 --pack-id CHANGWON_JINJEON_APPROVED_2025 --out-dir reference_packs

# 새 케이스에 적용(빈 시트만 채움)
eia-gen apply-reference-pack --xlsx output/case_new/case.xlsx --sources output/case_new/sources.yaml --pack-dir reference_packs/CHANGWON_JINJEON_APPROVED_2025
```

### 2-3d) (권장) SSOT 300p 본안 + 페이지 치환(SSOT_PAGE_OVERRIDES)

“샘플 PDF(≈300p)와 동일한 분량/서식”을 최우선으로 잠그려면,
샘플 PDF 페이지를 SSOT로 삽입한 뒤 **특정 페이지만 케이스 도면(PDF 페이지)로 교체**하는 방식이 가장 빠릅니다.

- 입력: `case.xlsx(v2)`의 `SSOT_PAGE_OVERRIDES` 시트
  - `sample_page`(샘플 PDF 1-based) → `override_file_path`(PDF) + `override_page`(1-based)
- 효과: “샘플의 자리/캡션/분량”은 유지하면서 “케이스별 도면”만 반영

도면 PDF의 페이지 번호 찾기(OCR, 스캔본 대응):
```bash
./.venv/bin/python scripts/ocr_pdf_page_titles.py \
  --pdf output/case_new/attachments/normalized/ATT-0001__00토목도면-자동차야영장.pdf \
  --page-start 1 --page-end 50 \
  --keywords '계획평면도,배수계획평면도,단면도,토공' \
  --max-print 200
```

자세한 운영 가이드:
- `docs/20_user_manual_ko.md`
- `docs/21_input_contract_case_xlsx_v2.md`



## sources.yaml 포맷
- v1(리스트): `examples/sources.sample.yaml`
- v2(권장, version/project/sources 구조): `examples/sources.v2.sample.yaml`

### 3) API 서버 실행(선택)
```bash
eia-gen serve --host 0.0.0.0 --port 8000
```

### 4) API 호출(선택)
`POST /v1/reports/small-eia:generate`는 `report.docx + validation_report.json + source_register.xlsx + draft.json`을 ZIP으로 반환합니다.
```bash
# (샘플 케이스는 examples/assets/* 경로를 참조하므로, 동일 경로로 zip을 만듭니다)
mkdir -p output
zip -r output/assets.zip examples/assets

curl -X POST 'http://127.0.0.1:8000/v1/reports/small-eia:generate?use_llm=false' \
  -F case_file=@examples/case.sample.yaml \
  -F sources_file=@examples/sources.sample.yaml \
  -F assets_zip=@output/assets.zip \
  -o output/small_eia_bundle.zip
```
템플릿 앵커 기반 생성이 필요하면 `use_template_map=true`를 추가합니다:
`/v1/reports/small-eia:generate?use_llm=false&use_template_map=true`

`case_file`은 `case.yaml`뿐 아니라 `case.xlsx`도 업로드할 수 있습니다(파일 확장자/헤더로 자동 감지).

## LLM 설정(선택)
환경변수:
- `OPENAI_API_KEY`: 설정 시 본문 서술을 LLM으로 생성
- 미설정 시: 규칙기반/자리표시자로 문서를 생성

## DOCX 출력(서식 일치성)
기본값은 “샘플 PDF와 최대한 서식 일치”를 위해, 본문/캡션의 인라인 출처표기(`〔SRC:...〕`)를 **DOCX에서 숨김** 처리합니다(출처 추적은 `source_register.xlsx`로 유지).

환경변수:
- `EIA_GEN_DOCX_RENDER_CITATIONS=true`: DOCX에 인라인 출처표기를 표시(기본값=false)
- `EIA_GEN_DOCX_STRICT_TEMPLATE=true`: 템플릿 앵커/표 구조 누락 시 즉시 실패(서식 드리프트 방지)

## 운영 문서
- `docs/20_user_manual_ko.md`: 실사용 가이드(템플릿 운영/그림/PDF/QA/스모크 체크)
- `docs/21_input_contract_case_xlsx_v2.md`: v2 case.xlsx 입력 계약(현재 구현 기준)

## SSOT(spec)
- `spec/sections.yaml`: 섹션/조건부 출력/금칙어
- `spec/table_specs.yaml`: 표 정의 + 검증(합계/필수값 등)
- `spec/figure_specs.yaml`: 그림 정의(필수/조건부)
- `spec/template_map.yaml`: DOCX 앵커 ↔ 삽입 매핑
- `spec_dia/*`: 소규모 재해영향평가(DIA) 스펙/앵커/표/그림 정의

## 문서
- `docs/00_v1_design_freeze.md`
- `docs/01_v1_acceptance_checklist.md`
- `docs/02_deep_research.md`
- `docs/03_case_xlsx_spec.md`
- `docs/04_v2_implementation_spec.md`
- `docs/05_v2_complete_spec.md`

## VSCode 작업 플로우(권장)
Word(.docx)는 **템플릿(앵커 포함)**이고, 내용(텍스트/표/그림)은 엔진이 채웁니다. 보통 Word를 직접 타이핑하지 않습니다.

### 1) 입력 준비(프로젝트별 1회)
- `case.xlsx` 작성: 공통 → EIA → DIA 순서로 채움
- `sources.yaml` 작성: `SRC-...` 출처 ID를 먼저 만들고, `case.xlsx`의 `src_id/src_ids`에 연결
- `attachments/` 배치: 사진/도면/PDF/GeoJSON 등 원본 파일을 넣고, `case.xlsx`에서 파일 경로/ID로 참조

### 2) 생성(반복)
- EIA만: `eia-gen generate-xlsx --xlsx case.xlsx --sources sources.yaml --out out/report_eia.docx --use-template-map --no-use-llm`
- EIA+DIA 병렬: `eia-gen generate-xlsx-both --xlsx case.xlsx --sources sources.yaml --out-dir out --use-template-map --no-use-llm`

### 3) 확인/보완(반복)
- 결과물: `report_*.docx`, `validation_report*.json`, `source_register.xlsx`
- 누락값은 문서에 `【작성자 기입 필요】`/`【첨부 필요】`로 남기고, `validation_report*.json`에 TODO/WARN으로 정리됩니다.
