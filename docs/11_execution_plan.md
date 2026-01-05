# 2025-12 실행 플랜 — 창원 샘플 기반 Input 세팅 + 도면/이미지 파이프라인(체크리스트)

> 목적: “샘플(창원, 2025) 수준의 문서 구성/도면·이미지 품질”을 **새 프로젝트(지번/좌표 + 첨부 파일)**에도 재현 가능하게 만든다.  
> 원칙: **설계도서급 도면을 AI로 ‘그럴듯하게 생성’하는 기능은 금지**하고, (1) 제공된 공식 도면/사진은 **정리·배치**하며 (2) 공식 도면이 없을 때만 **참고도(개략도)** 를 생성하고 워터마크/캡션에 강제 표기한다.  
> (정책 근거) refs: `eia-gen/docs/04_v2_implementation_spec.md:32`, `eia-gen/docs/04_v2_implementation_spec.md:55`

> 운영 원칙(필수): 본 실행 플랜의 모든 신규 설계/구현 결정은 First Principles Thinking(FPT)를 기본 전제로 한다.  
> - SSOT: `eia-gen/docs/25_first_principles_thinking_engine.md`  
> - 비사소한 변경은 이슈/PR/코멘트에 `FPT:` 블록(문제 재정의→가정→분해→재조립→5단계→비용→리스크→10배 옵션→권장안)을 남긴다.

---

## 0) SSOT(단일 진실) / 이번 플랜의 “참조 우선순위”

- [x] 스펙/템플릿/QA는 `eia-gen/` 아래가 SSOT이며, 루트의 `02_~05_*.md`는 중복본(레거시)로 취급한다.  
  refs: `AGENTS.md:11`, `eia-gen/README.md:214`
- [x] 샘플(창원, 2025) 기반 템플릿화 작업 지시/우선순위 SSOT: `eia-gen/docs/sample_changwon_gingerfarm_2025_prompt_pack.md`  
  refs: `eia-gen/docs/sample_changwon_gingerfarm_2025_prompt_pack.md:1`
- [x] 도면/이미지(샘플 수준) 파이프라인(레시피/규격) SSOT: `eia-gen/docs/09_figure_generation_pipeline.md`  
  refs: `eia-gen/docs/09_figure_generation_pipeline.md:1`
- [x] 공통 로드맵(장기) SSOT: `eia-gen/docs/06_eia_dia_shared_core_plan.md`  
  refs: `eia-gen/docs/06_eia_dia_shared_core_plan.md:1`
- [x] 지도/오버레이(WMS/WMTS/REST) 설정 SSOT: `eia-gen/config/wms_layers.yaml`, `eia-gen/config/basemap.yaml`, `eia-gen/config/cache.yaml`  
  refs: `eia-gen/config/wms_layers.yaml:1`, `eia-gen/config/basemap.yaml:1`, `eia-gen/config/cache.yaml:1`
- [x] 키/환경변수 템플릿: `eia-gen/.env.example`  
  refs: `eia-gen/.env.example:1`
- [x] 출처/증빙/사용처(Traceability) “잠금” 스펙: `eia-gen/docs/12_traceability_spec.md`  
  refs: `eia-gen/docs/12_traceability_spec.md:1`
- [x] DATA_REQUESTS(자동 수집 플래너) + 커넥터/실행기(Enrich) 스펙 SSOT: `eia-gen/docs/13_data_requests_and_connectors_spec.md`  
  - [x] 연계 설정 SSOT: `eia-gen/config/kosis_datasets.yaml`, `eia-gen/config/data_acquisition_rules.yaml`, `eia-gen/config/stations/README.md`  
  refs: `eia-gen/docs/13_data_requests_and_connectors_spec.md:1`, `eia-gen/config/kosis_datasets.yaml:1`, `eia-gen/config/data_acquisition_rules.yaml:1`, `eia-gen/config/stations/README.md:1`
- [x] (2025.05 실무지침 반영) “부록/별지서식/도면(A3)” 기반 플랜 델타: `eia-gen/docs/26_drr_guideline_2025_05_plan_delta.md`  
  refs: `eia-gen/docs/26_drr_guideline_2025_05_plan_delta.md:1`

---

## 1) 현재 구현 현황(레포 기준) — “이미 있음 / 부족 / 없음”

### 1.1 이미 구현됨(계획에서 “중복”으로 간주)

- [x] EIA + DIA 병렬 생성 CLI(`generate-xlsx-both`) + EIA/DIA 각각 spec 로드 + 산출물 분리 저장  
  refs: `eia-gen/src/eia_gen/cli.py:393`
- [x] v2 `case.xlsx` 템플릿에 `ZONING_OVERLAY`, `FIGURES`, `ATTACHMENTS`, `FIELD_SURVEY_LOG` 포함(입력 폼 구조 확보)  
  refs: `eia-gen/src/eia_gen/services/xlsx/case_template_v2.py:18`, `eia-gen/src/eia_gen/services/xlsx/case_template_v2.py:125`, `eia-gen/src/eia_gen/services/xlsx/case_template_v2.py:204`
- [x] v2 로더가 핵심 시트를 읽어 `Case`/`model_extra`로 매핑(기초 파이프라인)  
  refs: `eia-gen/src/eia_gen/services/xlsx/case_reader_v2.py:418`, `eia-gen/src/eia_gen/services/xlsx/case_reader_v2.py:448`, `eia-gen/src/eia_gen/services/xlsx/case_reader_v2.py:473`
- [x] `source_register.xlsx`에 `Evidence Register` / `Claims` 시트 생성(“문장/표/그림” 단위 traceability 뼈대)  
  refs: `eia-gen/src/eia_gen/services/export/source_register_xlsx.py:266`, `eia-gen/src/eia_gen/services/export/source_register_xlsx.py:350`
- [x] (traceability 정합성) 템플릿에 없는 섹션(예: `SSOT_*_REUSE_PDF`)이 `source_register.xlsx`의 `Claims/USAGE_REGISTER`에 포함되어 “샘플 혼입”처럼 보이던 노이즈를 차단 (done-by: Codex)  
  refs: `eia-gen/src/eia_gen/cli.py:84`, `eia-gen/src/eia_gen/cli.py:283`, `eia-gen/src/eia_gen/services/export/source_register_xlsx.py:152`, `eia-gen/src/eia_gen/services/export/source_register_xlsx.py:516`

  FPT:
  1) 문제 재정의: 실제 DOCX에 렌더되지 않은 섹션의 인용/출처가 `source_register.xlsx`에 “사용됨”으로 기록되면, 샘플(타 지역) 소스가 혼입된 것처럼 보이고 검증/감사 리스크가 커진다.
  2) 가정(삭제 가능/검증 필요/필수 제약): (a) 소스 레지스터는 “실제 산출물에 반영된 사용처” 기준이어야 한다. (b) 템플릿 기반 렌더링에서는 ‘템플릿에 존재하는 앵커/블록’이 렌더 가능 섹션의 상한이다.
  3) 제1원칙 분해(Primitives): (a) draft 섹션 집합 (b) 템플릿 앵커 존재 여부 (c) 출처 ID 집계 (d) Claims/Usage sheet 생성.
  4) 재조립(새 시스템 설계): 템플릿(docx)에서 렌더 가능한 섹션 allowlist를 계산하고, Source Register/Claims/USAGE_REGISTER의 집계 대상 섹션을 allowlist로 필터링한다.
  5) 최종 권장안(+지표/검증 루프): (a) `S-CHANGWON-SAMPLE/진전/진저`가 `source_register.xlsx` 전 시트에서 0건, (b) 실제 보고서 텍스트/그림은 변하지 않음을 게이트로 확인한다.
- [x] 샘플 PDF(스캔)용 OCR 인덱싱(2-pass) + 요약 생성 스크립트 존재  
  refs: `eia-gen/scripts/extract_pdf_index_twopass.py:59`, `eia-gen/scripts/summarize_pdf_index.py:11`
- [x] 샘플(창원, 2025) 서식에 맞춘 Word 템플릿(앵커 포함) 생성 스크립트 존재  
  refs: `eia-gen/scripts/make_template_sample_changwon_2025.py:45`
- [x] 템플릿 린트/스캐폴딩 CLI(`template-check`, `template-scaffold`) + JSON 리포트(템플릿 품질 게이트)  
  refs: `eia-gen/src/eia_gen/cli.py:590`, `eia-gen/src/eia_gen/services/docx/template_tools.py:61`
- [x] `DATA_REQUESTS` 플래너/러너가 WMS evidence를 `attachments/evidence/*`로 저장 + `ATTACHMENTS` 등록(기본 파이프라인)  
  refs: `eia-gen/src/eia_gen/services/data_requests/planner.py:117`, `eia-gen/src/eia_gen/services/data_requests/runner.py:280`
- [x] 키/엔드포인트 최소 검증 CLI(`verify-keys`) 존재  
  refs: `eia-gen/src/eia_gen/cli.py:1302`
- [x] 첨부 수집/정규화(파일명/ID) + sha256 manifest(`attachments/attachments_manifest.json`) 존재  
  refs: `eia-gen/src/eia_gen/cli.py:901`, `eia-gen/src/eia_gen/services/ingest_attachments.py:98`
- [x] reference pack export/apply + init-case 존재(샘플팩/재사용팩 운영 기반)  
  refs: `eia-gen/src/eia_gen/cli.py:664`, `eia-gen/src/eia_gen/cli.py:940`, `eia-gen/src/eia_gen/services/reference_packs.py:123`
- [x] SSOT(창원 샘플) 페이지를 `[[PDF_PAGE:...]]`로 삽입 + `SSOT_PAGE_OVERRIDES`로 “샘플 페이지 → 케이스 PDF 페이지” 치환 가능  
  refs: `eia-gen/src/eia_gen/services/xlsx/case_template_v2.py:158`, `eia-gen/src/eia_gen/services/writer.py:556`, `eia-gen/src/eia_gen/services/docx/spec_renderer.py:137`

### 1.2 부분 구현(갭) — “요구 품질 대비 부족”

- [x] 지도/도면 자동생성(샘플 스타일 지도 1장 완성 레시피: 북표/축척/범례/한글 폰트 고정) 구현 (done-by: Codex)  
  refs: `eia-gen/src/eia_gen/services/figures/map_generate.py:1`, `eia-gen/docs/09_figure_generation_pipeline.md:41`, `eia-gen/tests/test_map_generate.py:1`
- [x] DOCX figure materialize 산출물의 “증빙화(ATTACHMENTS/source_register 연결)”을 보강(저장: `attachments/derived/figures/_materialized/*`) (done-by: Codex)  
  refs: `eia-gen/src/eia_gen/services/docx/spec_renderer.py:211`, `eia-gen/src/eia_gen/services/figures/derived_evidence.py:42`, `eia-gen/src/eia_gen/services/export/source_register_xlsx.py:266`
- [x] 샘플 반복 “환경기준 표(대기/생활환경/소음 등)”를 표 스펙/앵커(템플릿) SSOT로 고정(정적 표 3종) (done-by: Codex)  
  refs: `eia-gen/spec/table_specs.yaml:1`, `eia-gen/spec/template_map.yaml:1`, `eia-gen/templates/report_template.sample_changwon_2025.scaffolded.docx:1`, `eia-gen/docs/32_env_standards_tables_plan.md:1`

- [x] (제출본/1인 사용) `case.xlsx(v2)` 업그레이드(누락 시트 자동 추가 + 기존 값 보존) CLI 추가 (done-by: Codex)  
  - 예: `ENV_BASE_SOCIO` 누락 시 QA/WARN 이전에 시트부터 보정(버전 갭 제거)  
  refs: `eia-gen/src/eia_gen/cli.py:255`, `eia-gen/src/eia_gen/services/xlsx/upgrade_v2.py:1`, `eia-gen/src/eia_gen/services/xlsx/case_template_v2.py:223`
- [x] (제출본/1인 사용) “submission mode” 게이트: 핵심 시트 0행/누락을 WARN→ERROR로 격상(스코핑/대책/이행관리/기초현황) (done-by: Codex)  
  - `config/data_acquisition_rules_submission*.yaml`로 기본 규칙 대비 ERROR를 격상  
  - CLI `--submission`으로 선택 적용(기본값은 비활성)  
  refs: `eia-gen/config/data_acquisition_rules_submission.yaml:1`, `eia-gen/config/data_acquisition_rules_submission_dia.yaml:1`, `eia-gen/src/eia_gen/services/qa/run.py:355`, `eia-gen/src/eia_gen/cli.py:558`, `eia-gen/src/eia_gen/cli.py:1395`
- [x] (데이터 확보) `verify-keys` 커버리지 확대: EcoZmp WMS(생태자연도 WMS조회)까지 네트워크 검증 포함 + 실패 유형(403/HTML alert) 분류/가이드 (done-by: Codex)  
  refs: `eia-gen/src/eia_gen/cli.py:1691`, `eia-gen/src/eia_gen/services/data_requests/wms.py:172`, `eia-gen/config/wms_layers.yaml:1`
- [x] (데이터 확보) SAFEMAP WMS 키/승인 이슈 해결(또는 공식 발급본 첨부로 fallback 고정) (done-by: Codex)  
  - WMS 실패 시에도 pipeline이 “조용히 증빙을 누락”하지 않도록, placeholder evidence(PNG)를 생성해 `ATTACHMENTS`에 등록(재실행 가능)  
  - DATA_REQUESTS(WMS)에서 `params_json.fallback_file_path`를 지정하면, WMS 불가(승인/키/장애) 시 로컬 공식 이미지(스크린샷 등)로 대체 저장 가능  
  refs: `eia-gen/src/eia_gen/services/data_requests/runner.py:417`, `eia-gen/src/eia_gen/services/data_requests/planner.py:78`, `eia-gen/config/wms_layers.yaml:1`, `eia-gen/docs/34_installation_and_fallbacks.md:1`, `eia-gen/tests/test_data_requests.py:193`

  FPT:
  1) 문제 재정의: SAFEMAP WMS는 키/승인/서비스 상태에 따라 실패가 빈번한데, 실패 시 증빙이 ‘생성되지 않은 채’로 남으면 사용자가 누락을 늦게 발견하고 제출/감사 리스크가 커진다.
  2) 가정(삭제 가능/검증 필요/필수 제약): (1) 외부 WMS 장애는 코드로 해결할 수 없으므로, 실패를 “명시적 산출물(placeholder) + next action”으로 바꿔야 한다. (2) 공식 발급본/공식 화면 캡처가 있으면 첨부로 대체 가능해야 한다. (3) 허위기재 방지를 위해 ‘그럴듯한 지도 생성’은 하지 않는다.
  3) 제1원칙 분해(Primitives): (a) WMS 요청 (b) 실패 분류/메시지 (c) evidence 파일 존재 (d) ATTACHMENTS/traceability 연결 (e) rerun 가능성.
  4) 재조립(새 시스템 설계): WMS 실패 시 placeholder PNG를 만들고 ATTACHMENTS에 등록하며, 사용자가 `fallback_file_path`로 공식 이미지를 지정하면 동일 경로로 대체 저장한다.
  5) 최종 권장안(+지표/검증 루프): 게이트에서 WMS 실패가 발생해도 (a) 증빙이 누락되지 않고 (b) placeholder/fallback이 명시적으로 남으며 (c) key/승인 해결 후 재실행으로 정상 이미지로 치환되는지로 검증한다.

- [x] (데이터 확보) data.go.kr 생태자연도(EcoZmp) API 활용신청/승인(403 해결) 또는 기본 disabled 유지(명시적 opt-in) (done-by: Codex)  
  - 기본 DATA_REQUESTS 플랜에서 ECO_NATURE는 opt-in(disabled)으로 시작(승인/권한 확인 후 enabled=TRUE로 전환)  
  refs: `eia-gen/config/wms_layers.yaml:87`, `eia-gen/src/eia_gen/services/data_requests/planner.py:73`, `eia-gen/src/eia_gen/cli.py:2161`
- [x] (데이터 확보) KOSIS `ENV_BASE_SOCIO` 프리셋 확정 + 자동 수집 경로 고정(Param/statisticsParameterData) + multi-`itmId` fan-out 지원 (done-by: Codex)  
  refs: `eia-gen/config/kosis_datasets.yaml:1`, `eia-gen/src/eia_gen/services/data_requests/kosis.py:1`, `eia-gen/src/eia_gen/cli.py:1784`, `eia-gen/docs/20_user_manual_ko.md:600`

FPT:
1) 문제 재정의: KOSIS 자동수집이 “키는 있는데도” err=20/21로 실패하거나(엔드포인트/파라미터 불일치), `TBD` 설정 때문에 아예 실행이 막혀 `ENV_BASE_SOCIO` 입력이 반복적으로 수기로 남는다.
2) 가정(삭제 가능/검증 필요/필수 제약): (a) SSOT(config)로 orgId/tblId/itmId를 잠가야 재사용/감사가 가능하다, (b) 케이스는 admin_code만 바뀌게 해야 한다, (c) endpoint는 KOSIS 공식 가이드에 맞춰야 한다.
3) 제1원칙 분해(Primitives): (a) 고정 프리셋(테이블/항목코드) (b) 변하는 값(행정구역 코드) (c) 호출 엔드포인트 (d) 증빙(evidence) 기록.
4) 재조립(새 시스템 설계): DT_1IN1502 프리셋을 SSOT로 고정하고, `openapi/Param/statisticsParameterData.do`로 호출을 통일하며, multi-`itmId`는 client-side fan-out으로 흡수한다.
5) 최종 권장안(+지표/검증 루프): `verify-keys`에서 KOSIS 1-cell 스모크를 통과하고, `ENV_BASE_SOCIO`가 5개 연도 행으로 채워지는지로 검증한다.

FPT:
1) 문제 재정의: case.xlsx(v2) 스키마/시트 드리프트와 “WARN으로만 남는 핵심 누락” 때문에, 혼자 제출본을 만들 때 누락 탐지/수정 루프가 길어진다.
2) 가정(삭제 가능/검증 필요/필수 제약): (a) 누락은 생성으로 메우지 않고 입력/근거로 채운다, (b) 제출본은 핵심 누락을 ERROR로 차단해야 한다, (c) 허위기재 방지가 최우선이다.
3) 제1원칙 분해(Primitives): (a) 입력 계약(case.xlsx 스키마), (b) 누락 탐지(QA 규칙), (c) 반복 비용(업그레이드/게이트 자동화), (d) 증빙/출처 연결.
4) 재조립(새 시스템 설계): (1) `xlsx-upgrade-v2`로 시트/컬럼 드리프트를 제거하고, (2) `--submission`으로 핵심 누락을 ERROR로 격상(보고서 종류별 규칙 분리)한다.
5) 머스크 5단계 실행안(순서 고정):
   5-1) 요구사항 의심: “WARN이면 나중에 해도 된다”가 제출본에 성립하는가?
   5-2) 삭제: 버전 드리프트로 생기는 “시트가 없어서 채울 수 없음”을 제거
   5-3) 단순화/최적화: 제출 필수 항목만 우선 ERROR로 고정
   5-4) 가속: `next_actions`로 어디를 채울지 즉시 안내
   5-5) 자동화: 누락 시트 자동 추가 + 게이트로 회귀 방지
6) 비용/Idiot Index(낭비/마찰): “어디가 비었는지 찾기/시트 만들기” 반복을 0에 수렴시킨다.
7) 리스크(인간/규제/2차효과/불확정): 자동 서술로 빈칸을 숨기면 허위기재 리스크가 증가한다(금지).
8) 10배 옵션(2~4개): 제출 모드에서 필수 컬럼 단위(빈칸)까지 ERROR로 확장 + 케이스별 체크리스트 자동 생성.
9) 최종 권장안(+지표/검증 루프): `xlsx-upgrade-v2` + `--submission` + next_actions로 “누락→수정→재검증” 루프를 짧게 고정한다.

FPT:
1) 문제 재정의: DOCX 렌더 중 materialize된 PNG가 만들어져도 Evidence/Claims로 연결되지 않으면 “무슨 파일이 어디에 쓰였는지”가 역추적되지 않는다.
2) 가정(삭제 가능/검증 필요/필수 제약): 렌더 과정은 입력(case.xlsx/attachments)과 산출물(output/attachments/derived)만으로 재현 가능해야 한다.
3) 제1원칙 분해(Primitives): (a) materialize는 결정적(out_dir+옵션+입력)이고 (b) 증빙은 manifest에 기록되어 (c) source_register.xlsx로 export되어야 한다.
4) 재조립(새 시스템 설계): 템플릿 섹션의 `[[PDF_PAGE:...]]` 경로까지 포함해 materialize 산출물을 `derived_evidence_manifest`로 기록하고, export 시 Evidence Register/Claims에 반영한다.
5) 머스크 5단계 실행안(순서 고정):
   5-1) 요구사항 의심: “파일이 생성됐다”는 사실만으로는 제출/감사 요구를 만족하지 못한다(사용처/출처 연결이 필요).
   5-2) 삭제: 입력 XLSX를 렌더 과정에서 직접 수정(ATTACHMENTS 자동 업데이트)하는 부작용은 피한다.
   5-3) 단순화/최적화: 렌더 단계에서 in-memory manifest만 기록하고, 기존 export 경로에서 source_register로 모은다.
   5-4) 가속: 테스트로 `[[PDF_PAGE:...]]` 경로의 manifest 기록을 잠가 회귀를 방지한다.
   5-5) 자동화: Gate/생성 커맨드가 항상 `source_register.xlsx`를 함께 생성해 evidence 누락을 조기에 드러낸다.
6) 비용/Idiot Index(낭비/마찰): “어느 페이지/이미지가 어디에 들어갔는지 확인”은 사람이 하면 가장 비싸므로 자동 기록으로 비용을 0에 수렴시킨다.
7) 리스크(인간/규제/2차효과/불확정): 출처/증빙 누락은 규제 리스크로 직결된다(미삽입/오삽입 사고).
8) 10배 옵션(2~4개): 다음 단계로 sha256/recipe/input_manifest까지 확장해 완전한 traceability(재현·감사)를 달성한다.
9) 최종 권장안(+지표/검증 루프): materialize 경로에서 derived evidence 기록→export 연결을 기본값으로 두고, 테스트로 지속 검증한다.

### 1.3 미구현(필요) — “이번 요청 핵심(샘플 수준 도면/이미지)”

- [x] `CALLOUT_COMPOSITE`(사진대지/콜아웃 판넬) 생성(2~6장, 고정 그리드, 번호/캡션바) (done-by: Codex)  
  refs: `eia-gen/docs/09_figure_generation_pipeline.md:61`, `eia-gen/src/eia_gen/services/figures/photo_sheet_generate.py:77`, `eia-gen/src/eia_gen/services/figures/callout_composite.py:151`, `eia-gen/tests/test_photo_sheet_generate.py:1`
  - [x] (B) 스크립트 프로토타입으로 2/4/6 그리드 사진대지 합성 + runbook/예제 산출물 추가(코어 연결 이전 단계 참고/디버그용) (done-by: B)  
    refs: `eia-gen/scripts/compose_callout_composite.py:1`, `eia-gen/scripts/make_callout_composite_from_case.py:1`, `eia-gen/docs/44_callout_composite_runbook.md:1`, `eia-gen/output/annotation_examples/att0008_p1_4_callout_composite.png`
- [x] “제출용 Asset Normalization” 단계(원본→정규화PNG→판넬PNG→DOCX)로 고정(패딩/테두리/워터마크/표준 DPI) (done-by: Codex)  
  refs: `eia-gen/config/asset_normalization.yaml:1`, `eia-gen/src/eia_gen/services/figures/materialize.py:650`, `eia-gen/tests/test_materialize_policy.py:1`
- [x] FIGURES 스키마에 `asset_role/authenticity/source_class/usage_scope/fallback_mode` 등 “정책 강제 컬럼” 도입(= 사람 주의가 아니라 데이터로 강제) (done-by: Codex)  
  refs: `eia-gen/docs/26_drr_guideline_2025_05_plan_delta.md:77`, `eia-gen/src/eia_gen/services/xlsx/case_template_v2.py:135`, `eia-gen/src/eia_gen/services/xlsx/case_reader_v2.py:517`, `eia-gen/src/eia_gen/services/figures/spec_figures.py:18`, `eia-gen/src/eia_gen/services/qa/run.py:305`, `eia-gen/tests/test_figures_authenticity.py:1`
  FPT:
  1) 문제 재정의: FIGURES의 “공식/참고도/사용범위/실패 시 처리”가 사람 기억에 의존하면 워터마크 누락·오인용(허위기재) 리스크가 커진다.
  2) 가정(삭제 가능/검증 필요/필수 제약): 입력 XLSX에 정책 메타데이터를 담고, 렌더/QA는 그 메타데이터로만 판단해야 한다(가드레일: 그럴듯한 도면 생성 금지).
  3) 제1원칙 분해(Primitives): (a) 정책 컬럼 (b) 템플릿 드롭다운 (c) 리더 파싱 (d) 렌더 힌트(gen_method) (e) QA 검증.
  4) 재조립(새 시스템 설계): FIGURES에 정책 컬럼 추가 → 리더가 읽어 asset에 보관 → resolve_figure가 gen_method에 AUTHENTICITY 힌트로 접합 → QA가 REFERENCE 가드레일을 컬럼 기반으로 적용.
  5) 머스크 5단계 실행안(순서 고정):
     5-1) 요구사항 의심: “주의해서 입력”은 실패한다 → 데이터로 강제해야 한다.
     5-2) 삭제: 레거시 `source_origin`만으로의 묵시적 정책을 제거하지 않고(호환), 새로운 컬럼으로 대체 경로를 추가한다.
     5-3) 단순화/최적화: AUTHENTICITY는 gen_method 힌트로 합쳐 downstream(materialize/워터마크)가 그대로 재사용한다.
     5-4) 가속: 단위테스트로 authenticity=REFERENCE → 캡션 prefix + 힌트가 항상 생성됨을 잠근다.
     5-5) 자동화: 템플릿 드롭다운으로 값 범위를 제한해 입력 오류를 감소시킨다.
  6) 비용/Idiot Index(낭비/마찰): 워터마크/캡션 누락 확인을 사람이 반복하지 않도록 비용을 0에 수렴시킨다.
  7) 리스크(인간/규제/2차효과/불확정): REFERENCE가 공식처럼 제출되면 규제 리스크가 즉시 발생한다 → 데이터/QA로 차단.
  8) 10배 옵션(2~4개): 제출 모드에서 `usage_scope/fallback_mode` 비어있음도 ERROR로 승격 + source_register에 컬럼 export.
  9) 최종 권장안(+지표/검증 루프): “컬럼→힌트→QA” 고정 루프로 누락·오류를 조기에 발견한다.
- [x] table cell 내부 앵커 스캔(선택) — 샘플 템플릿에서 표 내부 삽입 요구가 생기면 필요 (done-by: Codex)  
  refs: `eia-gen/src/eia_gen/services/docx/spec_renderer.py:127`, `eia-gen/tests/test_template_renderer.py:147`

### 1.4 완성도/편의성/효율성 평가(핵심 문제 → 개선 방향)

- [x] (Flow) 입력→산출 파이프라인을 “코드 기준”으로 정리(현재 진척/약점 평가의 기준점으로 고정)  
  - Input: `case.xlsx` + `sources.yaml` + `attachments/*`  
  - Flow: (선택) `ingest-attachments` → (선택) `generate-xlsx-both --enrich` → draft 생성 → DOCX 렌더(앵커 치환 + figure materialize) → QA → `source_register.xlsx`  
  - Output: `output/report_*.docx`, `output/validation_report_*.json`, `output/source_register.xlsx` (파생: `attachments/derived/figures/_materialized/*`)  
  refs: `eia-gen/docs/28_project_status_review.md:1`, `eia-gen/src/eia_gen/cli.py:393`, `eia-gen/src/eia_gen/services/docx/spec_renderer.py:359`, `eia-gen/src/eia_gen/services/docx/spec_renderer.py:448`, `eia-gen/src/eia_gen/services/figures/materialize.py:295`, `eia-gen/src/eia_gen/services/qa/run.py:1`, `eia-gen/src/eia_gen/services/export/source_register_xlsx.py:266`

- [x] (완성도: 핵심 문제) “샘플 수준”에서 빠르게 체감이 나는 3대 공백을 특정하고, 실행 플랜(P1/P4/P6)에 해결 항목으로 반영 (체크리스트 진행율: 99.6% = 245/246, 2026-01-04 기준)  
  - 문제점: (1) Asset Normalization 단계 부재 → 해결(P1-0 완료), (2) 사진대지(CALLOUT_COMPOSITE) 코어 연결은 완료(확장: 사진 메타 시트/다장 정책은 미결), (3) normalize 단계 산출물의 evidence/usage 연결 미흡 → 해결(P1-0 완료)  
  - 원인: materialize가 “DOCX 삽입 시점”에만 수행되어 품질(패딩/테두리/워터마크)·증빙(sha256/recipe)·재사용(케이스 폴더 기준)이 분리됨  
  - 해결 방향: Ingest→Normalize→Compose→Insert로 고정하고, 생성 시점에 evidence/usage를 기록(“나중에 붙이기” 금지)  
  refs: `eia-gen/docs/26_drr_guideline_2025_05_plan_delta.md:57`, `eia-gen/src/eia_gen/services/docx/spec_renderer.py:463`, `eia-gen/src/eia_gen/services/figures/materialize.py:304`, `eia-gen/docs/12_traceability_spec.md:65`

- [x] (편의성: 핵심 문제) 사용자가 “페이지 찾기/치환/미리보기”에서 반복 시행착오가 생기는 병목을 특정하고(P0-4/P1-5/P6), 해결 체크리스트로 반영  
  - 문제점: PDF 페이지 탐색/치환은 가능하지만(SSOT_PAGE_OVERRIDES) “검색→프리뷰→검증” UX가 약함  
  - 해결 방향: 키워드 검색(OCR) + 프리뷰 PNG + `SSOT_PAGE_OVERRIDES` 자동 작성 CLI + QA 링크를 한 번에 제공  
  refs: `eia-gen/src/eia_gen/services/writer.py:556`, `eia-gen/src/eia_gen/services/docx/spec_renderer.py:137`, `eia-gen/scripts/ocr_pdf_page_titles.py:91`

- [x] (효율성: 핵심 문제) 실행시간/산출물 용량을 크게 만드는 비용 상위 항목을 특정하고(래스터라이즈/이미지 포맷/중복 생성), 캐시/포맷 정책을 플랜으로 승격  
  - 문제점: (1) PDF→PNG 반복 래스터라이즈, (2) 사진까지 PNG 고정으로 DOCX가 과대해질 수 있음, (3) `_materialized`의 evidence/usage 연결이 미흡(케이스 `attachments/derived` 고정은 완료)  
  - 해결 방향: sha256+recipe 기반 안정 캐시(케이스 `attachments/derived` 중심) + asset_role 기반 포맷 정책(도면=PNG/사진=JPEG) + 증빙화(evidence_id)  
  refs: `eia-gen/src/eia_gen/services/figures/materialize.py:304`, `eia-gen/src/eia_gen/services/figures/materialize.py:352`, `eia-gen/src/eia_gen/cli.py:664`

---

## 2) 선별/정리(중복·불필요 판단) — 이번 플랜에 “넣지 않는 것”

- [x] “샘플 PDF의 표/본문 수치 자체를 OCR로 정답 데이터화”는 오탈자/허위기재 리스크가 커서 제외(구성/캡션/장절 흐름만 참고)  
  refs: `eia-gen/docs/sample_changwon_gingerfarm_2025_prompt_pack.md:11`, `eia-gen/docs/sample_changwon_gingerfarm_2025_prompt_pack.md:138`
- [x] “유사 사진 기반으로 현장 상태/수치/판단을 추정해 본문 자동 작성”은 금지(분류/배치만 가능)  
  refs: `eia-gen/docs/04_v2_implementation_spec.md:52`, `eia-gen/docs/04_v2_implementation_spec.md:60`
- [x] `ZONING_OVERLAY` 같은 이미 구현된 입력/스펙/앵커는 신규 작업에서 제외(필요 시 품질만 개선)  
  refs: `eia-gen/spec/template_map.yaml:37`, `eia-gen/spec/table_specs.yaml:9`

---

## 3) 실행 플랜(우선순위) — “무엇을, 어떤 순서로, 무엇이 끝이면 끝인가”

> 우선순위 원칙(실무지침 반영): **지도(P2)보다 첨부 정규화+삽입 품질(P1)**을 먼저 잠그고, traceability(P4)는 “나중”이 아니라 **생성 시점에 동시**로 붙인다.  
> refs: `eia-gen/docs/26_drr_guideline_2025_05_plan_delta.md:48`

### 3.0 실질 진척(게이트) 스냅샷(권장)

- 체크박스 진척률은 “문서/도구 준비” 지표이며, **제출/재현 품질의 실질 진척은 Gate 통과**로 판정한다.
- Gate-0(doctor): `./.venv/bin/python scripts/doctor_env.py --case-dir output/case_new_max_reuse --strict-case` (레거시/미완성 케이스는 `--doctor-relaxed`로 Gate Runner만 완화)
- Gate-1(template-check):
  - EIA(SSOT full): `./.venv/bin/eia-gen template-check --template templates/report_template.sample_changwon_2025.ssot_full.scaffolded.docx --spec-dir spec --allow-missing-anchors`
  - EIA(normal): `./.venv/bin/eia-gen template-check --template templates/report_template.sample_changwon_2025.scaffolded.docx --spec-dir spec --ignore-missing-anchor-prefix '[[BLOCK:SSOT_'`
  - DIA: `./.venv/bin/eia-gen template-check --template templates/dia_template.scaffolded.docx --spec-dir spec_dia`
  - 참고: SSOT 템플릿은 구조상(샘플 PDF 페이지 삽입) CH* 앵커를 의도적으로 생략하므로 `--allow-missing-anchors`로 운용한다.
- Gate-2A(check-xlsx-both, full spec QA): `./.venv/bin/eia-gen check-xlsx-both --xlsx output/case_new_max_reuse/case.xlsx --sources output/case_new_max_reuse/sources.yaml --out-dir output/case_new_max_reuse/_out_check_current`
  - 의미: “스펙 전체” 기준으로 누락/placeholder/TBD를 드러내는 지표(케이스 입력 완성도).
- Gate-2B(check-xlsx-both, template-scoped QA):
  - SSOT(default): `./.venv/bin/eia-gen check-xlsx-both --scope-to-default-templates --xlsx output/case_new_max_reuse/case.xlsx --sources output/case_new_max_reuse/sources.yaml --out-dir output/case_new_max_reuse/_out_check_scoped_ssot`
  - normal(backlog): `./.venv/bin/eia-gen check-xlsx-both --template-eia templates/report_template.sample_changwon_2025.scaffolded.docx --template-dia templates/dia_template.scaffolded.docx --xlsx output/case_new_max_reuse/case.xlsx --sources output/case_new_max_reuse/sources.yaml --out-dir output/case_new_max_reuse/_out_check_scoped_normal`
  - 의미: “실제 생성 템플릿에 존재하는 섹션/앵커”만 대상으로 QA(최종 산출물 체감 품질 / 입력 backlog).
  - 운영(거짓-그린 방지): `scripts/run_quality_gates.py`는 케이스 폴더에 `templates/report_template_{eia,dia}.docx`가 있으면 이를 우선 사용해 scoped QA/생성을 수행한다(다른 케이스의 SSOT PDF 페이지/증빙이 섞이는 사고 방지).
- Gate-3(verify-keys, 선택):
  - 일상/회귀(결정적): `./.venv/bin/eia-gen verify-keys --mode presence --out output/_tmp_verify_keys_presence.json` (env-only, 네트워크 없음)
  - 제출 전 최종 확인: `./.venv/bin/eia-gen verify-keys --mode network --out output/_tmp_verify_keys_network.json --strict` (외부 장애 영향 가능)
- Gate-4(enrich/DATA_REQUESTS, 선택): `./.venv/bin/python scripts/run_quality_gates.py --mode check --enrich --case-dir output/case_new_max_reuse` (case.xlsx/증빙 in-place 갱신)
- Gate-ALL(권장, one-shot): `./.venv/bin/python scripts/run_quality_gates.py --mode check --write-next-actions --append-data-requests-summary --append-ssot-overrides-summary --case-dir output/case_new_max_reuse` (옵션: `--verify-keys`(기본 presence), `--verify-keys-mode network`, `--enrich`)
- 회귀(베이스라인 N개): `./.venv/bin/python scripts/run_regression_gates.py`

현재 스냅샷(2026-01-02, 회귀 + 생성까지):
- 회귀 커맨드: `./.venv/bin/python scripts/run_regression_gates.py --mode both --verify-keys --verify-keys-strict --cases output/case_new output/case_new_max_reuse output/case_changwon_2025`

- `output/case_new` (quality_gates=`output/case_new/_quality_gates/20260102_021938`):
  - check_full: EIA ERROR 0, WARN 44 (placeholder 30, S-TBD 25) / DIA ERROR 0, WARN 10 (placeholder 5, S-TBD 4)
  - check_scoped_default(SSOT): EIA ERROR 0, WARN 14 (placeholder 13, S-TBD 6) / DIA ERROR 0, WARN 10 (placeholder 5, S-TBD 4)
  - check_scoped_normal: EIA ERROR 0, WARN 42 (placeholder 29, S-TBD 24) / DIA ERROR 0, WARN 10 (placeholder 5, S-TBD 4)

- `output/case_new_max_reuse` (quality_gates=`output/case_new_max_reuse/_quality_gates/20260102_022028`):
  - check_full: EIA ERROR 0, WARN 31 (placeholder 17, S-TBD 17) / DIA ERROR 0, WARN 0
  - check_scoped_default(SSOT): EIA ERROR 0, WARN 0 / DIA ERROR 0, WARN 0
  - check_scoped_normal: EIA ERROR 0, WARN 29 (placeholder 16, S-TBD 16) / DIA ERROR 0, WARN 0

- `output/case_changwon_2025` (quality_gates=`output/case_changwon_2025/_quality_gates/20260102_022136`):
  - check_full: EIA ERROR 0, WARN 36 (placeholder 25, S-TBD 23) / DIA ERROR 0, WARN 4 (placeholder 4, S-TBD 4)
  - check_scoped_default(SSOT): EIA ERROR 0, WARN 6 (placeholder 8, S-TBD 6) / DIA ERROR 0, WARN 4 (placeholder 4, S-TBD 4)
  - check_scoped_normal: EIA ERROR 0, WARN 34 (placeholder 24, S-TBD 22) / DIA ERROR 0, WARN 4 (placeholder 4, S-TBD 4)

FPT:
1) 문제 재정의: 샘플(SSOT full) PDF 페이지 재사용 결과가 신규 케이스 산출물에 그대로 섞이면(지명/수치/도면) 허위기재·감사·보완요구 리스크가 급증한다.
2) 가정(삭제 가능/검증 필요/필수 제약): (1) 케이스 폴더 `templates/`는 “그 케이스에 적용할 템플릿”의 SSOT다. (2) 샘플 PDF는 서식/앵커 검증용이며, 타 케이스 제출본에 내용이 섞이면 안 된다. (3) 부족한 데이터는 TBD로 남기되, 틀린 데이터는 제거한다.
3) 제1원칙 분해(Primitives): 템플릿 선택 우선순위, scoped QA 범위, sources/attachments의 source_id/evidence_id, DOCX에 삽입되는 이미지(SSOT PDF_PAGE).
4) 재조립(새 시스템 설계): Gate Runner는 case-local 템플릿이 있으면 이를 우선 사용하고, 샘플 기반 증빙/출처는 케이스 입력에서 제거해 “진짜 입력 backlog”만 QA로 노출한다.
5) 머스크 5단계 실행안(순서 고정):
   5-1) 요구사항 의심: ‘보고서가 그럴듯하게 보이는 것’이 목표인지 ‘케이스 사실이 맞는 것’이 목표인지 분리한다.
   5-2) 삭제: 타 케이스에서 온 지명/수질값/그림/증빙을 케이스 입력에서 제거한다.
   5-3) 단순화/최적화: 템플릿은 케이스 로컬 템플릿 우선으로 고정해 실수 가능성을 줄인다.
   5-4) 가속: scoped QA/생성 결과가 같은 템플릿을 기준으로 나오도록 Gate Runner를 정렬한다.
   5-5) 자동화: 산출물 Desktop 복사 + “금지 지명/샘플 source_id” 탐지(사전/사후)로 재발을 막는다.
6) 비용/Idiot Index(낭비/마찰): ‘샘플 내용 섞임’은 수정 비용이 아니라 신뢰/리스크 비용이 가장 크다.
7) 리스크(인간/규제/2차효과/불확정): 데이터 공백(TBD)보다 잘못된 데이터가 더 위험하다(거짓-그린/거짓-완성).
8) 10배 옵션(2~4개): (A) DOCX 내 금지 토큰(지명/샘플 source_id) 스캔 게이트 (B) case.xlsx에서 특정 source_id 사용 금지 규칙 (C) SSOT reuse는 명시적 플래그/명시적 허용 리스트로만 활성화.
9) 최종 권장안(+지표/검증 루프): 신규 케이스는 “case-local 템플릿 + ERROR 0”을 기본으로 하고, 샘플 재사용은 SSOT_PAGE_OVERRIDES/승인된 범위에서만 제한적으로 사용한다.

FPT:
1) 문제 재정의: 체크박스/산출물 스냅샷이 서로 어긋나 “진척”이 과대평가/과소평가될 수 있다.
2) 가정(삭제 가능/검증 필요/필수 제약): 최소 제출 기준은 Gate-2에서 ERROR 0(가능하면 WARN 축소)이다.
3) 제1원칙 분해(Primitives): 입력(엑셀/출처/첨부) → 생성 → QA는 “재현 가능한 커맨드 + 결과 파일”로만 신뢰할 수 있다.
4) 재조립(새 시스템 설계): 체크박스(계획/준비) + Gate 스냅샷(품질/재현) 2축으로 진척을 운영한다.
5) 머스크 5단계 실행안(순서 고정):
   5-1) 요구사항 의심: Gate-2B WARN은 “입력 backlog”이며, template 누락으로 WARN이 인위적으로 0이 되지 않도록 Gate-1을 먼저 통과시킨다.
   5-2) 삭제: Gate-2 ERROR 0과 무관한 작업(지도/고급 자동화)은 우선순위에서 뒤로 민다.
   5-3) 단순화/최적화: Gate 커맨드를 `.venv/bin/*`로 고정해 환경 차이로 인한 오탐을 줄인다.
   5-4) 가속: 스냅샷을 문서에 남겨 “현재 상태” 확인 비용을 0에 가깝게 만든다.
   5-5) 자동화: 필요 시 Gate 결과를 요약해 문서에 반영하는 스크립트로 자동화한다.
6) 비용/Idiot Index(낭비/마찰): “ERROR 0로 착각 → 재작업” 루프가 가장 비싸므로 Gate 스냅샷을 SSOT로 둔다.
7) 리스크(인간/규제/2차효과/불확정): ERROR 0이라도 ‘근거 확보’가 끝난 것은 아니므로 출처/증빙(Traceability) 기준을 병행한다.
8) 10배 옵션(2~4개): 케이스 폴더마다 `_out_check_current/`를 표준화하고, Gate 결과를 자동 링크(핸드오프/리뷰)한다.
9) 최종 권장안(+지표/검증 루프): Gate-2(EIA) ERROR 0을 먼저 복구하고, 이후 WARN을 “모드별 정책”으로 단계적 축소한다.

### P-1 — 재현성/운영 게이트(전제)

- [x] P-1-1. 프로젝트 단위 git/repo 경계 및 산출물 관리 정책 고정(`eia-gen/` + 대용량 PDF/이미지/`output/` 처리 규약) (done-by: B)  
  - [x] reference pack 기반 재사용 구조는 존재(샘플팩/무결성)  
  - [x] (B) 대용량(샘플 PDF/첨부) 장기 보관/공유 규약(외부 아카이브/pack) SSOT 문서화 (done-by: B)  
  refs: `eia-gen/src/eia_gen/cli.py:940`, `eia-gen/reference_packs/CHANGWON_JINJEON_APPROVED_2025/manifest.json:1`, `eia-gen/docs/36_large_assets_storage_policy.md:1`

- [x] P-1-2. 실행 재현성 고정(필수 의존성/옵션 의존성/폴백 정책) (done-by: B)  
  - [x] PDF rasterize는 PyMuPDF 사용(`fitz`)  
  - [x] (B) OS별(Windows/macOS/Linux) 설치/실패 폴백(미설치 시 disabled + 안내) SSOT 문서화 (done-by: B)  
  refs: `eia-gen/src/eia_gen/services/figures/materialize.py:225`, `eia-gen/pyproject.toml:1`, `eia-gen/docs/34_installation_and_fallbacks.md:1`

- [x] P-1-3. 품질 게이트 3종을 “릴리즈 전 필수”로 운영 고정 (done-by: B)  
  - [x] 커맨드 존재: `eia-gen template-check`, `eia-gen check-xlsx-both`, `eia-gen verify-keys`  
  - [x] (B) 운영 체크리스트(runbook)로 고정(실패 분류/로그 위치 포함). CI는 선택(로컬 우선) (done-by: B)  
  refs: `eia-gen/src/eia_gen/cli.py:590`, `eia-gen/src/eia_gen/cli.py:997`, `eia-gen/src/eia_gen/cli.py:1302`, `eia-gen/docs/35_quality_gates_runbook.md:1`

- [x] P-1-4. 샘플 PDF/샘플 케이스의 SSOT 위치를 `reference_packs/` 중심으로 재정의(재현 가능한 “샘플팩”) (done-by: Codex)  
  - [x] `export-reference-pack`/`apply-reference-pack`/`init-case --reference-pack-id` 구현  
  - [x] `writer.py`의 샘플 PDF 경로 하드코딩을 pack/설정 기반으로 치환(레포 분리/공유 시 깨짐 방지) (done-by: Codex)  
  - [x] (Codex) (템플릿) 첨부/증빙 제외 “워터마크 없는 템플릿팩” 추가: `reference_packs/CHANGWON_SAMPLE_TEMPLATE_NO_ASSETS_2025` (done-by: Codex, 2026-01-04)  
  refs: `eia-gen/src/eia_gen/cli.py:940`, `eia-gen/src/eia_gen/services/reference_packs.py:123`, `eia-gen/src/eia_gen/services/writer.py:133`

- [x] P-1-5. “성능/캐시(효율성)”을 SSOT로 고정(래스터라이즈/정규화/지도 요청의 캐시 키·위치·만료 정책) (done-by: Codex)  
  - [x] figure materialize 캐시 키/재사용 구현 존재(단, mtime 기반)  
  - [x] WMS/WMTS 캐시 설계 SSOT는 존재(`cache.yaml`)  
  - [x] 케이스 스켈레톤에 `attachments/normalized`/`attachments/derived`/`attachments/evidence` 디렉터리 포함  
  - [x] 캐시 키를 “입력 sha256 + recipe(정규화 파라미터)” 기반으로 통일(재현성/재사용/디버그 용이) (done-by: Codex)  
  - [x] 캐시/파생 산출물 저장 위치를 케이스 폴더 중심(`attachments/derived`)으로 통일하고 out_dir은 “최종 산출물”로 한정 (done-by: Codex)  
  - [x] TTL/삭제 정책(예: 30일) + purge CLI 지원(선택) (done-by: Codex) — `eia-gen purge-derived --case-dir ... [--days 30] [--apply]`  
  - FPT(2026-01-02, Codex):
    ```text
    FPT:
    1) 문제 재정의: 파생 산출물(materialize/자동 생성 지도 등)이 out_dir(최종 산출물) 쪽에 흩어지면 (1) 케이스 이관/재현이 깨지고 (2) 중복 래스터라이즈로 시간·용량이 증가하며 (3) 정리/삭제 기준(TTL)이 흐려진다.
    2) 가정(삭제 가능/검증 필요/필수 제약): (a) 케이스 폴더는 입력+증빙+파생의 SSOT다. (b) out_dir은 “제출물/리포트” 용도다. (c) 파생 산출물은 재생성 가능해야 하며, 허위 도면 생성은 금지다.
    3) 제1원칙 분해(Primitives): 입력 바이트(sha256), 레시피(recipe), 파생 이미지, 케이스 폴더, 최종 산출물 폴더, TTL 기준(시간), 삭제/재생성 루프.
    4) 재조립(새 시스템 설계): 모든 파생 이미지는 케이스 `attachments/derived/` 하위로 고정하고(재사용/백업/이관), out_dir에는 DOCX/QA/리포트만 남긴다.
    5) 머스크 5단계 실행안(순서 고정):
       5-1) 요구사항 의심: “out_dir에 이미지가 있어야 한다”는 요구는 재현성/이관성 관점에서 재검토.
       5-2) 삭제: out_dir 기반 `_materialized` 산출을 제거/축소.
       5-3) 단순화/최적화: 파생 경로를 `attachments/derived`로 단일화.
       5-4) 가속: sha256+recipe 캐시로 반복 래스터라이즈 비용 제거.
       5-5) 자동화: `purge-derived`로 TTL 정리 루프 제공.
    6) 비용/Idiot Index(낭비/마찰): 디스크 무한 증가, 재실행마다 래스터라이즈 반복, “파일 어디 갔지?” 탐색 비용.
    7) 리스크(인간/규제/2차효과/불확정): 파생 파일 삭제로 재생성 필요(시간↑), 경로 변경으로 기존 스크립트/문서 드리프트(문서 갱신 필요).
    8) 10배 옵션(2~4개): (A) gates에 derived 용량/파일 수 경고 추가, (B) case 종료 시 자동 purge, (C) 원격 캐시(선택).
    9) 최종 권장안(+지표/검증 루프): 파생은 `attachments/derived` 고정 + TTL purge 제공. 지표: out_dir 내 그림 파일 0개 유지, 동일 입력 재실행 시 materialize 재사용(생성시간↓).
    ```
  refs: `eia-gen/src/eia_gen/services/figures/materialize.py:304`, `eia-gen/config/cache.yaml:1`, `eia-gen/src/eia_gen/cli.py:692`

- [x] P-1-6. 사용자 관점 “최소 작업 흐름(3단계)”를 SSOT로 고정하고 문서/CLI 출력에 일관되게 반영(실수/재작업 감소) (done-by: B+Codex)  
  - [x] 기본 골격 CLI는 이미 존재: `init-case` → `ingest-attachments` → `generate-xlsx-both`  
  - [x] (B) 권장 실행 시퀀스를 `docs/20_user_manual_ko.md`에 “한 화면”으로 고정(필수 입력/폴백/에러 후 다음 액션 포함) (done-by: B)  
  - [x] CLI 에러/QA에서 “다음 액션”을 구체적으로 출력(예: 어떤 시트/컬럼/행을 채워야 하는지)  
    - [x] (B) QA 결과를 “다음 액션(시트/컬럼 힌트)”로 변환하는 후처리 스크립트 제공(운영 우회) (done-by: B)  
      - NOTE(2026-01-02, Codex): `qa_next_actions.py`가 (1) 누락되던 입력 경로 매핑(예: TOPO/POP_TRAFFIC/TRACKER)과 (2) table specs 기반 폴백(예: CH3_SCOPING/CH2_BASELINE_SUMMARY), (3) DATA_REQUESTS/S-TBD 가이드를 추가로 포함하도록 개선됨.  
    - [x] (A) 엔진/CLI에 내장(실시간 출력) 여부 결정 및 연결 (done-by: Codex)  
  - [x] (입력 UX) v2 템플릿/문서 드리프트 제거  
    - [x] (A) `SSOT_PAGE_OVERRIDES` 중복 선언 제거(템플릿 생성부) (done-by: Codex)  
    - [x] (B) 문서의 “ASSETS” 용어를 “FIGURES/ATTACHMENTS”로 통일(혼동 제거) (done-by: B)  
  refs: `eia-gen/docs/20_user_manual_ko.md:72`, `eia-gen/scripts/qa_next_actions.py:1`, `eia-gen/src/eia_gen/cli.py:283`, `eia-gen/src/eia_gen/cli.py:418`, `eia-gen/src/eia_gen/cli.py:1043`, `eia-gen/src/eia_gen/services/xlsx/case_template_v2.py:159`

- [x] P-1-7. “환경/설치/의존성” 자동 점검(doctor) + 실패 시 가이드(편의성/운영 안정화) (done-by: B)  
  - [x] 키/엔드포인트 최소 검증 CLI(`verify-keys`) 존재  
  - [x] (B) PDF/OCR/후처리 의존성 점검: PyMuPDF/Pages/Word/LibreOffice(soffice) 가용성 + 폴백 힌트 출력 (done-by: B)  
  - [x] (B) venv 미사용(Anaconda/base)에서 “의존성 누락” 오탐을 줄이기 위해, `sys.prefix` 기반 venv 감지 + 실행 힌트 출력 (done-by: B)  
    refs: `eia-gen/scripts/doctor_env.py:420`
  - [x] (B) (선택) 워터마크/캡션용 폰트팩(레포 포함) 도입 시, 폰트 파일 존재 여부까지 점검(현재는 기본폰트 폴백) (done-by: B)  
    refs: `eia-gen/scripts/doctor_env.py:104`, `eia-gen/.env.example:1`, `eia-gen/docs/34_installation_and_fallbacks.md:1`
  - [x] (B) 실행 전 필수 파일 체크: `case.xlsx` 필수 시트 + `sources.yaml` 파싱 + “엑셀 참조 src_id 누락”/“file_path 누락” 점검 (done-by: B)  
  - [x] (B) (운영 편의) `case.xlsx`가 참조하지만 `sources.yaml`에 없는 `source_id`를 stub로 자동 추가하는 스크립트 제공(선택) (done-by: B)  
  refs: `eia-gen/scripts/doctor_env.py:181`, `eia-gen/scripts/doctor_env.py:394`, `eia-gen/scripts/fix_missing_sources.py:1`, `eia-gen/docs/20_user_manual_ko.md:254`

- [x] P-1-8. Gate-2 지표 이중화(Full vs Template scope) + CLI 지원(드리프트 방지) (done-by: Codex)  
  - [x] `check-xlsx-both`에서 template-scoped QA 지원: `--template-eia/--template-dia/--scope-to-default-templates` (done-by: Codex)  
  - [x] Gate-2A/B를 runbook/플랜에 표준화하고, 케이스별 `_out_check_current/`/`_out_check_scoped/` 경로를 고정 (done-by: Codex)  
  refs: `eia-gen/src/eia_gen/cli.py:1013`, `eia-gen/src/eia_gen/cli.py:539`, `eia-gen/docs/35_quality_gates_runbook.md:1`, `eia-gen/docs/11_execution_plan.md:128`

- [x] P-1-9. quality gates one-shot에서 “SSOT(default) vs normal(backlog)” scoped QA를 함께 산출 + 통계 요약 저장(거짓-그린 방지) (done-by: Codex)  
  - [x] `scripts/run_quality_gates.py`가 `check_scoped_default` + `check_scoped_normal`을 모두 실행하고 `summary.json`에 error/warn/placeholder를 기록 (done-by: Codex)  
  - [x] `generate-xlsx-both`/`check-xlsx-both --scope-to-default-templates` 기본 템플릿은 사용자 매뉴얼과 동일하게 SSOT full 우선으로 복원 (done-by: Codex)  
  refs: `eia-gen/scripts/run_quality_gates.py:1`, `eia-gen/src/eia_gen/cli.py:395`, `eia-gen/docs/35_quality_gates_runbook.md:1`, `eia-gen/docs/20_user_manual_ko.md:330`

### P0 — 창원 샘플 기반 “입력 세팅 SSOT” 고정(기준점)

- [x] P0-1. 샘플 PDF 인덱스(챕터/표/그림) 결과를 “버전 고정” 폴더로 승격(재현 가능한 SSOT) (done-by: B)  
  - [x] 인덱스/요약 산출물은 존재(현 위치)  
  - [x] reference pack으로 승격(경로 안정화) (done-by: B)  
  - [x] (B) reference pack `manifest.json` sha256 무결성 검증 스크립트 추가(운영/이관 시 깨짐 조기 탐지) (done-by: B)  
    refs: `eia-gen/scripts/verify_reference_pack_manifest.py:1`, `eia-gen/reference_packs/CHANGWON_JINJEON_APPROVED_2025/manifest.json:1`
  - [x] (B) reference pack `manifest.json` sha256 갱신 스크립트 추가(재생성 후 수동 편집 제거) (done-by: B)  
    refs: `eia-gen/scripts/update_reference_pack_manifest.py:1`, `eia-gen/docs/36_large_assets_storage_policy.md:26`
  refs: `eia-gen/reference_packs/CHANGWON_JINJEON_APPROVED_2025/pdf_index/summary.md:1`, `eia-gen/src/eia_gen/cli.py:1144`

- [x] P0-2. “샘플 표/그림 캡션 목록 → 우리 SSOT 스펙 ID” 대응표(coverage matrix) 작성 (done-by: B)  
  - [x] EXIST/NEED/IGNORE 3분류 + NEED는 티켓화 (done-by: B)  
    refs: `eia-gen/scripts/coverage_matrix_auto_status.py:1`, `eia-gen/reference_packs/CHANGWON_JINJEON_APPROVED_2025/pdf_index/coverage_matrix_seed.suggested.xlsx:1`, `eia-gen/reference_packs/CHANGWON_JINJEON_APPROVED_2025/pdf_index/coverage_need_tickets.md:1`, `eia-gen/docs/31_coverage_matrix_workflow.md:1`
  - [x] (B) 캡션 목록을 coverage matrix seed(XLSX/CSV)로 자동 생성(초안 템플릿) (done-by: B)
    refs: `eia-gen/scripts/make_coverage_matrix_seed.py:1`, `eia-gen/reference_packs/CHANGWON_JINJEON_APPROVED_2025/pdf_index/coverage_matrix_seed.xlsx:1`
  - [x] (B) our_spec_id 후보 자동 추천(초안, fuzzy match) 스크립트 추가 + suggested seed 생성 (done-by: B)  
    refs: `eia-gen/scripts/suggest_coverage_spec_ids.py:1`, `eia-gen/reference_packs/CHANGWON_JINJEON_APPROVED_2025/pdf_index/coverage_matrix_seed.suggested.xlsx:1`, `eia-gen/docs/31_coverage_matrix_workflow.md:26`
  - [x] (B) suggest_1 고신뢰 행 기반 our_spec_id 자동 채움 스크립트 추가(복붙 감소, 오매핑 리스크↓) (done-by: B)  
    refs: `eia-gen/scripts/coverage_matrix_auto_fill_from_suggestions.py:1`, `eia-gen/docs/31_coverage_matrix_workflow.md:1`
  - [x] (B) suggested seed를 reference pack(SSOT)로 승격 + manifest에 sha256 기록(경로 안정화) (done-by: B)  
    refs: `eia-gen/reference_packs/CHANGWON_JINJEON_APPROVED_2025/manifest.json:1`, `eia-gen/reference_packs/CHANGWON_JINJEON_APPROVED_2025/pdf_index/coverage_matrix_seed.suggested.xlsx:1`
  - [x] (B) coverage matrix에서 NEED/UNCLASSIFIED 티켓 목록(md) 자동 생성 스크립트 추가 + SSOT 보관본 생성 (done-by: B)  
    refs: `eia-gen/scripts/make_coverage_need_tickets.py:1`, `eia-gen/reference_packs/CHANGWON_JINJEON_APPROVED_2025/pdf_index/coverage_need_tickets.md:1`, `eia-gen/docs/31_coverage_matrix_workflow.md:43`
  - [x] (B) coverage matrix의 `our_spec_id` 유효성 검증 + status(EXIST/NEED) 자동 정리 스크립트 추가(반복 수작업 감소) (done-by: B)  
    refs: `eia-gen/scripts/coverage_matrix_auto_status.py:1`, `eia-gen/docs/31_coverage_matrix_workflow.md:1`
  - [x] (B) 사용자 매뉴얼에 coverage matrix 워크플로/티켓 자동 생성을 “한 화면”으로 추가(운영 편의) (done-by: B)  
    refs: `eia-gen/docs/20_user_manual_ko.md:292`, `eia-gen/docs/31_coverage_matrix_workflow.md:1`
  refs: `eia-gen/docs/31_coverage_matrix_workflow.md:1`, `eia-gen/reference_packs/CHANGWON_JINJEON_APPROVED_2025/pdf_index/summary.md:17`, `eia-gen/spec/figure_specs.yaml:7`, `eia-gen/spec_dia/figure_specs.yaml:7`

- [x] P0-3. 샘플 반복 “환경기준 표”를 최소 3종으로 우선 정의(EIA 필수 앵커/표 스펙/입력 폼) (done-by: Codex)  
  - [x] (B) “환경기준 표 3종” 스펙/앵커/출처 초안 문서화(옵션 A/B 비교) (done-by: B)
    refs: `eia-gen/docs/32_env_standards_tables_plan.md:1`, `eia-gen/reference_packs/CHANGWON_JINJEON_APPROVED_2025/pdf_index/summary.md:43`
  - [x] (Codex) 옵션 A(정적 표 내장)로 3종 구현 + 템플릿 앵커/스캐폴딩 반영
    refs: `eia-gen/spec/table_specs.yaml:1`, `eia-gen/spec/template_map.yaml:1`, `eia-gen/scripts/make_template_sample_changwon_2025.py:1`, `eia-gen/templates/report_template.sample_changwon_2025.scaffolded.docx:1`
  refs: `eia-gen/docs/32_env_standards_tables_plan.md:1`, `eia-gen/docs/sample_changwon_gingerfarm_2025_prompt_pack.md:119`, `eia-gen/reference_packs/CHANGWON_JINJEON_APPROVED_2025/pdf_index/summary.md:58`, `eia-gen/spec/template_map.yaml:5`

- [x] P0-4. (창원 케이스) 새 첨부 PDF를 SSOT 페이지 치환/부록에 반영(운영 예시를 SSOT로 고정) (done-by: B)  
  - [x] `ATT-0008__진동광암해수욕장.pdf` ingest 완료 + SSOT 치환 4건 적용(샘플 p2 → p1, p71 → p2, p324 → p6, p65 → p17) (done-by: B)  
  - [x] `ATT-0001__00토목도면-자동차야영장.pdf`의 위치도/계획평면/배수계획 등 치환 적용  
  - [x] (B) `ATT-0001` 페이지 맵(OCR): 위치도/계획평면/배수계획 페이지를 빠르게 찾고 SSOT_PAGE_OVERRIDES로 연결(운영 메모) (done-by: B)  
    refs: `eia-gen/docs/41_att0001_civil_drawings_usage.md:1`, `output/att0001_page_titles_matches.json:1`
  - [x] `ATT-0008`의 추가 페이지(컨셉/연계/접근성 등)를 “빈 샘플 슬롯/부록”으로 추가 치환(우선순위: 본문 앞단) (done-by: B)  
  - [x] (B) “빈 슬롯 부족/챕터 디바이더 치환 리스크” 운영 주의사항 문서화(후속: insert 기능 검토) (done-by: B)  
    refs: `eia-gen/docs/29_jindong_gwangam_pdf_usage.md:69`, `eia-gen/scripts/find_sparse_pages_in_pdf.py:1`
  refs: `eia-gen/docs/29_jindong_gwangam_pdf_usage.md:1`, `eia-gen/scripts/ssot_page_override_wizard.py:1`, `eia-gen/docs/24_changwon_progress_report.md:39`, `eia-gen/docs/23_changwon_input_checklist.md:45`, `eia-gen/output/case_new_max_reuse/case.xlsx:SSOT_PAGE_OVERRIDES#row8`, `eia-gen/output/case_new_max_reuse/case.xlsx:SSOT_PAGE_OVERRIDES#row14`, `eia-gen/output/case_new_max_reuse/case.xlsx:SSOT_PAGE_OVERRIDES#row15`, `eia-gen/output/case_new_max_reuse/case.xlsx:SSOT_PAGE_OVERRIDES#row16`

- [x] P0-5. (후속) `SSOT_PAGE_OVERRIDES`의 “치환(replace)” 한계(빈 슬롯 부족) 대응: **추가 삽입(insert)** 방식 설계/구현(샘플 흐름 유지) (done-by: Codex)  
  - [x] 템플릿에 “부록/삽입 전용 앵커(예: `[[BLOCK:APPENDIX_INSERTS]]`)”를 두고, 치환이 아니라 **docx 삽입**으로 페이지(PDF page → PNG)를 추가 (done-by: Codex)
  - [x] (B) insert 방식 설계 스펙(입력 시트/앵커/trace) 문서화(초안) — A 구현 착수용 (done-by: B)  
    refs: `eia-gen/docs/38_appendix_insert_spec.md:1`, `eia-gen/docs/29_jindong_gwangam_pdf_usage.md:69`
  - [x] (B) `APPENDIX_INSERTS` 페이지 목록 자동 등록(UPSERT) + 프리뷰 PNG 생성 스크립트 추가 + runbook + 데모 케이스 적용 (done-by: B)  
    refs: `eia-gen/scripts/make_appendix_inserts_from_pdf.py:1`, `eia-gen/docs/50_appendix_inserts_from_pdf_runbook.md:1`, `eia-gen/output/case_new_max_reuse/case.xlsx:APPENDIX_INSERTS#row24`, `eia-gen/output/case_new_max_reuse/case.xlsx:APPENDIX_INSERTS#row25`, `eia-gen/output/case_new_max_reuse/attachments/derived/previews/appendix/ATT-0008__진동광암해수욕장/INS-0023__p0003.png:1`
  refs: `eia-gen/docs/29_jindong_gwangam_pdf_usage.md:1`, `eia-gen/docs/38_appendix_insert_spec.md:1`, `eia-gen/spec/template_map.yaml:1`, `eia-gen/src/eia_gen/services/docx/spec_renderer.py:368`, `eia-gen/src/eia_gen/services/writer.py:916`, `eia-gen/scripts/run_quality_gates.py:1`

### P1 — “첨부 도면/이미지”를 샘플 수준으로 출력(핵심)

- [x] P1-0. (핵심 공백) Asset Normalization 단계를 서비스로 고정(원본→정규화PNG→판넬PNG→DOCX) (done-by: Codex)  
  - [x] 현재도 DOCX 삽입 시점에 PDF rasterize/crop/resize를 수행(materialize)  
  - [x] (B) “패딩/테두리/워터마크/표준 DPI”를 포함한 제출용 정규화 레시피(SSOT) 초안 추가 (done-by: B)  
    refs: `eia-gen/config/asset_normalization.yaml:1`, `eia-gen/docs/09_figure_generation_pipeline.md:1`
  - [x] 정규화 산출물(정규화PNG/판넬PNG)을 evidence로 승격(sha256/recipe/input_manifest)  
    refs: `eia-gen/src/eia_gen/services/figures/materialize.py:629`, `eia-gen/src/eia_gen/services/docx/spec_renderer.py:256`, `eia-gen/src/eia_gen/services/docx/builder.py:219`, `eia-gen/src/eia_gen/services/figures/derived_evidence.py:42`, `eia-gen/src/eia_gen/services/export/source_register_xlsx.py:739`, `eia-gen/tests/test_materialize_traceability.py:15`
  - [x] 정규화/파생 산출물 저장 위치를 케이스 폴더 `attachments/normalized`/`attachments/derived`로 통일(재사용/이력/백업 편의)  
    refs: `eia-gen/src/eia_gen/cli.py:823`, `eia-gen/src/eia_gen/cli.py:1145`, `eia-gen/src/eia_gen/services/docx/builder.py:287`, `eia-gen/src/eia_gen/services/docx/spec_renderer.py:805`, `eia-gen/src/eia_gen/services/figures/photo_sheet_generate.py:244`
  - [x] 캐시 키를 mtime 기반이 아니라 “입력 sha256 + recipe” 기반으로 전환(재현성 + 효율성 + 디버그 용이) (done-by: Codex)  
  refs: `eia-gen/docs/26_drr_guideline_2025_05_plan_delta.md:57`, `eia-gen/src/eia_gen/cli.py:692`, `eia-gen/src/eia_gen/services/ingest_attachments.py:98`, `eia-gen/src/eia_gen/services/figures/materialize.py:295`, `eia-gen/src/eia_gen/services/figures/materialize.py:304`

  FPT:
  1) 문제 재정의: Normalize/Compose로 생성되는 PNG(정규화PNG/판넬PNG)가 “파일은 생성되지만” 재현에 필요한 입력/레시피/해시가 증빙으로 남지 않으면, `source_register.xlsx`에서 감사/보완요구 대응(어떤 파일이 어떤 규격으로 생성되었나)을 자동으로 할 수 없다.
  2) 가정(삭제 가능/검증 필요/필수 제약): (1) 입력 `case.xlsx`는 자동 수정하지 않는다. (2) 동일 입력+레시피면 동일 결과를 재사용한다. (3) 공식 도면/사진을 ‘그럴듯하게 생성’하지 않고, 정규화(크롭/패딩/워터마크/포맷)만 수행한다.
  3) 제1원칙 분해(Primitives): (a) src_sha256 (b) recipe (c) out_sha256 (d) file_path (e) fig_id/anchor (f) src_id 연결.
  4) 재조립(새 시스템 설계): materialize/compose 결과에 `recipe`+`input_manifest`를 포함한 note JSON을 부여하고, `derived_evidence_manifest`→`source_register.xlsx:EVIDENCE_INDEX/USAGE_REGISTER`로 자동 연결한다.
  5) 머스크 5단계 실행안(순서 고정):
     5-1) 요구사항 의심: “완벽한 전수 입력” 이전에, 생성물의 sha256+레시피를 먼저 잠가 재현성을 확보한다.
     5-2) 삭제: 생성 후 수동으로 메모/증빙을 붙이는 프로세스를 배제한다(누락 원인).
     5-3) 단순화/최적화: 생성 시점에 note(JSON)만 기록하고, export에서 표준 시트로 모은다.
     5-4) 가속: 유닛테스트로 recipe/input_manifest 기록을 회귀 방지한다.
     5-5) 자동화: Gate 산출물(`source_register.xlsx`)에서 증빙/사용처 연결을 점검한다.

- [x] P1-1. FIGURES 시트의 `width_mm/crop/gen_method`를 실제 렌더에 반영(사진/도면 품질 최소선) (done-by: Codex)  
  - [x] 템플릿/빌트인 렌더러 모두 materialize(PDF rasterize/crop/resize) 적용 중  
  - [x] materialized 산출물을 `attachments/derived`로 통일 저장(케이스 폴더 기준) (done-by: Codex)  
  - [x] 삽입된 그림(특히 materialize 결과)이 자동으로 `ATTACHMENTS`/`source_register.xlsx`의 사용처(Claims)로 연결 (done-by: Codex)  
  - [x] PDF 페이지 지정이 없을 때 자동 선택(휴리스틱) 결과를 QA WARN로 노출 (done-by: Codex)  
  - [x] DOCX 삽입 폭을 “페이지 가용폭(여백 제외)” 기준으로 clamp(폭 과다로 레이아웃 깨짐 방지) (done-by: Codex)  
  refs: `eia-gen/src/eia_gen/services/docx/spec_renderer.py:1`, `eia-gen/src/eia_gen/services/docx/builder.py:1`, `eia-gen/src/eia_gen/services/figures/materialize.py:1`, `eia-gen/src/eia_gen/services/figures/derived_evidence.py:1`, `eia-gen/src/eia_gen/services/figures/spec_figures.py:1`, `eia-gen/src/eia_gen/services/qa/run.py:1`, `eia-gen/src/eia_gen/services/export/source_register_xlsx.py:1`, `eia-gen/tests/test_source_register_xlsx.py:1`, `eia-gen/tests/test_pdf_page_selection.py:1`

  FPT:
  1) 문제 재정의: DOCX 삽입 과정에서 생성되는 materialize PNG가 “증빙(evidence)”로 승격되지 않아, `source_register.xlsx`에서 그림의 재현/추적이 끊긴다.
  2) 가정(삭제 가능/검증 필요/필수 제약): (1) 입력 XLSX를 자동 수정하지 않는다. (2) 동일 입력/옵션이면 동일 파생 산출물을 재사용한다. (3) 사용자는 보완요구 대응에서 ‘어떤 파일/페이지가 어디에 쓰였는지’를 먼저 찾는다.
  3) 제1원칙 분해(Primitives): (a) 실제 삽입된 파일 경로(derived) (b) 그림 식별자(fig_id) (c) 출처(src_ids) (d) 레지스트리 출력(Evidence/Claims).
  4) 재조립(새 시스템 설계): 렌더러가 materialize 산출물을 `case.model_extra.derived_evidence_manifest`에 기록하고, `source_register.xlsx`에서 Evidence/Claims로 연결한다.
  5) 머스크 5단계 실행안(순서 고정):
     5-1) 요구사항 의심: “완벽한 4시트 Traceability” 이전에, 최소한 그림 evidence_id만이라도 자동 연결한다.
     5-2) 삭제: 입력 XLSX 자동 업데이트(부작용/충돌)를 우선 범위에서 제외한다.
     5-3) 단순화/최적화: evidence_id는 파생 파일명(stem) 기반으로 안정 생성하고 중복을 제거한다.
     5-4) 가속: 유닛테스트로 Claims evidence_id 채움 회귀를 고정한다.
     5-5) 자동화: quality gates에서 생성된 `source_register.xlsx`로 실사용 스냅샷을 검증한다.
  6) 비용/Idiot Index(낭비/마찰): “어느 그림이 어디에 쓰였는지” 수동 추적 시간을 제거한다.
  7) 리스크(인간/규제/2차효과/불확정): 템플릿에 존재하지만 draft에 없는 그림은 Claims에 누락될 수 있다(차기: 렌더 결과 기반 claim 보강).
  8) 10배 옵션(2~4개): (A) normalize/compose 산출물까지 동일 방식으로 evidence/claims 연결 (B) figure 삽입 폭 clamp + 레시피/sha256로 결정론 강화.
  9) 최종 권장안(+지표/검증 루프): `EV-DERIVED-*`가 Evidence에 포함되고, 해당 fig_id의 Claims evidence_id가 자동 채워지는지(Gate 산출물)로 검증한다.

- [x] P1-2. “사진대지/콜아웃 판넬” 생성 레시피(`CALLOUT_COMPOSITE`) 구현(샘플 스타일 고정) (done-by: Codex)  
  - [x] (B) 스크립트 프로토타입 + runbook + 예제 산출물 추가(현장 운용: “합성 PNG 생성 → FIGURES.file_path로 등록”) (done-by: B)  
    refs: `eia-gen/scripts/compose_callout_composite.py:1`, `eia-gen/scripts/make_callout_composite_from_case.py:1`, `eia-gen/docs/44_callout_composite_runbook.md:1`, `eia-gen/output/annotation_examples/att0008_p1_4_callout_composite.png`, `eia-gen/docs/20_user_manual_ko.md:471`
  - [x] (B) 데모 케이스로 “원클릭” E2E 검증(ingest→related_fig_id 매핑→callout 합성→FIGURES.file_path 갱신) (done-by: B)  
    refs: `eia-gen/output/case_callout_demo/case.xlsx:FIELD_SURVEY_LOG#row2`, `eia-gen/output/case_callout_demo/case.xlsx:ATTACHMENTS#row10`, `eia-gen/output/case_callout_demo/case.xlsx:ATTACHMENTS#row11`, `eia-gen/output/case_callout_demo/case.xlsx:ATTACHMENTS#row12`, `eia-gen/output/case_callout_demo/case.xlsx:ATTACHMENTS#row13`, `eia-gen/output/case_callout_demo/case.xlsx:ATTACHMENTS#row14`, `eia-gen/output/case_callout_demo/case.xlsx:FIGURES#row2`, `eia-gen/output/case_callout_demo/attachments/derived/figures/callout/FIG-PHOTO-SHEET-01.png`
  - [x] (B) 고정 그리드 3종만(2/4/6) 지원 + 슬롯 구성(cover 크롭 + 캡션바 + 번호 배지) (done-by: B, script-prototype)  
    refs: `eia-gen/scripts/compose_callout_composite.py:1`, `eia-gen/config/figure_style.yaml:69`, `eia-gen/output/case_callout_demo/attachments/derived/figures/callout/FIG-PHOTO-SHEET-01.png`
  - [x] (B) 입력 연결 규약 고정(권장 1안): `ATTACHMENTS.related_fig_id == FIGURES.fig_id`로 사진 묶음을 정의(캡션은 `ATTACHMENTS.title/note` 우선) (done-by: B, wrapper-prototype)  
    refs: `eia-gen/scripts/make_callout_composite_from_case.py:164`, `eia-gen/scripts/make_callout_composite_from_case.py:169`, `eia-gen/output/case_callout_demo/case.xlsx:ATTACHMENTS#row10`, `eia-gen/output/case_callout_demo/case.xlsx:FIGURES#row2`
  - [x] (B) (선택 2안) `FIELD_SURVEY_LOG.photo_folder` 기반 자동 수집→ATTACHMENTS 자동 등록→related_fig_id 매핑까지 “원클릭” UX 제공 (done-by: B, wrapper-prototype)  
    refs: `eia-gen/scripts/make_callout_composite_from_case.py:316`, `eia-gen/scripts/make_callout_composite_from_case.py:354`, `eia-gen/scripts/make_callout_composite_from_case.py:365`, `eia-gen/output/case_callout_demo/case.xlsx:FIELD_SURVEY_LOG#row2`, `eia-gen/output/case_callout_demo/attachments/attachments_manifest.json:1`
  - [x] (확장) 사진 메타(촬영일/방향/지점/좌표) 전용 시트(예: `PHOTO_LOG`) 도입 여부 결정(샘플 재현 요구 강도에 따라) — **현 단계에서는 미도입(보류)**: `FIELD_SURVEY_LOG` + `ATTACHMENTS.note`로 운영하고, 필요 시 P3-2(현장조사 가드레일)에서 재검토 (done-by: Codex)  
  - [x] (B) 결과 PNG는 “결정론적(동일 입력이면 동일 sha256)”으로 생성(골든 테스트/회귀 검증 가능) (done-by: B)  
    refs: `eia-gen/output/case_callout_demo/_tmp_callout_sha_before.txt:1`, `eia-gen/output/case_callout_demo/_tmp_callout_sha_after.txt:1`
  - [x] (B) (stopgap) 결과 PNG를 ATTACHMENTS에 등록(UPSERT, evidence_id=`DER-{fig_id}`, note에 sha256 기록) (done-by: B)  
    refs: `eia-gen/scripts/make_callout_composite_from_case.py:280`, `eia-gen/scripts/make_callout_composite_from_case.py:420`, `eia-gen/scripts/make_callout_composite_from_case.py:442`, `eia-gen/output/case_callout_demo/case.xlsx:ATTACHMENTS#row14`
  - [x] (A) (정식) 결과 PNG를 evidence/source_register.xlsx에 연결(sha256/recipe/usage까지 “완전한 traceability”) (done-by: Codex)  
    refs: `eia-gen/src/eia_gen/services/figures/photo_sheet_generate.py:1`, `eia-gen/src/eia_gen/cli.py:313`, `eia-gen/tests/test_photo_sheet_generate.py:1`
  - [x] (B) 임시 운영: wrapper가 `FIGURES.file_path` 갱신까지 수행(코어 미연결 상태에서도 삽입 가능) (done-by: B)  
    refs: `eia-gen/scripts/make_callout_composite_from_case.py:446`, `eia-gen/scripts/make_callout_composite_from_case.py:455`, `eia-gen/output/case_callout_demo/case.xlsx:FIGURES#row2`
  - [x] (B) 비주얼 스타일 토큰 SSOT 초안 추가(콜아웃/범례/사진대지) — 구현은 A가 파싱하여 사용 (done-by: B)  
    refs: `eia-gen/config/figure_style.yaml:1`, `eia-gen/docs/37_sample_figure_style_guide.md:1`
  - [x] (B) 제공 이미지/도면(PNG/JPG/PDF→PNG)에 polygon/label/번호/범례를 “스펙 기반”으로 합성하는 프로토타입(스키마+스크립트) 추가 (done-by: B)  
    refs: `eia-gen/docs/40_image_annotation_spec.md:1`, `eia-gen/scripts/annotate_image.py:1`, `eia-gen/config/figure_style.yaml:1`
  - [x] (B) 반투명 오버레이(슬라이드/항공사진)에서 폴리곤을 자동 추출→주석 스펙으로 변환하는 스크립트+runbook 추가(샘플 수준 라벨/경계 재현 보조) (done-by: B)  
    refs: `eia-gen/scripts/extract_overlay_polygons.py:1`, `eia-gen/docs/42_overlay_polygon_extraction_runbook.md:1`, `eia-gen/output/annotation_examples/att0008_p2_overlays.yaml:1`, `eia-gen/output/annotation_examples/att0008_p2_annotated_ko.png`
  - [x] (B) 오버레이 자동 추출 시 노이즈 영역을 줄이기 위한 ROI/제외영역(`--roi/--exclude-rect`) 옵션 추가(반복 튜닝 비용↓) (done-by: B)  
    refs: `eia-gen/scripts/extract_overlay_polygons.py:1`, `eia-gen/docs/42_overlay_polygon_extraction_runbook.md:45`
  - [x] (B) 오버레이 자동 추출 결과에 번호 배지/범례 레이어를 옵션으로 포함(`--emit-number-badges/--emit-legend`) + 예제 산출물 추가 (done-by: B)  
    refs: `eia-gen/scripts/extract_overlay_polygons.py:1`, `eia-gen/output/annotation_examples/att0008_p2_overlays_with_badges.yaml:1`, `eia-gen/output/annotation_examples/att0008_p2_annotated_badges_ko.png`
- [x] (B) GeoJSON(경계) + bbox(=WMS/WMTS) → `WORLD_LINEAR_BBOX` 주석 YAML 자동 생성 스크립트+runbook+예제 추가(“사업부지 경계”를 지도에 바로 합성) (done-by: B)  
    refs: `eia-gen/scripts/geojson_to_annotations.py:1`, `eia-gen/docs/43_geojson_to_annotations_runbook.md:1`, `eia-gen/output/annotation_examples/site_boundary_on_wms_landslide.yaml:1`, `eia-gen/output/annotation_examples/site_boundary_on_wms_landslide_annotated.png`, `eia-gen/output/case_new_max_reuse/attachments/evidence/gis/EV-REQ-AUTO-GIS-ZONING-OVERLAY-WMS-20251230-112037_REQ-AUTO-GIS-ZONING-OVERLAY-WMS_wms_overlay.csv:1`
  - [x] (B) annotate(합성) 스크립트에서 `image_size: 512x512` 문자열 파싱 + WORLD_LINEAR_BBOX 스케일링(리사이즈된 지도에도 정합) 보강 (done-by: B)  
    refs: `eia-gen/scripts/annotate_image.py:1`, `eia-gen/scripts/geojson_to_annotations.py:1`, `eia-gen/output/annotation_examples/site_boundary_on_wms_landslide.yaml:45`
  - [x] (B) (추가) “마스크(흑백 PNG) → 폴리곤 → 주석 YAML” 변환 도구 + runbook + 예제 추가(오버레이/경계가 이미 마스크로 주어진 경우) (done-by: B)  
    refs: `eia-gen/scripts/mask_to_polygons.py:1`, `eia-gen/docs/47_mask_to_polygons_runbook.md:1`, `eia-gen/output/annotation_examples/att0008_p2_debug/mask.png:1`, `eia-gen/output/annotation_examples/att0008_p2_mask_polys.yaml:1`, `eia-gen/output/annotation_examples/att0008_p2_mask_polys_annotated.png:1`
  - [x] (B) (wrapper) case.xlsx + (base image + annotations.yaml) → 주석 합성 PNG 생성 + `FIGURES.file_path`/`ATTACHMENTS(note=sha256)` 등록(운영형 “원클릭” UX) (done-by: B)  
    refs: `eia-gen/scripts/make_annotated_figure_from_case.py:1`, `eia-gen/docs/48_annotate_figure_from_case_runbook.md:1`, `eia-gen/output/case_callout_demo/case.xlsx:FIGURES#row3`, `eia-gen/output/case_callout_demo/case.xlsx:ATTACHMENTS#row15`, `eia-gen/output/case_callout_demo/attachments/derived/figures/annotated/FIG-ANNO-01.png:1`, `eia-gen/output/case_callout_demo/_tmp_anno_sha_before.txt:1`, `eia-gen/output/case_callout_demo/_tmp_anno_sha_after.txt:1`
  - [x] (B) (wrapper) case.xlsx + (mask PNG + base image) → polygons → annotations.yaml → 주석 합성 PNG 생성 + `FIGURES.file_path`/`ATTACHMENTS(note=sha256)` 등록(“마스크 기반” 원클릭) (done-by: B)  
    refs: `eia-gen/scripts/make_mask_annotated_figure_from_case.py:1`, `eia-gen/docs/49_mask_annotate_from_case_runbook.md:1`, `eia-gen/scripts/mask_to_polygons.py:1`, `eia-gen/output/case_callout_demo/case.xlsx:FIGURES#row4`, `eia-gen/output/case_callout_demo/case.xlsx:ATTACHMENTS#row16`, `eia-gen/output/case_callout_demo/attachments/derived/annotations/FIG-ANNO-MASK-01_mask.yaml:1`, `eia-gen/output/case_callout_demo/attachments/derived/figures/annotated/FIG-ANNO-MASK-01.png:1`, `eia-gen/output/case_callout_demo/_tmp_anno_mask_sha_before.txt:1`, `eia-gen/output/case_callout_demo/_tmp_anno_mask_sha_after.txt:1`
  - [x] (B) 샘플(창원/기허가) 스타일 토큰/패턴을 문서화(초안) — 구현은 A가 SSOT로 승격하여 사용 (done-by: B)  
    refs: `eia-gen/docs/37_sample_figure_style_guide.md:1`, `eia-gen/docs/09_figure_generation_pipeline.md:1`
  refs: `eia-gen/docs/09_figure_generation_pipeline.md:61`, `eia-gen/docs/26_drr_guideline_2025_05_plan_delta.md:90`, `eia-gen/src/eia_gen/services/xlsx/case_template_v2.py:125`, `eia-gen/src/eia_gen/services/xlsx/case_template_v2.py:144`, `eia-gen/src/eia_gen/services/xlsx/case_template_v2.py:204`

  FPT:
  1) 문제 재정의: 스크립트로 생성된 사진대지 PNG가 코어 파이프라인/`source_register.xlsx`의 Evidence/Claims에 자동 연결되지 않아, 보완요구 대응 시 “어떤 사진 묶음이 어떤 그림(figure)에 사용되었는지” 추적이 끊긴다.
  2) 가정(삭제 가능/검증 필요/필수 제약): (1) 입력 XLSX를 자동 수정하지 않는다. (2) 사진 묶음은 `ATTACHMENTS.related_fig_id`로만 정의한다. (3) 최대 6장(2/4/6 그리드)만 지원한다. (4) 합성은 배치/표시만 수행한다(허위 상태/수치 생성 금지).
  3) 제1원칙 분해(Primitives): (a) 입력 사진 evidence_id/파일경로 (b) figure_id (c) 출력 PNG 경로 (d) sha256 (e) 레시피(grid/캡션).
  4) 재조립(새 시스템 설계): `ensure_photo_sheets_from_attachments()`가 `attachments/derived/figures/callout/{fig_id}.png`를 생성하고, `case.assets`와 `case.derived_evidence_manifest`를 갱신해 Evidence/Claims가 자동 연결되게 한다.
  5) 머스크 5단계 실행안(순서 고정):
     5-1) 요구사항 의심: “사진 메타 전용 시트” 도입 전, figure evidence_id 자동 연결을 최소 목표로 한다.
     5-2) 삭제: `ATTACHMENTS`/`FIGURES` 시트 in-place 업데이트를 정식 범위에서 제외한다.
     5-3) 단순화/최적화: 출력 경로/파일명은 fig_id 기반으로 고정하고, note에 sha256+recipe(JSON)만 기록한다.
     5-4) 가속: 유닛테스트로 PNG 생성/asset 반영/derived evidence 기록을 고정한다.
     5-5) 자동화: quality gates 산출물의 `source_register.xlsx`에서 해당 figure의 Claims evidence_id 채움을 검증한다.
  6) 비용/Idiot Index(낭비/마찰): 수동으로 “사진대지 파일 생성→FIGURES 등록→증빙 연결”을 반복하는 비용을 제거한다.
  7) 리스크(인간/규제/2차효과/불확정): 사용자가 이미 수동으로 다른 사진대지를 등록한 경우 자동 생성이 덮어쓰지 않도록(존중) skip 정책을 둔다.
  8) 10배 옵션(2~4개): (A) 입력 사진 캡션을 별도 시트로 구조화(향후) (B) 사진대지 생성도 sha256+recipe 캐시로 재사용 (C) figure 삽입 결과까지 claim 기반 역추적 강화.
  9) 최종 권장안(+지표/검증 루프): `EV-DERIVED-FIG-...`가 Evidence Register에 기록되고, 해당 figure의 Claims(figure).evidence_id가 자동 채워지는지로 검증한다.

- [x] P1-3. “도면 PDF → 보고서용 PNG 정리” 레시피 고정(페이지 지정/자동 크롭/여백/테두리/워터마크)  
  - [x] PDF→PNG 래스터라이즈 + AUTO crop/resize + pad/border frame(삽입 시점) 존재  
  - [x] 이미지 EXIF 방향 교정 존재(사진 입력 안정화)  
  - [x] “도면은 A3 이하 권장” 같은 제출 규격(폭/해상도) 규칙을 레시피에 반영(= QA 경고로 조기 노출) (done-by: Codex)  
    refs: `eia-gen/src/eia_gen/services/qa/run.py:1`, `eia-gen/docs/08_qa_checklist_rules.md:1`
  - [x] REFERENCE일 때 워터마크/캡션 자동 강제(스키마+QA로 연결) (done-by: Codex)  
    refs: `eia-gen/src/eia_gen/services/figures/spec_figures.py:1`, `eia-gen/src/eia_gen/services/figures/materialize.py:1`, `eia-gen/src/eia_gen/services/docx/spec_renderer.py:1`, `eia-gen/config/asset_normalization.yaml:1`
  - [x] (효율성) 사진/도면 포맷 정책 고정: 도면=PNG(선명), 사진=JPEG(용량↓) 등 “asset_role 기반” 선택(선택 구현) (done-by: Codex)  
    refs: `eia-gen/config/asset_normalization.yaml:1`, `eia-gen/src/eia_gen/services/figures/materialize.py:1`
  - [x] (효율성) `target_dpi/max_width_px` 기본값/override 정책을 SSOT로 고정(DOCX 용량/속도 예측 가능) (done-by: Codex)  
    refs: `eia-gen/config/asset_normalization.yaml:1`, `eia-gen/src/eia_gen/services/docx/spec_renderer.py:1`, `eia-gen/src/eia_gen/services/figures/materialize.py:1`
  refs: `eia-gen/src/eia_gen/services/figures/materialize.py:17`, `eia-gen/src/eia_gen/services/figures/materialize.py:295`, `eia-gen/src/eia_gen/services/figures/materialize.py:348`, `eia-gen/docs/26_drr_guideline_2025_05_plan_delta.md:41`, `eia-gen/docs/26_drr_guideline_2025_05_plan_delta.md:77`

  FPT:
  1) 문제 재정의: “참고도/개략도”로 제공된 도면(PDF/이미지)이 보고서에 그대로 삽입되면, 공신력 오해(허위기재) 리스크가 생기고 추후 보완요구 대응에서 근거/구분이 흔들린다.
  2) 가정(삭제 가능/검증 필요/필수 제약): (1) 참고도 여부는 자동 추정하지 않고 입력(FIGURES.source_origin/ gen_method)으로만 결정한다. (2) 워터마크는 “추정/생성”이 아니라 표기 강제다. (3) 이미지 품질/용량은 레시피(결정론)로 통제한다.
  3) 제1원칙 분해(Primitives): (a) REFERENCE 플래그 (b) 워터마크 텍스트/투명도/각도/폰트 (c) 삽입되는 실제 파일(derived) (d) QA 신호(규격/페이지/참고도 표기).
  4) 재조립(새 시스템 설계): `AUTHENTICITY:REFERENCE` 토큰을 gen_method에 결속하고, materialize 단계에서 워터마크를 적용한 derived PNG를 삽입/증빙화한다.
  5) 머스크 5단계 실행안(순서 고정):
     5-1) 요구사항 의심: “이미지 분석으로 공식/참고도 판별”은 금지(오판/허위기재 위험) → 입력 기반만 허용.
     5-2) 삭제: 사용자가 수동으로 워터마크를 넣는 운영을 기본값에서 제거(누락 위험).
     5-3) 단순화/최적화: 기존 FIGURES 컬럼(source_origin/gen_method)만으로 토큰을 생성하고 렌더러가 일관 적용한다.
     5-4) 가속: 유닛테스트로 REFERENCE 토큰/워터마크 적용을 회귀 고정한다.
     5-5) 자동화: QA에서 PDF 규격(A3 초과) 신호를 WARN로 노출해 조기 수정 루프를 만든다.
  6) 비용/Idiot Index(낭비/마찰): “참고도 표기 누락”으로 인한 보완요구/재제출 비용을 낮춘다.
  7) 리스크(인간/규제/2차효과/불확정): 폰트 환경 차이로 워터마크 렌더 결과가 달라질 수 있음 → env 폰트 경로로 결정론 확보(권장).
  8) 10배 옵션(2~4개): (A) FIGURES에 authenticity/usage_scope 컬럼을 정식 추가(P1-0/P3) (B) normalize 단계에서 패딩/테두리까지 함께 고정 (C) 삽입된 derived의 sha256/레시피를 usage로 기록(P4).
  9) 최종 권장안(+지표/검증 루프): REFERENCE 플래그가 있으면 워터마크 적용된 derived 파일이 삽입되고, QA에 참고도/규격 관련 WARN가 재현 가능하게 남는지로 검증한다.

- [x] P1-4. DOCX 삽입기(앵커 치환) 안정화 + 테스트 케이스 확장 (done-by: Codex)  
  - [x] 앵커 스캔/치환은 “문단 단위”로 구현되어 있음(단독 문단 앵커 원칙)  
  - [x] 표는 in-place fill(템플릿에 표 뼈대가 있으면) 지원  
  - [x] (선택) 표 셀 내부 앵커 스캔/치환(필요 시) (done-by: Codex)  
    refs: `eia-gen/src/eia_gen/services/docx/spec_renderer.py:1`, `eia-gen/tests/test_template_renderer.py:1`
  - [x] 중복 앵커 정책(초기 ERROR 권장) + 테스트 보강 (done-by: Codex)  
    refs: `eia-gen/src/eia_gen/services/docx/template_tools.py:1`, `eia-gen/src/eia_gen/cli.py:1`, `eia-gen/tests/test_template_tools.py:1`

  FPT:
  1) 문제 재정의: DOCX 템플릿에서 앵커가 “표 셀 내부”에 있거나 “중복”되면, 치환 대상이 모호해지고(누락/중복 삽입) 결과물이 샘플과 달라지거나 보완요구 대응에서 재현이 깨진다.
  2) 가정(삭제 가능/검증 필요/필수 제약): (1) 앵커는 단독 문단 텍스트로 존재한다(부분 문자열 치환은 범위 외). (2) 표/그림/섹션 앵커는 SSOT `spec/template_map.yaml`에 정의된다. (3) 중복 앵커는 “템플릿 오류”로 보고 조기 실패가 바람직하다.
  3) 제1원칙 분해(Primitives): (a) spec 앵커 목록 (b) 템플릿 내 앵커 위치(문서/표/셀) (c) 앵커 중복 여부 (d) 치환 로직(섹션/표/그림).
  4) 재조립(새 시스템 설계): 템플릿의 “모든 문단”(표/셀 포함)을 순회하여 앵커를 치환하고, template-check에서 중복 앵커를 에러로 노출해 수정 루프를 만든다.
  5) 머스크 5단계 실행안(순서 고정):
     5-1) 요구사항 의심: 표 셀 내부 치환은 필요한 범위(문단 단위)만 지원한다(복잡한 레이아웃 자동 생성은 금지).
     5-2) 삭제: “중복 앵커도 그냥 진행” 같은 암묵 동작을 제거하고, 검사 단계에서 실패하게 만든다.
     5-3) 단순화/최적화: 동일한 문단 순회 유틸로 렌더러/검사기 모두를 일관화한다.
     5-4) 가속: 유닛테스트로 (1) 셀 내부 앵커 치환 (2) 중복 앵커 감지를 회귀 고정한다.
     5-5) 자동화: quality gates의 template-check에 dup_spec/dup_template 지표를 포함해 CI/로컬에서 즉시 확인한다.
  6) 비용/Idiot Index(낭비/마찰): 템플릿 실수(앵커 중복/위치)로 인한 “다시 렌더→확인” 반복 비용을 낮춘다.
  7) 리스크(인간/규제/2차효과/불확정): 중복 앵커를 에러로 만들면 기존 템플릿이 깨질 수 있음 → template-check로 조기 발견/수정 유도(정상화).
  8) 10배 옵션(2~4개): (A) header/footer 앵커 지원(필요 시) (B) table anchor가 cell 안에 있어도 scaffold가 보강하도록 확장 (C) 렌더러 strict 모드에서 중복 앵커를 더 자세히 보고(위치/인덱스).
  9) 최종 권장안(+지표/검증 루프): template-check에서 dup_spec=0, dup_template=0이 유지되고, 표 셀 내부 앵커가 실제 치환되는 회귀 테스트가 통과하는지로 검증한다.

- [x] P1-5. `SSOT_PAGE_OVERRIDES` 운영 UX(페이지 찾기/미리보기/검증) 고도화 (done-by: B)  
  - [x] 시트/치환 로직 존재(샘플 페이지 → 첨부 PDF 페이지)  
  - [x] 스캔 PDF 페이지 제목 OCR 유틸 존재(`ocr_pdf_page_titles.py`)  
  - [x] (창원 케이스) `ATT-0008/ATT-0001` 페이지 치환 적용 사례가 존재  
  - [x] (B) “샘플 PDF의 빈/저밀도 슬롯” 후보 탐지 스크립트 추가(near-blank 페이지 빠른 식별) (done-by: B)  
  - [x] 키워드 검색→프리뷰→선택→`SSOT_PAGE_OVERRIDES` UPSERT(wizard) + QA 후처리 연계 (done-by: B)  
  - [x] (B) QA 후처리에서 sample_page 중복(WARN) 탐지 추가(치환 결과가 행 순서에 의존하는 리스크 노출) (done-by: B)  
  - [x] (B) QA 후처리에서 override_page 범위(WARN) 점검 추가(pdfinfo 기반; 페이지 범위 오류 조기 탐지) (done-by: B)  
  refs: `eia-gen/src/eia_gen/services/writer.py:556`, `eia-gen/scripts/ssot_page_override_wizard.py:1`, `eia-gen/scripts/find_sparse_pages_in_pdf.py:1`, `eia-gen/scripts/qa_ssot_overrides_summary.py:1`, `eia-gen/docs/20_user_manual_ko.md:210`, `eia-gen/docs/24_changwon_progress_report.md:39`

### P2 — 지번/좌표 기반 지도 자동생성(샘플 수준 “지도 계열”)

- [x] P2-0. 설정 SSOT 확정 + FIGURES `gen_method` 연결 규약 확정 (done-by: Codex)  
  - [x] 설정 파일은 이미 존재(`wms_layers.yaml`/`basemap.yaml`/`cache.yaml`)  
  - [x] FIGURES에서 `MAP_*` 선언 시 “완성 PNG” 생성 + `Case.assets.file_path` 갱신(케이스 `attachments/derived/figures/maps/`)  
  - [x] `gen_method` 규약(요약): `MAP_BASE|MAP_BUFFER|MAP_OVERLAY` + `size=WxH` + (선택) `zoom=` + `basemap_provider=`/`basemap_layer=` + (선택) `wms_layers=`/`rings=`  
  refs: `eia-gen/config/wms_layers.yaml:7`, `eia-gen/config/basemap.yaml:7`, `eia-gen/config/cache.yaml:16`, `eia-gen/src/eia_gen/services/figures/map_generate.py:1`, `eia-gen/src/eia_gen/cli.py:1`

  FPT:
  1) 문제 재정의: 위치도/영향권/규제중첩도 같은 “지도 계열” 그림은 반복 작업인데, 사람이 매번 타일/WMS 스크린샷을 만들면 재현성(요청 파라미터)·출처(source_id)·변경 이력(sha256)이 끊겨 제출/감사 대응이 약해진다.
  2) 가정(삭제 가능/검증 필요/필수 제약): (1) 지도 계열은 “공식 설계도서급 도면 생성”이 아니라 공공 WMTS/WMS 기반 참고/근거용 도면을 결정론적으로 조합한다. (2) basemap/WMS 설정은 SSOT(`config/*.yaml`)로 고정한다. (3) 네트워크/키 이슈가 있더라도 파이프라인은 best-effort로 진행하고 실패는 evidence/QA로 노출한다.
  3) 제1원칙 분해(Primitives): (a) 사업지 경계/중심점 입력 (b) basemap 타일 소스/캐시 (c) 오버레이 레이어 카탈로그 (d) 캡션/출처 연결 (e) 파생 산출물 위치/sha256 (f) 요청 메타(request_url/params) 기록.
  4) 재조립(새 시스템 설계): FIGURES.gen_method에 `MAP_*` 레시피를 선언하면, 케이스 폴더 `attachments/derived/figures/maps/`에 PNG를 생성하고 `derived_evidence_manifest`에 sha256/요청 메타를 기록해 source_register로 연결한다.
  5) 머스크 5단계 실행안(순서 고정):
     5-1) 요구사항 의심: “지도 이미지를 AI로 생성” 대신, 타일/WMS 조합으로만 만들고 레이어/출처가 SSOT로 설명 가능해야 한다.
     5-2) 삭제: 생성 결과를 output 폴더에 흩뿌리거나, XLSX를 렌더 단계에서 임의 수정하는 부작용을 제거한다(케이스 derived 고정).
     5-3) 단순화/최적화: 레시피 파라미터는 최소(`size/zoom/provider/layers/rings`)로 시작하고, 나머지는 config defaults로 통일한다.
     5-4) 가속: 네트워크 없는(mock) 유닛테스트로 레시피 파싱/PNG 생성/증빙 기록을 회귀 고정한다.
     5-5) 자동화: quality gates에서 생성된 maps PNG가 존재하고 derived evidence가 누락되지 않는지로 검증 루프를 만든다.
  6) 비용/Idiot Index(낭비/마찰): “지도 만들기→캡션/출처 붙이기→재수정” 반복 비용을 자동화로 줄인다.
  7) 리스크(인간/규제/2차효과/불확정): API 키 누락/서비스 장애로 배경지도가 비는 리스크 → 캐시+best-effort(placeholder)로 진행하되 note에 실패 원인을 남기고 QA에서 확인 가능하게 한다.
  8) 10배 옵션(2~4개): (A) 경계 기반 버퍼/영향권 계산 강화 (B) 좌표격자/스케일/도곽 스타일 SSOT화 (C) 레이어별 기간/버전 메타를 sources.yaml retrieval로 자동 연결.
  9) 최종 권장안(+지표/검증 루프): MAP_BASE/MAP_BUFFER/MAP_OVERLAY가 케이스 입력만으로 재현되고, source_register.xlsx에서 basemap/WMS 출처 및 request 메타가 누락되지 않는지로 검증한다.

- [x] P2-1. `MAP_BASE`(베이스맵+경계) 구현: WMTS/타일 + 경계 + 북쪽표시/축척/범례(기본) (done-by: Codex)  
  refs: `eia-gen/docs/09_figure_generation_pipeline.md:41`, `eia-gen/src/eia_gen/services/figures/map_generate.py:1`, `eia-gen/tests/test_map_generate.py:1`

- [x] P2-2. `MAP_BUFFER`(영향권 버퍼) 구현: 중심점+반경 또는 경계 기반 버퍼 + 라벨 규칙(300/500/1000m) (done-by: Codex)  
  refs: `eia-gen/docs/09_figure_generation_pipeline.md:46`, `eia-gen/src/eia_gen/services/figures/map_generate.py:1`, `eia-gen/tests/test_map_generate.py:1`

- [x] P2-3. `MAP_OVERLAY`(규제/위험 레이어 중첩) 구현: WMS 레이어 + 경계 중첩 + 캡션에 source_id/레이어명/요청메타 포함 (done-by: Codex)  
  refs: `eia-gen/docs/09_figure_generation_pipeline.md:51`, `eia-gen/docs/12_traceability_spec.md:65`, `eia-gen/src/eia_gen/services/figures/map_generate.py:1`, `eia-gen/src/eia_gen/services/data_requests/wms.py:1`, `eia-gen/tests/test_map_generate.py:1`

- [x] P2-4. GeoJSON 속성 누락/라벨 충돌 Fallback 규칙 반영(“완벽” 품질 보완 핵심) (done-by: Codex)  
  - [x] 경계 GeoJSON이 lon/lat(4326)로 보이면 EPSG 자동 감지(케이스 epsg와 불일치 시에도 안전)  
  - [x] 중심점 누락 시 boundary bbox 중심점으로 MAP_BUFFER 링 렌더 폴백  
  - [x] VWorld 키 미설정 시 `osm_tile`로 폴백(기본값 WMTS blank 방지)  
  - [x] `MAP_OVERLAY`는 레이어 title을 캡션에 보강(중복/충돌은 dedup)  
  refs: `eia-gen/src/eia_gen/services/figures/map_generate.py:1`, `eia-gen/tests/test_map_generate.py:1`

### P3 — QA/가드레일을 “도면/이미지/부록”까지 확장(허위기재 방지)

- [x] P3-1. “공식 도면 vs 참고도” 판별/표기 정책을 QA로 강제(워터마크/캡션/사용 제한) (done-by: Codex)  
  - `FIGURES.source_origin`을 드롭다운(OFFICIAL/REFERENCE/UNKNOWN)으로 고정해 운영 실수(오타/변형값) 감소  
  - QA에서 REFERENCE 표시와 `gen_method(AUTHENTICITY:REFERENCE)` 불일치 시 ERROR로 차단(워터마크 미적용 위험)  
  - QA에서 필수 그림이 REFERENCE인 경우 WARN으로 조기 노출(공식 도면 필요 가능성)  
  - 회귀 테스트로 “REFERENCE인데 AUTHENTICITY:OFFICIAL 강제” 케이스를 ERROR로 고정  
  refs: `eia-gen/src/eia_gen/services/xlsx/case_template_v2.py:1`, `eia-gen/src/eia_gen/services/qa/run.py:1`, `eia-gen/docs/08_qa_checklist_rules.md:1`, `eia-gen/tests/test_reference_watermark.py:1`, `eia-gen/docs/26_drr_guideline_2025_05_plan_delta.md:77`

  FPT:
  1) 문제 재정의: 공식 도면/사진이 없는 상태에서 “그럴듯한 도면”이 산출물에 섞이면 허위기재·감사·규제 리스크가 폭증한다(가드레일이 데이터/QA로 강제되지 않으면 운영 실수로 쉽게 발생).
  2) 가정(삭제 가능/검증 필요/필수 제약): (1) 참고도는 허용하되 반드시 워터마크/캡션으로 “공식 도면 아님”을 표시한다. (2) 분류는 사람 주의가 아니라 `FIGURES.source_origin` 같은 입력 필드로 강제한다. (3) 워터마크 적용 여부는 렌더러 동작(=gen_method 힌트)로 간접 검증 가능하다.
  3) 제1원칙 분해(Primitives): (a) 분류값(OFFICIAL/REFERENCE) (b) 캡션 표기(참고도) (c) 워터마크 적용 트리거(AUTHENTICITY:REFERENCE) (d) QA 룰(불일치 차단) (e) 회귀 테스트.
  4) 재조립(새 시스템 설계): 템플릿에서 분류값을 드롭다운으로 고정하고, QA가 “REFERENCE인데 워터마크 힌트 누락” 같은 위험 상태를 ERROR로 차단한다.
  5) 머스크 5단계 실행안(순서 고정):
     5-1) 요구사항 의심: “이미지 픽셀 분석으로 워터마크 존재 여부 판별” 대신, 트리거 힌트(gen_method) 정합성으로 먼저 안전장치를 둔다.
     5-2) 삭제: REFERENCE 표시인데 AUTHENTICITY:OFFICIAL 같은 위험한 조합을 허용하는 암묵 동작을 제거한다.
     5-3) 단순화/최적화: source_origin 드롭다운 + QA 1~2개 룰로 최소 정책을 잠근다.
     5-4) 가속: 유닛테스트로 mismatch 케이스를 회귀 고정한다.
     5-5) 자동화: quality gates에서 해당 ERROR가 0인지 확인해 릴리즈 전에 차단한다.
  6) 비용/Idiot Index(낭비/마찰): “이 도면이 공식인지 참고도인지”를 나중에 수동 검수로 찾아내는 비용을 QA 단계로 당긴다.
  7) 리스크(인간/규제/2차효과/불확정): 과도한 ERROR는 초반 샘플 생성에 마찰 → mismatch(워터마크 미적용 위험)만 ERROR로 두고, ‘필수 도면이 REFERENCE’는 WARN으로 시작한다.
  8) 10배 옵션(2~4개): (A) FIGURES에 authenticity/source_class/usage_scope/fallback_mode 컬럼을 추가해 정책을 더 정교하게 강제 (B) 참고도 사용 시 `DISPLAY_ONLY` 범위를 자동 체크.
  9) 최종 권장안(+지표/검증 루프): REFERENCE 그림이 항상 워터마크/캡션 표기(=정합 gen_method)를 갖추고, mismatch는 QA에서 ERROR로 차단되는지로 검증한다.

- [x] P3-2. “현장조사 표현” 가드레일을 첨부(ATTACHMENTS/FIELD_SURVEY_LOG) 메타와 더 강하게 결속 (done-by: Codex)  
  - [x] QA에서 “현장/현지조사/현장측정/탐문 결과/실시” 표현 탐지 + 현장 근거 출처 미인용 시 ERROR로 차단(`E-FIELD-001`~`E-FIELD-005`)  
  - [x] 위 문단에서 인용한 현장 근거 `src_id`는 `FIELD_SURVEY_LOG` 또는 `ATTACHMENTS` 메타로 연결되도록 WARN(`W-FIELD-META-001`)  
  - [x] QA에 “현장/탐문 표현 + FIELD_SURVEY_LOG 비어있음” WARN(진실 게이트, `W-FIELD-LOG-001`)  
  - [x] “현장/탐문 표현이 결론/요약 근거로 쓰인 경우”는 메타 연결을 제출 모드에서 ERROR로 승격(선택, `W-FIELD-META-002`) (done-by: Codex)  
  refs: `eia-gen/src/eia_gen/services/qa/run.py:574`, `eia-gen/docs/08_qa_checklist_rules.md:1`, `eia-gen/tests/test_field_survey_meta_gate.py:1`

  FPT:
  1) 문제 재정의: “현장조사 결과” 같은 표현이 문서에 등장해도, 그 근거(현장로그/사진/탐문기록)가 case.xlsx 메타로 연결되지 않으면 허위기재·감사·보완요구 리스크가 커진다.
  2) 가정(삭제 가능/검증 필요/필수 제약): (1) 현장 근거는 반드시 출처 인용으로 드러나야 한다. (2) 인용된 현장 근거는 최소 1개의 메타(현장로그 또는 첨부 증빙)로 연결되어야 한다. (3) 과도한 차단은 초기 작성 마찰이므로 단계적으로 강화한다.
  3) 제1원칙 분해(Primitives): (a) 위험 표현 탐지(정규식) (b) 문단 인용 출처 (c) sources.yaml의 source kind (d) FIELD_SURVEY_LOG/ATTACHMENTS 메타 (e) QA rule severity.
  4) 재조립(새 시스템 설계): (1) 표현이 나오면 현장 근거 출처 인용을 ERROR로 강제하고, (2) 인용된 src_id가 case.xlsx 메타로 연결되지 않으면 WARN으로 조기 노출한다.
  5) 머스크 5단계 실행안(순서 고정):
     5-1) 요구사항 의심: “나중에 수동 검수”로는 누락/과장을 막기 어렵다 → QA에서 자동 탐지한다.
     5-2) 삭제: 현장 근거 없이 “현장조사 결과”가 등장하는 경로를 제거한다(ERROR).
     5-3) 단순화/최적화: 메타 연결은 FIELD_SURVEY_LOG/ATTACHMENTS 중 하나면 우선 통과(WARN로 시작).
     5-4) 가속: 유닛테스트로 탐지/경고를 회귀 고정한다.
     5-5) 자동화: quality gates에서 해당 ERROR/WARN 추이를 확인해 운영 기준을 강화한다.
  6) 비용/Idiot Index(낭비/마찰): “근거 파일이 뭐였지?”를 찾는 사후 추적 비용을 QA 단계로 앞당겨 줄인다.
  7) 리스크(인간/규제/2차효과/불확정): 규칙이 너무 엄격하면 초기 케이스에서 ERROR가 늘 수 있음 → ‘근거 미인용’만 ERROR로 두고, ‘메타 미연결’은 WARN으로 시작한다.
  8) 10배 옵션(2~4개): (A) 제출 모드에서 `W-FIELD-META-001`을 ERROR로 승격 (B) 결론/요약 섹션에서만 더 강하게 적용 (C) ATTACHMENTS.evidence_type(사진/탐문기록) 기반 세분화.
  9) 최종 권장안(+지표/검증 루프): 현장 표현이 있을 때 (a) 현장 근거 출처가 인용되고 (b) case.xlsx 메타로 연결되는지가 QA 결과에 항상 남는지로 검증한다.

### P4 — 출처/증빙/사용처(Traceability) 스펙 “잠금”을 구현로 연결

- [x] P4-1. `sources.yaml`은 v2 포맷을 유지하면서 reliability/confidential/citation 등 확장 필드를 보존 (done-by: Codex)  
  refs: `eia-gen/docs/12_traceability_spec.md:22`, `eia-gen/src/eia_gen/models/sources.py:1`, `eia-gen/examples/sources.v2.sample.yaml:1`, `eia-gen/tests/test_models.py:1`

  FPT:
  1) 문제 재정의: `sources.yaml` v2는 top-level 메타(version/project)와 per-source 확장 필드(citation/retrieval 등)를 포함하는데, 로더가 이를 버리면 추적성/재현성이 약해진다.
  2) 가정(삭제 가능/검증 필요/필수 제약): (1) v1/v2 모두 허용한다. (2) 알 수 없는 필드는 삭제하지 않고 보존(model_extra)한다. (3) 기존 코드가 `.sources` 리스트만 사용해도 깨지지 않아야 한다.
  3) 제1원칙 분해(Primitives): (a) sources 리스트 (b) top-level 메타(version/project) (c) per-entry 확장 필드(딕트) (d) unique id 제약.
  4) 재조립(새 시스템 설계): SourceRegistry가 dict 입력을 그대로 보존하고(extra=allow), sources 키만 정규화하여 내부 `.sources`로 접근 가능하게 한다.
  5) 머스크 5단계 실행안(순서 고정):
     5-1) 요구사항 의심: “전 필드 스키마화” 대신, 우선 보존(allow)+최소 alias만 제공한다.
     5-2) 삭제: 로딩 시 버전/프로젝트 메타를 강제로 삭제하는 코드를 제거한다.
     5-3) 단순화/최적화: dict 입력은 그대로 반환하고, entries→sources만 정규화한다.
     5-4) 가속: v2 샘플 로드 테스트로 version/project/citation 보존을 고정한다.
     5-5) 자동화: quality gates(unit-tests)에서 회귀를 상시 검증한다.
  6) 비용/Idiot Index(낭비/마찰): “왜 source_register에 메타가 없지?” 같은 디버그 시간을 줄인다.
  7) 리스크(인간/규제/2차효과/불확정): extra 보존이 export에 불필요한 필드 누출을 유발할 수 있음 → export는 필요한 컬럼만 사용(현행 유지).
  8) 10배 옵션(2~4개): (A) v2 meta를 source_register SUMMARY로 노출(P4-3) (B) reliability/confidence 필드 표준화(P4-4).
  9) 최종 권장안(+지표/검증 루프): v2 sources.yaml 로드 후 model_extra에 version/project가 남고, 엔트리 model_extra에 citation이 남는지로 검증한다.

- [x] P4-2. evidence 저장/등록은 일부 구현됨(남은 과제: 통일/연결/확장) (done-by: Codex)  
  - [x] DATA_REQUESTS/WMS 결과는 `attachments/evidence/*` 저장 + `ATTACHMENTS` 등록까지 구현  
  - [x] figure materialize/판넬 산출물도 동일 규약(evidence_id, sha256, recipe, usage)으로 승격 (done-by: Codex)  
    refs: `eia-gen/src/eia_gen/services/figures/materialize.py:1`, `eia-gen/src/eia_gen/services/figures/photo_sheet_generate.py:1`, `eia-gen/src/eia_gen/services/figures/derived_evidence.py:1`
  - [x] `source_register.xlsx`의 Evidence/Claims로 자동 연결(사용처까지) (done-by: Codex)  
    refs: `eia-gen/src/eia_gen/services/export/source_register_xlsx.py:1`, `eia-gen/tests/test_source_register_xlsx.py:1`
  - [x] “DOCX에 실제 삽입된 파일 경로/페이지(선택)”를 usage로 기록(보완 요구 시 역추적 가능) (done-by: Codex)  
    refs: `eia-gen/src/eia_gen/services/docx/spec_renderer.py:1`, `eia-gen/tests/test_materialize_traceability.py:1`
  refs: `eia-gen/src/eia_gen/services/data_requests/runner.py:280`, `eia-gen/src/eia_gen/services/ingest_attachments.py:20`, `eia-gen/src/eia_gen/services/export/source_register_xlsx.py:266`

  FPT:
  1) 문제 재정의: 보고서에 실제 삽입된 파생 산출물(도면/사진/사진대지)이 “어떤 입력에서 어떤 레시피로 만들어졌는지(sha256/recipe)”가 남지 않으면, 보완요구·감사·재현에서 추적이 끊기고 허위기재 리스크 대응이 약해진다.
  2) 가정(삭제 가능/검증 필요/필수 제약): (1) 입력 XLSX를 in-place 수정하지 않는다. (2) 파생 산출물은 evidence_id로 식별되고, 관련 figure는 related_fig_id/used_in으로 연결한다. (3) 레시피는 결정론적 JSON으로 남긴다(환경 의존 최소화).
  3) 제1원칙 분해(Primitives): (a) 산출물 파일 경로 (b) out_sha256/src_sha256 (c) 레시피(target_dpi/max_width/format/워터마크/프레임) (d) 사용처(figure_id, doc_target, pdf_page) (e) source_ids.
  4) 재조립(새 시스템 설계): materialize가 recipe+sha256 메타를 반환하고, DOCX 렌더러가 그 메타를 derived_evidence_manifest에 기록 → source_register.xlsx의 Evidence/Claims로 자동 연결한다.
  5) 머스크 5단계 실행안(순서 고정):
     5-1) 요구사항 의심: “완전한 usage 레지스트리” 확장은 P4-3에서 다루고, 우선 figure 단위 연결을 고정한다.
     5-2) 삭제: 수동으로 “이미지 파일→증빙/사용처”를 적는 운영을 기본값에서 제거한다.
     5-3) 단순화/최적화: note는 JSON 1개(sha256+recipe+핵심 파라미터)로 표준화한다.
     5-4) 가속: 유닛테스트로 derived manifest가 sha256/asset_type을 포함하는지 회귀 고정한다.
     5-5) 자동화: quality gates에서 생성된 `source_register.xlsx`가 figure claim에 evidence_id를 채우는지로 검증한다.
  6) 비용/Idiot Index(낭비/마찰): “근거파일이 뭐였지?”를 찾는 수동 역추적 시간을 줄인다.
  7) 리스크(인간/규제/2차효과/불확정): note JSON이 커질 수 있음 → 최소 파라미터만 유지하고, 확장(Usage Register)은 별도 시트(P4-3)로 분리한다.
  8) 10배 옵션(2~4개): (A) materialize meta에 “truncated 이미지 처리 여부” 표기 (B) doc_target까지 usage로 표준화 (C) normalize 단계 산출물도 동일 규약으로 승격.
  9) 최종 권장안(+지표/검증 루프): derived_evidence_manifest(note)에 out_sha256/src_sha256/recipe가 남고, Claims(figure).evidence_id가 자동 채워지는지로 검증한다.

- [x] P4-3. `source_register.xlsx`를 4시트(SOURCE_CATALOG/EVIDENCE_INDEX/USAGE_REGISTER/VALIDATION_SUMMARY)로 확장(현행 시트 호환 유지) (done-by: Codex)  
  - canonical 4시트 생성 + 기존 `Source Register`/`Evidence Register`/`Claims` 시트는 유지(호환)  
  - `validation_report*.json` 내용을 `VALIDATION_SUMMARY` 시트로 요약(행 단위)  
  refs: `eia-gen/src/eia_gen/services/export/source_register_xlsx.py:1`, `eia-gen/src/eia_gen/cli.py:1`, `eia-gen/src/eia_gen/api.py:1`, `eia-gen/tests/test_source_register_xlsx.py:1`, `eia-gen/docs/12_traceability_spec.md:82`

  FPT:
  1) 문제 재정의: `source_register.xlsx`가 “출처/증빙/사용처/검증결과”를 한 화면에서 추적할 수 있는 정규화된 레지스트리(4시트) 형태가 아니면, 보완요구 대응·감사·재현에서 탐색 비용이 커지고 누락 리스크가 증가한다.
  2) 가정(삭제 가능/검증 필요/필수 제약): (1) 기존 시트(`Source Register`/`Evidence Register`/`Claims`)를 기대하는 운영/툴을 깨지 않는다. (2) 입력 XLSX를 in-place 수정하지 않는다. (3) QA 결과는 `validation_report*.json`이 SSOT이며 엑셀은 “요약 뷰”다.
  3) 제1원칙 분해(Primitives): (a) 출처 카탈로그(sources.yaml) (b) 증빙 인덱스(attachments/derived manifest) (c) 사용 레지스터(draft claims) (d) 검증 결과(QA RuleResult).
  4) 재조립(새 시스템 설계): export가 canonical 4시트를 생성하고, CLI/API가 QA 결과를 함께 전달해 `VALIDATION_SUMMARY`를 채운다(legacy 시트는 유지).
  5) 머스크 5단계 실행안(순서 고정):
     5-1) 요구사항 의심: “완전한 관계형 DB” 대신, openpyxl 기반 고정 4시트로 실무 점검 UX를 먼저 닫는다.
     5-2) 삭제: QA/증빙/사용처가 서로 다른 파일에 흩어져 사람이 수동으로 대조하는 운영을 기본값에서 제거한다.
     5-3) 단순화/최적화: VALIDATION_SUMMARY는 code/severity/message/path(related_anchor) 최소 셋만 먼저 고정한다.
     5-4) 가속: 유닛테스트로 4시트 생성/figure evidence 연결/validation 요약을 회귀 고정한다.
     5-5) 자동화: quality gates 산출물에 포함된 `source_register.xlsx`에서 누락 경고/오류를 즉시 필터링 가능하게 한다.
  6) 비용/Idiot Index(낭비/마찰): “근거/사용처/QA를 찾기 위해 여러 파일을 오가며 확인”하는 시간이 가장 큰 낭비이므로, 엑셀 요약 뷰로 탐색 비용을 줄인다.
  7) 리스크(인간/규제/2차효과/불확정): 시트가 늘어나면 혼란 가능 → legacy 시트는 유지하되, SSOT는 canonical 4시트로 문서화한다.
  8) 10배 옵션(2~4개): (A) `VALIDATION_SUMMARY`에 sheet/row_id까지 연결(P4-4) (B) `EVIDENCE_INDEX`에 sha256/recipe 컬럼 확정(P4-4) (C) `USAGE_REGISTER`를 문단/표/그림 외 추가 타입으로 확장.
  9) 최종 권장안(+지표/검증 루프): `source_register.xlsx`에 canonical 4시트가 존재하고, figure 사용행에 evidence_id가 채워지며, VALIDATION_SUMMARY에 QA 항목이 최소 1행 이상 기록되는지로 검증한다.

- [x] P4-4. “컬럼이 비면 어디서 가져오나” 안내를 규칙 파일로 고정하고 VALIDATION_SUMMARY로 노출 (done-by: Codex)  
  - `config/data_acquisition_rules.yaml`을 읽어 case.xlsx(v2)를 평가하고, `validation_report*.json`에 룰 결과를 추가(= 품질 게이트에서 조기 노출)  
  - `source_register.xlsx:VALIDATION_SUMMARY`에 related_sheet/related_row_id/related_anchor를 함께 기록(엑셀에서 즉시 필터링)  
  refs: `eia-gen/config/data_acquisition_rules.yaml:1`, `eia-gen/src/eia_gen/services/qa/run.py:1`, `eia-gen/src/eia_gen/services/qa/report.py:1`, `eia-gen/src/eia_gen/services/export/source_register_xlsx.py:1`, `eia-gen/tests/test_data_acquisition_rules.py:1`, `eia-gen/docs/12_traceability_spec.md:175`

  FPT:
  1) 문제 재정의: 케이스 입력에서 특정 시트/컬럼이 비어있을 때 “어디서 무엇을 확보해야 하는지”가 자동으로 안내되지 않으면, 작성/보완 루프가 느리고 누락이 반복된다.
  2) 가정(삭제 가능/검증 필요/필수 제약): (1) 안내 규칙은 코드 하드코딩이 아니라 YAML로 관리한다. (2) QA는 fail-fast가 아니라 “누락/요청 포인트를 조기에 노출”하는 역할도 수행한다. (3) 규칙 평가는 case.xlsx(v2)를 읽되, 실패 시 QA 전체는 fail-open으로 유지한다.
  3) 제1원칙 분해(Primitives): (a) 시트명/컬럼명 조건 (b) 빈칸/row_count 판정 (c) severity/message/fix_hint (d) 엑셀 요약(VALIDATION_SUMMARY)로의 노출.
  4) 재조립(새 시스템 설계): QA가 rules.yaml을 평가해 RuleResult를 생성하고, export가 related_sheet/row_id를 포함해 VALIDATION_SUMMARY에 기록한다.
  5) 머스크 5단계 실행안(순서 고정):
     5-1) 요구사항 의심: “완전한 자동 수집”보다 먼저, 누락 안내를 정확히 고정한다.
     5-2) 삭제: 누락 안내가 문서/구두로만 존재하는 운영을 제거한다.
     5-3) 단순화/최적화: sheet/column/row_count + 메시지/힌트만 최소 스키마로 시작한다.
     5-4) 가속: 유닛테스트로 규칙 평가 결과를 회귀 고정한다.
     5-5) 자동화: quality gates 산출물에서 VALIDATION_SUMMARY를 기반으로 보완요구 대응 체크리스트를 즉시 뽑는다.
  6) 비용/Idiot Index(낭비/마찰): “어떤 시트를 채워야 하지?”를 찾는 탐색 비용이 반복될수록 커지므로, 자동 안내로 상수시간화한다.
  7) 리스크(인간/규제/2차효과/불확정): 규칙이 과도하면 WARN이 늘어 노이즈가 될 수 있음 → severity/조건을 YAML에서 조절 가능하게 한다.
  8) 10배 옵션(2~4개): (A) suggested_sources를 DATA_REQUESTS 플랜과 직접 연결 (B) related_row_id를 sheet별 key로 고도화 (C) “대체 확보 경로”를 복수 제시.
  9) 최종 권장안(+지표/검증 루프): 빈칸 조건이 만족되면 validation_report에 규칙 결과가 추가되고, 동일 항목이 VALIDATION_SUMMARY에 나타나는지로 검증한다.

### P5 — DATA_REQUESTS(자동 수집) 하드닝/확장(기본 구현은 존재)

> 주의: “새 프로젝트 폴더 구조/신규 CLI”로 레포 구조를 깨는 제안은 채택하지 않음(현 v2 구조를 강화).  
> refs: `eia-gen/src/eia_gen/services/xlsx/case_reader.py:103`, `eia-gen/src/eia_gen/cli.py:144`

- [x] P5-1. (현황) v2 템플릿에 `DATA_REQUESTS` 포함 + 플래너/러너 동작. (과제) 트리거/증빙/반영 규칙을 SSOT로 잠금 (done-by: B)  
  - [x] 시트/플래너/러너 기본 동작  
  - [x] 케이스 간 재현(키 누락 시 disabled + 명확한 안내) 문서화/QA 연결 (done-by: B)  
  refs: `eia-gen/src/eia_gen/services/xlsx/case_template_v2.py:172`, `eia-gen/src/eia_gen/services/data_requests/planner.py:117`, `eia-gen/src/eia_gen/services/data_requests/runner.py:110`, `eia-gen/scripts/qa_data_requests_summary.py:1`, `eia-gen/docs/35_quality_gates_runbook.md:33`, `eia-gen/docs/13_data_requests_and_connectors_spec.md:29`

- [x] P5-2. Planner 트리거를 “현 v2 시트명”에 맞춰 잠금(빈칸 탐지 → DATA_REQUESTS 생성) (done-by: Codex)  
  refs: `eia-gen/docs/13_data_requests_and_connectors_spec.md:73`, `eia-gen/src/eia_gen/services/data_requests/planner.py:1`, `eia-gen/src/eia_gen/services/xlsx/case_template_v2.py:48`, `eia-gen/tests/test_data_requests.py:1`

  FPT:
  1) 문제 재정의: v2 템플릿 시트는 “행이 존재하지만 값이 비어있는” 경우가 흔한데, 단순 row_count만으로는 빈칸을 감지하지 못해 DATA_REQUESTS 자동 생성이 누락된다.
  2) 가정(삭제 가능/검증 필요/필수 제약): (1) Planner는 외부 API 호출 없이 “빈칸 탐지→요청 행 생성”만 한다. (2) 빈칸 판정은 시트별 핵심 컬럼(any_of) 기준으로 결정한다. (3) 기존 동작(진짜 데이터가 있으면 계획 생략)은 유지한다.
  3) 제1원칙 분해(Primitives): (a) 시트 rows(dict) (b) 핵심 컬럼 리스트(any_of) (c) 의미 있는 값 여부 (d) 생성될 req_id 목록.
  4) 재조립(새 시스템 설계): `_effective_empty(rows, any_of=...)`로 “실질적 빈 시트”를 판정하고, ENV_BASE_AIR/ZONING_BREAKDOWN 등 트리거를 결정론적으로 생성한다.
  5) 머스크 5단계 실행안(순서 고정):
     5-1) 요구사항 의심: “모든 시트 완전 자동” 대신, 먼저 핵심 2~3개 트리거만 고정한다.
     5-2) 삭제: row_count==0에만 의존하는 트리거를 제거한다.
     5-3) 단순화/최적화: any_of 컬럼만 보아 빈칸 여부를 판단한다(나머지는 무시).
     5-4) 가속: “빈 행 1개가 있어도 trigger 된다” 회귀 테스트를 추가한다.
     5-5) 자동화: unit-tests/quality gates로 Planner 회귀를 상시 검증한다.
  6) 비용/Idiot Index(낭비/마찰): 사용자가 DATA_REQUESTS를 수동으로 작성/수정하는 반복을 줄인다.
  7) 리스크(인간/규제/2차효과/불확정): any_of 선택이 부정확하면 과잉 트리거가 생길 수 있음 → 시트별 최소 핵심 컬럼만 사용.
  8) 10배 옵션(2~4개): (A) 트리거 테이블을 SSOT yaml로 외부화(P5-2 확장) (B) 빈칸 감지 결과를 QA에 요약(P5-3).
  9) 최종 권장안(+지표/검증 루프): ENV_BASE_AIR/ZONING_BREAKDOWN에 “빈 행만 있는” 경우에도 req_id가 생성되는지로 검증한다.

- [x] P5-3. Runner/Applier 규칙(우선순위/merge/evidence)을 확정하고 Traceability로 연결 (done-by: Codex)  
  - DATA_REQUESTS runner가 evidence 메타를 note(JSON)로 표준화(retrieved_at/request_url/request_params/hash_sha1, WMS bbox/srs 포함)
  - `source_register.xlsx:EVIDENCE_INDEX`가 note JSON을 파싱해 request_url/request_params/retrieved_at/hash_sha1 컬럼을 채움(typed-field dict도 지원)
  - 회귀 테스트 추가(WMS note JSON/legacy 파싱 + EVIDENCE_INDEX note JSON 추출)
  refs: `eia-gen/src/eia_gen/services/data_requests/runner.py:1`, `eia-gen/src/eia_gen/services/export/source_register_xlsx.py:1`, `eia-gen/tests/test_source_register_xlsx.py:1`, `eia-gen/tests/test_data_requests.py:1`, `eia-gen/docs/12_traceability_spec.md:65`

  FPT:
  1) 문제 재정의: DATA_REQUESTS가 생성한 증빙의 “요청 URL/파라미터/획득시각/해시”가 구조화되어 레지스트리로 올라오지 않으면, 감사·재현·보완요구 대응에서 역추적 비용이 사람에게 전가된다.
  2) 가정(삭제 가능/검증 필요/필수 제약): (1) 엑셀 입력(ATTACHMENTS.note)은 문자열이므로 구조화 메타는 JSON string으로 기록한다. (2) 레거시 note 형식(srs=... bbox=...)은 계속 파싱 지원한다. (3) export는 note 파싱 실패 시에도 fail-open(빈칸)으로 진행한다.
  3) 제1원칙 분해(Primitives): (a) 최소 메타(retrieved_at, request_url, request_params, hash) (b) v2 typed-field wrapper({"t":...}) (c) export 시 note JSON 파싱/정규화.
  4) 재조립(새 시스템 설계): runner가 note에 JSON 메타를 남기고, `source_register.xlsx:EVIDENCE_INDEX` export가 이를 파싱해 컬럼으로 materialize 한다.
  5) 머스크 5단계 실행안(순서 고정):
     5-1) 요구사항 의심: “완전한 DB/스키마” 대신, note JSON + export 파싱으로 실무 필요한 컬럼만 우선 닫는다.
     5-2) 삭제: request_url/params 누락으로 “파일만 남고 근거가 안 남는” 상태를 제거한다.
     5-3) 단순화/최적화: note JSON은 compact+stable(json sort_keys)로 저장해 diff/엑셀 가독성을 유지한다.
     5-4) 가속: 유닛테스트로 note JSON 파싱과 EVIDENCE_INDEX 컬럼 채움을 회귀 고정한다.
     5-5) 자동화: quality gates 산출물의 `source_register.xlsx`에서 즉시 필터/검색 가능하게 한다.
  6) 비용/Idiot Index(낭비/마찰): URL/파라미터 확인을 위해 코드/로그를 뒤지는 시간을 제거한다.
  7) 리스크(인간/규제/2차효과/불확정): note JSON이 커질 수 있음 → 최소 필드만 유지하고 확장은 별도 시트/manifest로 분리한다.
  8) 10배 옵션(2~4개): (A) sha256/recipe까지 확장 (B) normalize 단계 산출물도 동일 규약으로 승격.
  9) 최종 권장안(+지표/검증 루프): DATA_REQUESTS 증빙이 `EVIDENCE_INDEX`에서 request_url/request_params/hash/retrieved_at로 한 눈에 확인된다.

- [x] P5-4. KOSIS dataset_key/관측소 카탈로그를 SSOT로 고정(하드코딩 방지) (done-by: Codex)  
  - `DATA_REQUESTS(KOSIS)`가 `params_json.dataset_keys`를 `config/kosis_datasets.yaml`로 resolve(placeholder(TBD) 보호)  
  - 다중 dataset fetch 결과를 연도 기준으로 merge → 단일 evidence로 저장 + `ENV_BASE_SOCIO`를 한 번에 채움  
  - Planner가 `ENV_BASE_SOCIO` 빈칸 감지 시 dataset_keys 기반 disabled 가이드 행(최근 5년)을 생성  
  - (관측소) `config/stations/kma_asos_stations.csv` 경로 고정 + planner/runner auto-pick(없으면 온라인 fetch fallback)  
  refs: `eia-gen/config/kosis_datasets.yaml:1`, `eia-gen/src/eia_gen/services/data_requests/kosis.py:1`, `eia-gen/src/eia_gen/services/data_requests/planner.py:1`, `eia-gen/src/eia_gen/services/data_requests/runner.py:1`, `eia-gen/src/eia_gen/cli.py:2056`, `eia-gen/config/stations/README.md:1`, `eia-gen/tests/test_data_requests.py:1`

  FPT:
  1) 문제 재정의: 케이스마다 KOSIS 호출 파라미터(orgId/tblId/항목코드/기간)를 엑셀에 직접 하드코딩하면 재사용/감사/수정 비용이 폭증하고, 잘못된 테이블/코드로 인한 “조용한 오염” 리스크가 커진다.
  2) 가정(삭제 가능/검증 필요/필수 제약): (1) dataset_key는 프로젝트(SSOT)에서만 정의하고, 케이스는 dataset_key와 admin_code/기간만 제공한다. (2) `kosis_datasets.yaml`의 placeholder(TBD)는 실행을 막아 잘못된 호출을 예방한다. (3) 기존(레거시) `query_params+mappings` 케이스 입력도 호환 유지한다.
  3) 제1원칙 분해(Primitives): (a) dataset_key→query_params 템플릿 (b) 템플릿 변수(admin_code/start_year/end_year) (c) 항목 매핑(mappings) (d) 연도별 row merge (e) 단일 evidence 저장.
  4) 재조립(새 시스템 설계): runner가 dataset_keys를 SSOT config로 resolve하고, 각 dataset을 fetch→부분 row 생성→연도 키로 merge한 뒤 단일 evidence + 시트 반영으로 닫는다.
  5) 머스크 5단계 실행안(순서 고정):
     5-1) 요구사항 의심: “완전 자동 카탈로그” 대신, dataset_key/템플릿/매핑만 SSOT로 고정해 하드코딩을 제거한다.
     5-2) 삭제: 케이스 엑셀에 orgId/tblId/itmId를 직접 적는 운영을 기본값에서 제거한다.
     5-3) 단순화/최적화: query_params는 템플릿 치환만 지원하고, placeholder(TBD) 검증으로 안전장치를 둔다.
     5-4) 가속: 유닛테스트로 dataset_keys resolve/merge 동작을 회귀 고정한다.
     5-5) 자동화: `source_register.xlsx:EVIDENCE_INDEX`에서 KOSIS 요청 메타를 즉시 확인 가능하게 한다(P5-3 연계).
  6) 비용/Idiot Index(낭비/마찰): 매번 KOSIS 테이블/코드를 찾아 채우는 반복과 검증 비용을 줄인다.
  7) 리스크(인간/규제/2차효과/불확정): 잘못된 dataset 매핑은 보고서 수치 오염 위험 → placeholder 차단 + evidence 저장으로 재현/감사 가능하게 한다.
  8) 10배 옵션(2~4개): (A) dataset_key별 “공식 출처ID(src_id)”도 SSOT로 강제 (B) admin_code 후보 자동 추천(주소→행정코드) 연결.
  9) 최종 권장안(+지표/검증 루프): 케이스는 dataset_keys/admin_code/기간만 입력해도 ENV_BASE_SOCIO가 채워지고, evidence가 저장/추적되며, config placeholder가 남아있으면 실행이 명확히 실패해야 한다.

### P6 — Word 후처리/자동화(로컬 Word 또는 Antigravity)

- [x] P6-1. “python-docx로 DOCX 산출 → 필요 시 로컬 후처리(TOC/필드 업데이트/PDF)” 단계화를 SSOT로 고정 (done-by: B)  
  - [x] 기본 산출물은 `report_*.docx`까지(Word 의존 최소화) (done-by: B)  
  - [x] DOCX→PDF 후처리 스크립트 추가: macOS `Word(설치 시)→Pages→LibreOffice(soffice)` 자동 선택 + 명시 모드 지원 (done-by: B)  
  - [x] (macOS) Pages 기반 DOCX→PDF 변환 스모크 성공(Word 미설치 환경에서도 동작) (done-by: B)  
  - SSOT 문서화: `eia-gen/docs/39_docx_postprocess_pipeline.md` (done-by: B)  
  refs: `eia-gen/scripts/postprocess_docx_to_pdf.py:1`, `eia-gen/docs/20_user_manual_ko.md:512`, `eia-gen/docs/33_antigravity_docx_to_pdf_runbook.md:1`, `eia-gen/scripts/doctor_env.py:394`, `eia-gen/docs/39_docx_postprocess_pipeline.md:1`

- [ ] P6-2. (macOS) Word 설치 환경에서 “열기→필드/목차 업데이트→저장→PDF” E2E 확인(선택, NEED_CONFIRM)  
  refs: `eia-gen/scripts/postprocess_docx_to_pdf.py:64`, `eia-gen/docs/39_docx_postprocess_pipeline.md:36`
  - [x] (B) (증적) 현재 로컬 환경에서 Word 미설치 상태를 doctor 결과로 기록(=P6-2는 환경 의존) (done-by: B)  
    refs: `eia-gen/output/_tmp/doctor_env_2026-01-03_venv.txt:1`, `eia-gen/output/_tmp/doctor_env_2026-01-03_anaconda.txt:1`, `eia-gen/scripts/doctor_env.py:1`

- [x] P6-3. (fallback2) 로컬 후처리 불가 시 “Antigravity로 DOCX→PDF 후처리” 운영 절차/보안/실패 메시지 SSOT 확정(선택) (done-by: Codex)  
  - [x] Runbook(입출력 계약/보안/실패 처리) 초안 작성 (done-by: B)  
  - [x] (B) “파일 기반 Job 계약”(폴더 구조 + job.json) + 큐 생성 스크립트 추가(원격 러너는 환경 의존) (done-by: B)  
    - 러너 측 실행기(로컬/원격): job.json을 읽고 `postprocess_docx_to_pdf.py`를 호출하여 out/pdf + log/{conversion_log,done}.json 생성 (done-by: Codex)  
    - (Codex) (증적) 로컬 러너 스모크: Pages backend로 job 처리 성공(out/pdf + log/{conversion_log,done}.json 생성) (done-by: Codex)  
    refs: `eia-gen/docs/46_antigravity_job_contract.md:1`, `eia-gen/scripts/antigravity_queue_docx_to_pdf.py:1`, `eia-gen/scripts/antigravity_run_docx_to_pdf.py:1`, `eia-gen/docs/33_antigravity_docx_to_pdf_runbook.md:30`
  - [x] Antigravity 접속 방식/권한/보관·삭제 정책 확정(프로젝트 기본값으로 SSOT 고정) (done-by: Codex)  
    - 정책 템플릿/DoD: `eia-gen/docs/33_antigravity_docx_to_pdf_runbook.md`의 3.1.1(기본값) + 7(완료 기준)  
  refs: `eia-gen/docs/33_antigravity_docx_to_pdf_runbook.md:1`, `eia-gen/docs/39_docx_postprocess_pipeline.md:58`

---

### P7 — 케이스 입력 추출/정합(관광농원, placeholder 최소화)

> 목적: “허위기재 없이” `【작성자 기입 필요】`를 줄이기 위해, 스캔 PDF를 목차/장절 단위로 읽고
> `case.xlsx(v2)`에 **근거가 있는 값만** 구조화하여 입력한다.

- [x] P7-1. `case_changwon_2025` 기준 “전량 읽기/추출” 체크리스트를 SSOT로 고정하고 진행(2026-01)  
  refs: `eia-gen/docs/53_case_changwon_2025_pdf_extraction_plan.md:1`
- [x] P7-2. `case.xlsx:LOCATION/PARCELS/PROJECT/ZONING_BREAKDOWN`을 PDF 근거로 정합(우선순위 P0)  
  refs: `eia-gen/docs/53_case_changwon_2025_pdf_extraction_plan.md:55`
- [x] P7-3. 표 2.3-1(환경관련 지구·지역) 기반으로 `ZONING_OVERLAY`를 근거/출처까지 정리(P1)  
  refs: `eia-gen/docs/53_case_changwon_2025_pdf_extraction_plan.md:81`
- [x] P7-4. QA(스코프/전체)로 placeholder/WARN 감소를 계량하고 next-actions를 갱신  
  refs: `eia-gen/docs/35_quality_gates_runbook.md:1`

---

## 4) 완료 기준(DoD) — “새 프로젝트에도 샘플 수준으로 나온다”의 정의

- [x] (이미지) 모든 PDF/사진 기반 그림이 “정규화 레시피(회전/크롭/패딩/테두리/워터마크)”를 거쳐 제출용 품질로 삽입된다.  
  refs: `eia-gen/config/asset_normalization.yaml:1`, `eia-gen/src/eia_gen/services/docx/spec_renderer.py:809`, `eia-gen/src/eia_gen/services/figures/materialize.py:633`
- [x] (도면 규격) 도면은 A3 이하 권장 등 제출 규격(폭/해상도/가독성)이 레시피/QA로 통제된다.  
  refs: `eia-gen/src/eia_gen/services/qa/run.py:854`, `eia-gen/docs/26_drr_guideline_2025_05_plan_delta.md:41`
- [x] (누락 처리) PDF 페이지 지정 누락/파일 누락 시 “빈칸”이 아니라 placeholder(사유+필요 입력+액션)가 문서에 삽입되고 QA에 ERROR/WARN으로 남는다.  
  refs: `eia-gen/src/eia_gen/services/docx/spec_renderer.py:862`, `eia-gen/src/eia_gen/services/qa/run.py:831`
- [x] (EIA) 필수 그림이 “첨부 or 자동 생성 or placeholder” 중 하나로 반드시 채워진다.  
  refs: `eia-gen/src/eia_gen/services/docx/spec_renderer.py:862`, `eia-gen/src/eia_gen/services/qa/run.py:801`
- [x] (DIA) 최소 1개 부록/표준 서식(예: 주민탐문/유지관리대장)이 자동 생성 성공해야 통과(제출물 세트 체감). (done-by: Codex)  
  refs: `eia-gen/spec_dia/table_specs.yaml:139`, `eia-gen/spec_dia/table_specs.yaml:253`, `eia-gen/docs/26_drr_guideline_2025_05_plan_delta.md:110`, `eia-gen/src/eia_gen/services/dia/auto_generate.py:1`, `eia-gen/src/eia_gen/cli.py:404`, `eia-gen/src/eia_gen/services/writer.py:825`, `eia-gen/tests/test_dia_auto_generate.py:1`
- [x] (추적성) 어떤 그림/도면/사진이든 `source_register.xlsx`에서 “출처ID + 증빙(evidence) + 사용처(anchor)”가 추적된다(생성 시점 기록).  
  refs: `eia-gen/src/eia_gen/services/export/source_register_xlsx.py:266`, `eia-gen/src/eia_gen/services/docx/spec_renderer.py:815`, `eia-gen/docs/12_traceability_spec.md:65`

---

## 5) 다음 진행(추천 스프린트) — “샘플 체감(도면/사진) + 추적성”부터 닫기

- [x] (플랜 반영 완료, done-by: B) 완성도/편의성/효율성 약점→해결방안을 P-1/P1/P4/P6 체크리스트로 분해하여 반영  
  refs: `eia-gen/docs/11_execution_plan.md:86`, `eia-gen/docs/11_execution_plan.md:150`, `eia-gen/docs/11_execution_plan.md:194`, `eia-gen/docs/11_execution_plan.md:276`, `eia-gen/docs/11_execution_plan.md:310`

- [x] Sprint-1A(편의성, owner: B) `SSOT_PAGE_OVERRIDES` “검색→프리뷰→검증” UX를 CLI로 최소 구현(사용자 시행착오 감소) (done-by: B)  
  - [x] OCR(페이지 제목/키워드) 결과에서 후보 페이지 Top-N 추천 + 프리뷰 PNG 생성 (done-by: B)  
  - [x] 선택 결과를 `case.xlsx:SSOT_PAGE_OVERRIDES`에 자동 추가(UPSERT by sample_page) (done-by: B)  
  - [x] QA 리포트에 “치환된 페이지 목록/출처(src_id)/폭(width_mm)” 요약을 남김(후처리 스크립트로 append) (done-by: B)  
  refs: `eia-gen/scripts/qa_ssot_overrides_summary.py:1`, `eia-gen/scripts/ssot_page_override_wizard.py:1`, `eia-gen/scripts/ocr_pdf_page_titles.py:91`, `eia-gen/src/eia_gen/services/writer.py:556`, `eia-gen/src/eia_gen/services/xlsx/case_template_v2.py:158`, `eia-gen/docs/24_changwon_progress_report.md:39`

- [x] Sprint-1B(완성도+효율성, owner: A) “정규화 PNG(도면/사진) + 증빙 + DOCX 폭 반영”을 한 번에 end-to-end로 닫기 (done-by: Codex)  
  - [x] Normalize 단계(패딩/테두리/워터마크/표준 DPI)를 SSOT로 정의하고, 산출물을 `attachments/derived`로 통일 저장  
  - [x] sha256+recipe 기반 캐시 키로 전환(동일 입력이면 동일 산출물 재사용)  
  - [x] DOCX 삽입은 `attachments/derived/figures/_materialized`를 SSOT 산출물로 사용(미존재 시 생성, 존재 시 캐시 재사용)  
  - [x] evidence_id/sha256/선택된 PDF page(휴리스틱 포함)를 `source_register.xlsx` Claims/Usage에 기록  
  refs: `eia-gen/config/asset_normalization.yaml:1`, `eia-gen/src/eia_gen/services/figures/materialize.py:262`, `eia-gen/src/eia_gen/services/docx/spec_renderer.py:448`, `eia-gen/src/eia_gen/services/export/source_register_xlsx.py:266`, `eia-gen/tests/test_materialize_policy.py:1`, `eia-gen/docs/12_traceability_spec.md:65`

- [x] Sprint-2(완성도, owner: A) `CALLOUT_COMPOSITE`(사진대지) 2/4/6 그리드 고정 구현 + 입력 연결 규약 확정 (done-by: Codex)  
  - [x] 입력 묶음 규약(권장 1안): `ATTACHMENTS.related_fig_id == FIGURES.fig_id`로 사진 그룹 정의(순서/캡션 포함)  
  - [x] 결과 PNG를 증빙화(evidence_id/sha256/recipe)하고 DOCX에 삽입(폭 clamp 적용)  
  - [x] 골든 테스트: 동일 입력/동일 레시피 → 동일 sha256  
  refs: `eia-gen/docs/09_figure_generation_pipeline.md:61`, `eia-gen/src/eia_gen/services/figures/photo_sheet_generate.py:86`, `eia-gen/src/eia_gen/services/figures/callout_composite.py:151`, `eia-gen/src/eia_gen/cli.py:424`, `eia-gen/tests/test_photo_sheet_generate.py:1`, `eia-gen/tests/test_callout_composite_determinism.py:1`

- [x] Sprint-2(편의성, owner: B) “doctor(환경 점검)” + “다음 액션 출력”으로 사용자 셀프 해결률을 올리기 (done-by: B)  
  - [x] PyMuPDF/Word/LibreOffice/폰트/키 유무를 한 번에 점검하고, 누락 시 대체 경로(스킵/placeholder) 안내 (done-by: B)  
  - [x] QA/CLI 에러에서 “어느 시트의 어떤 컬럼을 채워야 하는지”를 직접 출력(후처리 스크립트) (done-by: B)  
  refs: `eia-gen/scripts/qa_next_actions.py:1`, `eia-gen/scripts/doctor_env.py:1`, `eia-gen/docs/11_execution_plan.md:166`, `eia-gen/src/eia_gen/cli.py:1302`, `eia-gen/src/eia_gen/services/qa/run.py:1`

- [x] Sprint-3(제출물 세트, owner: A) DIA “부록/별지 서식”을 케이스 입력→부록 표 자동 생성까지 잠금(실무 제출 체감 강화) (done-by: Codex)  
  - [x] 주민탐문/이행관리대장 등 표준 서식을 “입력 시트(로그)→부록 표”로 연결(추적 포함)  
  refs: `eia-gen/src/eia_gen/services/xlsx/case_template_v2.py:419`, `eia-gen/src/eia_gen/services/xlsx/case_reader_v2.py:1164`, `eia-gen/spec_dia/template_map.yaml:1`, `eia-gen/spec_dia/table_specs.yaml:139`, `eia-gen/spec_dia/table_specs.yaml:253`, `eia-gen/src/eia_gen/services/dia/auto_generate.py:1`

---

## 6) 협업/병렬 진행(Agent A/B) — 충돌 없이 동시에 진행하는 방법

- [x] (현황) Agent B는 “플랜/문서/분업” 영역에서 병렬 기여 가능(코어 구현 파일 충돌 회피)  
  refs: `eia-gen/docs/11_execution_plan.md:337`
- [x] (B) (운영) 체크리스트 진척률 자동 계산 스크립트 추가(문서 표기 드리프트 방지) (done-by: B)  
  - [x] (B) `\\1{pct}` backref 버그(`\\154.8` 오염) 수정 + 오염된 문구 복구(재발 방지) (done-by: B)  
  refs: `eia-gen/scripts/plan_progress.py:1`, `eia-gen/docs/11_execution_plan.md:94`, `eia-gen/docs/30_handoff_agent_b_to_a.md:15`
- [x] (B) (운영) 플랜/런북의 `path:line` refs 드리프트를 조기 탐지하는 검증 스크립트 추가(문서 유지보수 비용↓) (done-by: B)  
  - [x] 기본 대상: `docs/11_execution_plan.md`(SSOT)  
  - [x] (선택) `case.xlsx:SHEET#rowN` 형태의 refs도 시트/row 범위까지 검증(`--check-xlsx-rows`)  
  refs: `eia-gen/scripts/verify_md_refs.py:1`, `eia-gen/docs/11_execution_plan.md:1`
- [x] (필수) 작업 단위/소유자(Owner) 선언: Sprint/P 항목에 `owner: A|B`를 붙이고, 완료 시 `done-by: A|B`로 기록 (done-by: B)  
  refs: `eia-gen/docs/11_execution_plan.md:342`
- [x] (필수) 충돌 방지 규칙: “한 파일 한 에이전트” 원칙 + 공용 파일은 선점/락(문서로 기록) 후 수정 (done-by: B)  
  refs: `eia-gen/docs/27_parallel_working_protocol.md:1`
- [x] (권장) Git 브랜치 분리: `agent-a/*`, `agent-b/*`로 작업하고 rebase/merge로 합치기(동시 편집 충돌 최소화) — 로컬-only(노커밋) 운용으로 N/A 처리 (done-by: B)  
  refs: `eia-gen/docs/27_parallel_working_protocol.md:48`
