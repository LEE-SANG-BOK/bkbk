from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class SectionSpec:
    section_id: str
    title: str
    requirements: str


SECTION_SPECS: list[SectionSpec] = [
    SectionSpec(
        section_id="CH0_SUMMARY",
        title="요약",
        requirements=(
            "- 대상판정(근거) 요약\n"
            "- 사업규모/주요시설 요약\n"
            "- 핵심 이슈(토사유출·경관·소음 등)\n"
            "- 핵심 저감대책(침사지, 살수, 방음, 친환경배수 등)\n"
            "- 값/출처 없으면 【작성자 기입 필요】로 남김"
        ),
    ),
    SectionSpec(
        section_id="CH1_OVERVIEW",
        title="제1장 사업의 개요",
        requirements=(
            "- 목적 및 필요성\n"
            "- 위치 및 면적(지번별 면적표는 별도 표)\n"
            "- 사업 내용 및 규모(시설별)\n"
            "- 추진 일정(공정표)\n"
            "- 관련 계획/상위계획 정합성(가능 시)\n"
            "- 단정 금지, 입력 기반 서술"
        ),
    ),
    SectionSpec(
        section_id="CH1_PERMITS",
        title="인허가 현황 및 협의 대상 근거",
        requirements=(
            "- 소규모환경영향평가 대상근거(시행령 별표 등)\n"
            "- 재해영향성 검토 연계 여부(해당 시)\n"
            "- 선행평가 존재 시 제60조 생략 적용 근거"
        ),
    ),
    SectionSpec(
        section_id="CH2_METHOD",
        title="제2장 환경현황 조사 - 조사 범위·방법",
        requirements=(
            "- 조사범위(사업지+영향권 반경)\n"
            "- 조사방법(문헌/공공DB/현장)\n"
            "- 현장조사표/증빙은 부록 연계(값 없으면 【작성자 기입 필요】)"
        ),
    ),
    SectionSpec(
        section_id="CH2_NAT_TG",
        title="2.2 자연환경 현황 - 지형·지질",
        requirements=(
            "- 지형/경사/지질 개요\n"
            "- 입력 수치/지도 기반 요약\n"
            "- 출처/적용범위 명확히"
        ),
    ),
    SectionSpec(
        section_id="CH2_NAT_ECO",
        title="2.2 자연환경 현황 - 동·식물상(자연생태)",
        requirements=(
            "- 현장조사 요약(일시/방법/조사경로/주요 출현종)\n"
            "- 문헌조사 요약(최근자료/출처)\n"
            "- 보호종/서식지 관련 주의사항\n"
            "- 입력 없으면 【작성자 기입 필요】"
        ),
    ),
    SectionSpec(
        section_id="CH2_NAT_WATER",
        title="2.2 자연환경 현황 - 수환경",
        requirements=(
            "- 인근 하천/수계 및 거리\n"
            "- 수질현황(측정망/자료 기간/지점)\n"
            "- 지하수/비점유출 경로(가능 시)"
        ),
    ),
    SectionSpec(
        section_id="CH2_LIFE_AIR",
        title="2.3 생활환경 현황 - 대기질",
        requirements=(
            "- 기존자료(측정소/기간/거리) 기반 현황\n"
            "- 공사·운영 영향예측과는 3장에서 연결"
        ),
    ),
    SectionSpec(
        section_id="CH2_LIFE_NOISE",
        title="2.3 생활환경 현황 - 소음·진동",
        requirements=(
            "- 측정 또는 문헌 기반 현황\n"
            "- 민감수용체(주거 등) 거리\n"
            "- 입력 없으면 【작성자 기입 필요】"
        ),
    ),
    SectionSpec(
        section_id="CH2_LIFE_ODOR",
        title="2.3 생활환경 현황 - 악취",
        requirements=(
            "- 주변 발생원(축사/음식점 등) 확인\n"
            "- 풍향/거리 고려 정성평가(입력 기반)\n"
        ),
    ),
    SectionSpec(
        section_id="CH2_SOC_LANDUSE",
        title="2.4 사회·경제환경 현황 - 토지이용",
        requirements=(
            "- 토지이용계획/중첩규제 확인\n"
            "- 항공/드론사진(이격거리 표시) 첨부 여부 명시\n"
            "- 기존자료 중심 + 최근자료 사용 + 출처 기재"
        ),
    ),
    SectionSpec(
        section_id="CH2_SOC_LANDSCAPE",
        title="2.4 사회·경제환경 현황 - 경관",
        requirements=(
            "- 조망점 선정 사유/목록\n"
            "- 현장사진 및 시뮬레이션(가능 시)\n"
            "- 입력 없으면 【작성자 기입 필요】"
        ),
    ),
    SectionSpec(
        section_id="CH2_SOC_POP",
        title="2.4 사회·경제환경 현황 - 인구·주거",
        requirements=(
            "- 주변 마을/주거 분포 및 거리\n"
            "- 민원 가능성 등 정성평가(입력 기반)"
        ),
    ),
    SectionSpec(
        section_id="CH3_CONSTRUCTION",
        title="제3장 환경영향 예측 - 공사 단계",
        requirements=(
            "- 비산먼지/소음·진동/토사유출/생태 등\n"
            "- 현황→영향→저감→잔류영향/관리 흐름\n"
            "- '문제 없음' 같은 단정 금지\n"
            "- 정량값 없으면 추정 금지(필요자료를 TODO로)"
        ),
    ),
    SectionSpec(
        section_id="CH3_OPERATION",
        title="제3장 환경영향 예측 - 운영 단계",
        requirements=(
            "- 교통/오수·비점/폐기물/이용객 등\n"
            "- 현황→영향→저감→잔류영향/관리 흐름\n"
            "- 정량값 없으면 추정 금지"
        ),
    ),
    SectionSpec(
        section_id="CH4_TEXT",
        title="제4장 환경보전 및 저감방안",
        requirements=(
            "- 자연환경/생활환경 저감대책 요약\n"
            "- 공사/운영 관리방안 구분\n"
            "- 저감대책(표)과 일관되게 서술"
        ),
    ),
    SectionSpec(
        section_id="CH5_TEXT",
        title="제5장 환경관리 계획(협의조건 이행관리)",
        requirements=(
            "- 협의조건 이행관리대장 중심\n"
            "- 소규모 특성상 책임자/사후조사는 조건부 섹션"
        ),
    ),
]

