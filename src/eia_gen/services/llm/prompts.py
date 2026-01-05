from __future__ import annotations

import json

from eia_gen.services.sections import SectionSpec


SYSTEM_PROMPT_KO = """\
역할: 당신은 한국의 소규모환경영향평가서(관광농원) 작성 실무자이다.
목표: 제공된 FACTS_JSON과 요구사항만으로 해당 섹션의 ‘보고서 초안’을 작성한다.

절대 규칙:
1) FACTS_JSON에 없는 사실/수치/지명/법령문구/결론을 새로 만들지 말 것.
2) 값이 비어 있으면 추정하지 말고 반드시 "【작성자 기입 필요】"로 남길 것.
3) 모든 문단은 끝에 출처를 〔SRC:ID〕 형식으로 1개 이상 포함할 것.
   - 출처가 없으면 〔SRC-TBD〕로 표기하고, todos에 누락 사유/필요자료를 기록할 것.
4) "문제 없음", "영향 없음" 같은 단정 표현을 금지한다.
   - 영향이 작거나 기준 만족이 예상되더라도 전제(저감대책 적용 등)를 명시할 것.
5) 한국어 보고서 문체(객관/공식)로 작성한다.

출력 형식:
- 반드시 JSON만 출력한다(설명문, 마크다운 금지).
- 스키마:
  {"section_id": "...", "title": "...", "paragraphs": [...], "tables": [], "figures": [], "todos": []}
"""


def build_user_prompt(spec: SectionSpec, facts: dict) -> str:
    facts_json = json.dumps(facts, ensure_ascii=False, indent=2)
    return (
        "다음 섹션을 작성하라.\n\n"
        f"- section_id: {spec.section_id}\n"
        f"- title: {spec.title}\n"
        "- requirements:\n"
        f"{spec.requirements}\n\n"
        "FACTS_JSON:\n"
        f"{facts_json}\n"
    )
