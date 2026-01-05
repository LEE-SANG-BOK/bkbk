from __future__ import annotations

import json
from typing import Any

import httpx

from eia_gen.services.draft import SectionDraft
from eia_gen.services.llm.base import LLMClient
from eia_gen.services.llm.prompt_library import PromptLibrary
from eia_gen.services.llm.prompts import SYSTEM_PROMPT_KO, build_user_prompt
from eia_gen.services.sections import SectionSpec


def _extract_json(text: str) -> str:
    # Allow accidental code fences.
    t = text.strip()
    if t.startswith("```"):
        # remove first fence line and last fence line
        lines = t.splitlines()
        # drop first line
        lines = lines[1:]
        # drop last line if fence
        if lines and lines[-1].strip().startswith("```"):
            lines = lines[:-1]
        t = "\n".join(lines).strip()

    # Try to find the first {...} block.
    start = t.find("{")
    end = t.rfind("}")
    if start != -1 and end != -1 and end > start:
        return t[start : end + 1]
    return t


class OpenAIChatClient(LLMClient):
    def __init__(self, api_key: str, model: str, prompt_root: str = "prompts") -> None:
        self._api_key = api_key
        self._model = model
        self._prompts = PromptLibrary(prompt_root)

    def generate_section(self, spec: SectionSpec, facts: dict) -> SectionDraft:
        url = "https://api.openai.com/v1/chat/completions"
        headers = {
            "Authorization": f"Bearer {self._api_key}",
            "Content-Type": "application/json",
        }
        system_prompt = self._prompts.system_prompt() or SYSTEM_PROMPT_KO
        section_rules = self._prompts.section_prompt(spec.section_id) or ""
        user_prompt = build_user_prompt(spec, facts)
        if section_rules:
            user_prompt = f"{user_prompt}\n\nSECTION_RULES:\n{section_rules}\n"

        payload: dict[str, Any] = {
            "model": self._model,
            "temperature": 0.2,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        }

        with httpx.Client(timeout=60) as client:
            resp = client.post(url, headers=headers, json=payload)
            resp.raise_for_status()
            data = resp.json()

        content = data["choices"][0]["message"]["content"]
        json_text = _extract_json(content)
        obj = json.loads(json_text)
        return SectionDraft.model_validate(obj)
