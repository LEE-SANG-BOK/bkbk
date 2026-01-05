from __future__ import annotations

from pathlib import Path


class PromptLibrary:
    def __init__(self, root: str | Path = "prompts") -> None:
        self._root = Path(root)

    def system_prompt(self) -> str | None:
        p = self._root / "system_ko.txt"
        if not p.exists():
            return None
        return p.read_text(encoding="utf-8")

    def section_prompt(self, section_id: str) -> str | None:
        p = self._root / "section" / f"{section_id}.txt"
        if not p.exists():
            return None
        return p.read_text(encoding="utf-8")

    def postprocess_prompt(self, name: str) -> str | None:
        p = self._root / "postprocess" / f"{name}.txt"
        if not p.exists():
            return None
        return p.read_text(encoding="utf-8")

