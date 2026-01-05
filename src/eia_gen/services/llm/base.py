from __future__ import annotations

from abc import ABC, abstractmethod

from eia_gen.services.draft import SectionDraft
from eia_gen.services.sections import SectionSpec


class LLMClient(ABC):
    @abstractmethod
    def generate_section(self, spec: SectionSpec, facts: dict) -> SectionDraft: ...

