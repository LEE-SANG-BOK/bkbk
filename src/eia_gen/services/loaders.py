from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

from eia_gen.models.case import Case
from eia_gen.models.sources import SourceRegistry
from eia_gen.services.canonicalize import canonicalize_case


def load_yaml(path: str | Path) -> Any:
    p = Path(path)
    raw = p.read_text(encoding="utf-8")
    return yaml.safe_load(raw)


def load_case(path: str | Path) -> Case:
    data = load_yaml(path)
    case = Case.model_validate(data or {})
    return canonicalize_case(case)


def load_sources(path: str | Path) -> SourceRegistry:
    data = load_yaml(path)
    return SourceRegistry.model_validate(data or {})

