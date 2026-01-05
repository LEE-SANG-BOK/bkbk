from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

from eia_gen.spec.models import FigureSpecs, SectionsSpec, SpecBundle, TableSpecs, TemplateMap


def _load_yaml(path: Path) -> Any:
    return yaml.safe_load(path.read_text(encoding="utf-8"))


def load_spec_bundle(spec_dir: str | Path = "spec") -> SpecBundle:
    base = Path(spec_dir)
    sections = SectionsSpec.model_validate(_load_yaml(base / "sections.yaml") or {})
    tables = TableSpecs.model_validate(_load_yaml(base / "table_specs.yaml") or {})
    figures = FigureSpecs.model_validate(_load_yaml(base / "figure_specs.yaml") or {})
    template_map = TemplateMap.model_validate(_load_yaml(base / "template_map.yaml") or {})
    return SpecBundle(sections=sections, tables=tables, figures=figures, template_map=template_map)

