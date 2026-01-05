from __future__ import annotations

import re
from typing import Any

from eia_gen.services.tables.path import resolve_path


_COND_RE = re.compile(r"^\s*([A-Za-z0-9_\.]+)\s*==\s*(true|false)\s*$", re.IGNORECASE)


def _truthy(v: Any) -> bool:
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    if isinstance(v, dict) and "t" in v:
        return _truthy(v.get("t"))
    s = str(v).strip().lower()
    return s in {"true", "t", "yes", "y", "1"}


def eval_condition(obj: Any, expr: str | None) -> bool:
    """Evaluate simple boolean conditions used in spec.

    Supported:
    - "path.to.field == true|false"
    """

    if not expr:
        return True
    m = _COND_RE.match(expr)
    if not m:
        # Unknown expression -> be safe: treat as False (forces explicit)
        return False
    path, tf = m.group(1), m.group(2).lower()
    expected = tf == "true"
    v = resolve_path(obj, path)
    # unwrap TextField-like
    if hasattr(v, "t"):
        v = getattr(v, "t")
    return _truthy(v) == expected

