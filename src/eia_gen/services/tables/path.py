from __future__ import annotations

from collections.abc import Iterable
from typing import Any


def _get_attr(obj: Any, name: str) -> Any:
    if obj is None:
        return None
    if isinstance(obj, dict):
        return obj.get(name)
    return getattr(obj, name, None)


def resolve_path(obj: Any, path: str) -> Any:
    """Resolve a dotted path with optional '[]' list expansion.

    Supported patterns:
    - "a.b.c"
    - "a.list[].field"

    Returns:
    - scalar/dict/model when no [] used
    - list when [] used (flattened)
    """

    if not path:
        return None

    parts = path.split(".")
    cur: Any = obj

    for i, part in enumerate(parts):
        if part.endswith("[]"):
            name = part[:-2]
            lst = _get_attr(cur, name)
            if not isinstance(lst, list):
                return []
            rest = ".".join(parts[i + 1 :])
            if not rest:
                return lst
            out: list[Any] = []
            for item in lst:
                v = resolve_path(item, rest)
                if isinstance(v, list):
                    out.extend(v)
                else:
                    out.append(v)
            return out
        cur = _get_attr(cur, part)

    return cur


def coerce_text(v: Any, empty: str) -> str:
    if v is None:
        return empty
    if isinstance(v, str):
        s = v.strip()
        return s if s else empty
    return str(v)


def coerce_number(v: Any, unit: str | None, empty: str) -> str:
    if v is None:
        return empty
    try:
        fv = float(v)
        if fv.is_integer():
            n = f"{int(fv):,}"
        else:
            n = f"{fv:,.2f}"
    except Exception:
        n = str(v)
    return f"{n}{unit or ''}"


def infer_src_path(value_path: str | None) -> str | None:
    if not value_path:
        return None
    if value_path.endswith(".t") or value_path.endswith(".v"):
        return value_path.rsplit(".", 1)[0] + ".src"
    # raw dict values may store src at same level
    return None

