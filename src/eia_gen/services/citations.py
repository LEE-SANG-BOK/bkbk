from __future__ import annotations

import re


_CITATION_RE = re.compile(r"〔[^〕]+〕")


def normalize_ids(ids: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for s in ids:
        s2 = s.strip()
        if not s2:
            continue
        if s2 in seen:
            continue
        seen.add(s2)
        out.append(s2)
    return out


def format_citations(ids: list[str] | None) -> str:
    ids2 = normalize_ids(ids or [])
    if not ids2:
        ids2 = ["SRC-TBD"]
    prefixed: list[str] = []
    for s in ids2:
        s2 = s.strip()
        if not s2:
            continue
        if s2 in {"S-TBD", "SRC-TBD"}:
            prefixed.append("SRC-TBD")
            continue
        if s2.upper().startswith("SRC:"):
            prefixed.append(s2)
        else:
            prefixed.append(f"SRC:{s2}")
    return f"〔{','.join(prefixed)}〕"


def has_citation(text: str) -> bool:
    return bool(_CITATION_RE.search(text))


def ensure_citation(text: str, ids: list[str] | None = None) -> str:
    if has_citation(text):
        return text
    return f"{text} {format_citations(ids)}"


def strip_citations(text: str) -> str:
    """Remove inline citation blocks like `〔SRC:...〕` from text for layout-sensitive outputs."""
    if not text:
        return ""
    return _CITATION_RE.sub("", text).strip()


def extract_citation_ids(text: str) -> list[str]:
    """Extract normalized source ids from inline citation blocks.

    Accepts tokens like:
    - `〔SRC:S-01,SRC:S-02〕`
    - `〔SRC-TBD〕`
    """
    if not text:
        return []
    ids: list[str] = []
    for block in _CITATION_RE.findall(text):
        inner = str(block).strip()
        if inner.startswith("〔") and inner.endswith("〕"):
            inner = inner[1:-1]
        for token in inner.split(","):
            t = token.strip()
            if not t:
                continue
            if t.upper().startswith("SRC:"):
                t = t[4:].strip()
            if t.upper() == "SRC-TBD":
                t = "S-TBD"
            if t:
                ids.append(t)
    return normalize_ids(ids)
