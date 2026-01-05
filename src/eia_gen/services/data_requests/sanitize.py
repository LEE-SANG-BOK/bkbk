from __future__ import annotations

import re
from typing import Any
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit


# Keep this list intentionally small and conservative.
# These keys frequently carry secrets in Korean public-data/map APIs.
_SENSITIVE_QUERY_KEYS = {
    "servicekey",
    "apikey",
    "api_key",
    "key",
    "token",
    "access_token",
    "authorization",
}


def _is_sensitive_key(key: str) -> bool:
    k = str(key or "").strip().lower()
    return k in _SENSITIVE_QUERY_KEYS


def redact_url(url: str) -> str:
    """Redact common secret-like query params from a URL string."""
    s = str(url or "").strip()
    if not s:
        return s

    try:
        parts = urlsplit(s)
    except Exception:
        return redact_text(s)

    if not parts.query:
        return s

    q = []
    for k, v in parse_qsl(parts.query, keep_blank_values=True):
        if _is_sensitive_key(k):
            q.append((k, "REDACTED"))
        else:
            q.append((k, v))

    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(q), parts.fragment))


def strip_secrets_from_params(params: dict[str, Any] | None) -> dict[str, Any]:
    """Return a shallow-copied params dict without sensitive keys."""
    if not params:
        return {}
    out: dict[str, Any] = {}
    for k, v in params.items():
        if _is_sensitive_key(k):
            continue
        out[k] = v
    return out


_SECRET_KV_RE = re.compile(
    r'(?i)(\b(?:servicekey|apikey|api_key|access_token|token|authorization|key)\b)\s*=\s*([^&\s\'"]+)',
)


def redact_text(text: str) -> str:
    """Redact common `key=value` patterns in arbitrary text."""
    s = str(text or "")
    if not s:
        return s

    def _sub(m: re.Match[str]) -> str:
        k = m.group(1)
        return f"{k}=REDACTED"

    return _SECRET_KV_RE.sub(_sub, s)
