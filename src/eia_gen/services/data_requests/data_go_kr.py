from __future__ import annotations

from typing import Any
from urllib.parse import quote, urlencode


def build_url(
    *,
    base_url: str,
    service_key: str,
    params: dict[str, Any],
    key_param: str = "serviceKey",
) -> str:
    """Build a data.go.kr URL without double-encoding `serviceKey`.

    data.go.kr provides both:
    - Decoding(일반) key: contains reserved characters like +/=
    - Encoding key: already percent-encoded

    When using an HTTP client with `params=...`, an already-encoded key can get
    encoded again (% -> %25). This helper keeps the key stable by embedding it
    directly in the URL query string.
    """
    sk = str(service_key or "").strip()
    if not sk:
        raise ValueError("Missing service_key")

    # If the key looks already percent-encoded, keep it as-is.
    if "%" not in sk:
        sk = quote(sk, safe="")

    # Remove any accidental key param from params.
    params2 = {k: v for k, v in (params or {}).items() if str(k) != key_param}

    tail = urlencode(params2, doseq=True)
    if tail:
        return f"{base_url}?{key_param}={sk}&{tail}"
    return f"{base_url}?{key_param}={sk}"
