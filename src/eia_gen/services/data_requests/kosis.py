from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import httpx
import yaml


# NOTE:
# - KOSIS의 orgId/tblId 기반 조회는 `statisticsData.do`보다
#   `openapi/Param/statisticsParameterData.do` 엔드포인트가 안정적으로 동작한다.
# - `statisticsData.do`는 (일부 환경에서) userStatsId 기반 호출 위주로 문서화되어 있어,
#   본 프로젝트의 dataset_key(SSOT) 방식에는 Param 엔드포인트를 기본값으로 둔다.
KOSIS_DATA_URL = "https://kosis.kr/openapi/Param/statisticsParameterData.do"


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _api_key() -> str:
    return os.environ.get("KOSIS_API_KEY", "").strip()


def _as_int(v: Any) -> int | None:
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return v
    if isinstance(v, float):
        return int(v)
    s = str(v).strip()
    if not s or s in {"-", "NA", "N/A"}:
        return None
    s = s.replace(",", "")
    try:
        return int(float(s))
    except Exception:
        return None


def _safe_upper(v: Any) -> str:
    return str(v or "").strip().upper()


def _load_yaml(path: Path) -> dict[str, Any]:
    try:
        obj = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except Exception as e:
        raise ValueError(f"Failed to load YAML: {path} ({e})") from None
    if not isinstance(obj, dict):
        raise ValueError(f"Invalid YAML (expected mapping): {path}")
    return obj


def load_kosis_dataset_catalog(path: Path) -> dict[str, dict[str, Any]]:
    """Load `config/kosis_datasets.yaml` and return `dataset_key -> dataset dict`."""
    obj = _load_yaml(path)
    datasets = obj.get("datasets") or {}
    if not isinstance(datasets, dict):
        raise ValueError(f"Invalid kosis_datasets.yaml: datasets must be a mapping ({path})")
    out: dict[str, dict[str, Any]] = {}
    for k, v in datasets.items():
        key = str(k or "").strip()
        if not key or not isinstance(v, dict):
            continue
        out[key] = dict(v)
    return out


def _format_templates(obj: Any, ctx: dict[str, Any]) -> Any:
    if isinstance(obj, str):
        s = obj
        if "{" in s and "}" in s:
            try:
                return s.format(**ctx)
            except KeyError as e:
                missing = str(e.args[0]) if e.args else "?"
                raise ValueError(f"Missing template variable: {missing}") from None
            except Exception as e:
                raise ValueError(f"Failed to format template: {s!r} ({e})") from None
        return s
    if isinstance(obj, dict):
        return {k: _format_templates(v, ctx) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_format_templates(v, ctx) for v in obj]
    return obj


def _is_placeholder_value(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return True
        if s.upper() in {"TBD", "TODO"}:
            return True
    return False


def resolve_kosis_dataset(
    *,
    dataset_key: str,
    config_path: Path,
    context: dict[str, Any],
) -> tuple[KosisQuery, list[KosisMapping], dict[str, Any]]:
    """Resolve `dataset_key` into a concrete KOSIS query + mappings using SSOT config."""
    key = (dataset_key or "").strip()
    if not key:
        raise ValueError("KOSIS dataset_key is empty")
    catalog = load_kosis_dataset_catalog(config_path)
    ds = catalog.get(key)
    if not ds:
        raise ValueError(f"Unknown KOSIS dataset_key: {key} (check {config_path})")

    endpoint_url = str(ds.get("endpoint_url") or KOSIS_DATA_URL).strip() or KOSIS_DATA_URL
    method = str(ds.get("method") or "getList").strip() or "getList"
    fmt = str(ds.get("format") or "json").strip() or "json"
    json_vd = str(ds.get("json_vd") or "Y").strip() or "Y"

    query_tpl = ds.get("query_params")
    if not isinstance(query_tpl, dict):
        raise ValueError(f"KOSIS dataset {key} missing query_params (in {config_path})")
    query_params = _format_templates(query_tpl, context)
    if not isinstance(query_params, dict):
        raise ValueError(f"KOSIS dataset {key} query_params must resolve to an object (in {config_path})")

    missing_keys = [k for k, v in query_params.items() if _is_placeholder_value(v)]
    if missing_keys:
        raise ValueError(
            f"KOSIS dataset {key} has placeholder query_params values: {', '.join(sorted(missing_keys))} "
            f"(fill {config_path})"
        )

    mappings_raw = ds.get("mappings")
    if not isinstance(mappings_raw, list):
        # Backward-compat: allow `outputs` list from early drafts.
        mappings_raw = ds.get("outputs")
    if not isinstance(mappings_raw, list) or not mappings_raw:
        raise ValueError(f"KOSIS dataset {key} missing mappings list (in {config_path})")

    mappings: list[KosisMapping] = []
    for m in mappings_raw:
        if not isinstance(m, dict):
            continue
        output_col = str(m.get("output_col") or m.get("col") or "").strip()
        if not output_col:
            continue
        match_itm_id = str(m.get("match_itm_id") or m.get("itm_id") or "").strip()
        match_itm_nm_contains = str(m.get("match_itm_nm_contains") or m.get("itm_nm_contains") or m.get("field") or "").strip()
        if not match_itm_id and not match_itm_nm_contains:
            continue
        mappings.append(
            KosisMapping(
                output_col=output_col,
                match_itm_id=match_itm_id,
                match_itm_nm_contains=match_itm_nm_contains,
            )
        )
    if not mappings:
        raise ValueError(f"KOSIS dataset {key} mappings has no valid entries (in {config_path})")

    q = KosisQuery(
        query_params={str(k): v for k, v in query_params.items()},
        endpoint_url=endpoint_url,
        method=method,
        format=fmt,
        json_vd=json_vd,
    )

    ds_meta = {
        "dataset_key": key,
        "description": str(ds.get("description") or "").strip(),
        "endpoint_url": endpoint_url,
        "method": method,
        "format": fmt,
        "json_vd": json_vd,
        "query_params": q.query_params,
        "mappings": [m.__dict__ for m in mappings],
    }
    return q, mappings, ds_meta


@dataclass(frozen=True)
class KosisQuery:
    query_params: dict[str, Any]
    endpoint_url: str = KOSIS_DATA_URL
    method: str = "getList"
    format: str = "json"
    json_vd: str = "Y"


@dataclass(frozen=True)
class KosisMapping:
    output_col: str
    match_itm_id: str = ""
    match_itm_nm_contains: str = ""


@dataclass(frozen=True)
class KosisFetchResult:
    items: list[dict[str, Any]]
    evidence_json: dict[str, Any]


def fetch_kosis_series(*, q: KosisQuery, timeout_sec: int = 25) -> KosisFetchResult:
    """Fetch raw KOSIS series data (best-effort).

    The caller is responsible for passing correct KOSIS parameters (orgId/tblId/...)
    via `q.query_params`. This function only ensures reproducibility by keeping the
    request/response in evidence.
    """
    key = _api_key()
    if not key:
        raise ValueError("Missing KOSIS_API_KEY")

    params: dict[str, Any] = {
        "method": q.method,
        "apiKey": key,
        "format": q.format,
        "jsonVD": q.json_vd,
    }
    params.update(q.query_params or {})

    def _fetch_once(p: dict[str, Any]) -> list[dict[str, Any]]:
        with httpx.Client(timeout=timeout_sec, follow_redirects=True) as client:
            try:
                r = client.get(q.endpoint_url, params=p)
                r.raise_for_status()
                try:
                    data0 = r.json()
                except Exception:
                    raise ValueError("KOSIS response is not JSON")
            except httpx.HTTPStatusError as e:
                raise ValueError(f"KOSIS request failed: HTTP {e.response.status_code}") from None
            except httpx.HTTPError:
                raise ValueError("KOSIS request failed") from None

        # KOSIS typically returns a list of dicts on success, or a dict on error.
        if isinstance(data0, dict):
            raise ValueError(f"KOSIS error response: {data0}")
        if not isinstance(data0, list):
            raise ValueError("KOSIS response has unexpected type")
        return [dict(x or {}) for x in data0 if isinstance(x, dict)]

    # Some KOSIS endpoints reject multi-valued `itmId` (e.g., "T100,T200") with err=21.
    # When detected, fan-out into multiple calls and merge the returned rows.
    itm_key = None
    itm_raw = None
    for k in list(params.keys()):
        if str(k).strip().lower() == "itmid":
            itm_key = str(k)
            itm_raw = params.get(k)
            break

    itm_ids: list[str] = []
    if itm_key:
        s = str(itm_raw or "").strip()
        if "," in s or ";" in s:
            s2 = s.replace(";", ",")
            itm_ids = [p.strip() for p in s2.split(",") if p.strip()]

    requests_meta: list[dict[str, Any]] = []
    items: list[dict[str, Any]] = []
    if itm_ids and len(itm_ids) >= 2 and itm_key:
        for one in itm_ids:
            p = dict(params)
            p[itm_key] = one
            requests_meta.append({"url": q.endpoint_url, "params": {k: v for k, v in p.items() if k != "apiKey"}})
            items.extend(_fetch_once(p))
    else:
        requests_meta.append({"url": q.endpoint_url, "params": {k: v for k, v in params.items() if k != "apiKey"}})
        items = _fetch_once(params)

    evidence = {
        "generated_at": _now_iso(),
        # Backward-compat: keep a single `request` object for downstream parsers.
        "request": requests_meta[0] if requests_meta else {"url": q.endpoint_url, "params": {}},
        "requests": requests_meta,
        "response": items,
        "computed": {
            "row_count": len(items),
        },
    }
    return KosisFetchResult(items=items, evidence_json=evidence)


def build_env_base_socio_rows(
    *,
    items: list[dict[str, Any]],
    mappings: list[KosisMapping],
    admin_code: str,
    admin_name: str,
) -> list[dict[str, Any]]:
    """Convert KOSIS items into ENV_BASE_SOCIO rows (year → wide format).

    Expected common KOSIS fields:
    - year: PRD_DE
    - item id/name: ITM_ID / ITM_NM
    - value: DT
    """
    by_year: dict[str, dict[str, Any]] = {}

    for it in items:
        year = str(it.get("PRD_DE") or it.get("PRD") or "").strip()
        if not year:
            continue
        itm_id = _safe_upper(it.get("ITM_ID"))
        itm_nm = str(it.get("ITM_NM") or "").strip()
        dt = it.get("DT")

        row = by_year.setdefault(
            year,
            {
                "admin_code": admin_code,
                "admin_name": admin_name,
                "year": year,
            },
        )

        for m in mappings:
            if m.match_itm_id and itm_id != _safe_upper(m.match_itm_id):
                continue
            if m.match_itm_nm_contains and m.match_itm_nm_contains not in itm_nm:
                continue
            val = _as_int(dt)
            if val is None:
                continue
            row[m.output_col] = val

    # Stable ordering
    return [by_year[y] for y in sorted(by_year.keys())]


def evidence_bytes(evidence: dict[str, Any]) -> bytes:
    return json.dumps(evidence, ensure_ascii=False, indent=2).encode("utf-8")
