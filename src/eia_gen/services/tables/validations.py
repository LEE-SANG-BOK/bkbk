from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Literal

from eia_gen.models.case import Case
from eia_gen.services.tables.path import resolve_path
from eia_gen.spec.models import TableSpec, TableValidation


Severity = Literal["ERROR", "WARN", "INFO"]


@dataclass(frozen=True)
class ValidationFinding:
    rule_id: str
    severity: Severity
    message: str


_YM_RE = re.compile(r"^\d{4}-\d{2}$")


def validate_table(case: Case, spec: TableSpec) -> list[ValidationFinding]:
    findings: list[ValidationFinding] = []

    for v in spec.validations:
        params: dict[str, Any] = v.model_dump()
        vtype = params.get("type")
        rid = params.get("id") or "UNKNOWN"
        sev: Severity = params.get("severity") or "ERROR"

        if vtype == "min_rows":
            min_rows = int(params.get("min") or 0)
            data = resolve_path(case, spec.data_path or "")
            count = len(data) if isinstance(data, (list, dict)) else 0
            if count < min_rows:
                findings.append(
                    ValidationFinding(rid, sev, f"{spec.id}: 행 수 부족({count} < {min_rows})")
                )
            continue

        if vtype == "sum_equals":
            sum_path = str(params.get("sum_path") or "")
            total_path = str(params.get("total_path") or "")
            tol = float(params.get("tolerance_ratio") or 0.0)
            values = resolve_path(case, sum_path)
            total = resolve_path(case, total_path)
            s = 0.0
            if isinstance(values, list):
                for x in values:
                    try:
                        if x is None:
                            continue
                        s += float(x)
                    except Exception:
                        continue
            try:
                t = float(total) if total is not None else None
            except Exception:
                t = None
            if t is None:
                findings.append(ValidationFinding(rid, sev, f"{spec.id}: 총면적 값 누락/비정상"))
            else:
                allowed = abs(t) * tol
                if abs(s - t) > allowed:
                    findings.append(
                        ValidationFinding(
                            rid,
                            sev,
                            f"{spec.id}: 합계 불일치(sum={s:.2f}, total={t:.2f}, tol={tol})",
                        )
                    )
            continue

        if vtype == "date_format":
            path = str(params.get("path") or "")
            fmt = str(params.get("format") or "YYYY-MM")
            vals = resolve_path(case, path)
            if not isinstance(vals, list):
                vals = [vals]
            if fmt == "YYYY-MM":
                bad = [x for x in vals if isinstance(x, str) and x and not _YM_RE.match(x)]
                if bad:
                    findings.append(ValidationFinding(rid, sev, f"{spec.id}: 날짜 형식 오류({bad[0]!r})"))
            continue

        if vtype == "required_if":
            # Row-level conditional required check
            if not spec.data_path:
                continue
            cond = params.get("condition") or {}
            if not isinstance(cond, dict):
                continue
            cond_path = str(cond.get("path") or "")
            cond_equals = str(cond.get("equals") or "")
            required_path = str(params.get("required_path") or "")

            rows = resolve_path(case, spec.data_path)
            if not isinstance(rows, list):
                continue

            for row in rows:
                v_cond = resolve_path(row, cond_path)
                if isinstance(v_cond, dict) and "t" in v_cond:
                    v_cond = v_cond.get("t")
                if str(v_cond).strip() != cond_equals:
                    continue
                req = resolve_path(row, required_path)
                if isinstance(req, dict) and "t" in req:
                    req = req.get("t")
                if req is None or str(req).strip() == "":
                    findings.append(
                        ValidationFinding(rid, sev, f"{spec.id}: 조건부 필수값 누락({required_path})")
                    )
                    break

            continue

    return findings
