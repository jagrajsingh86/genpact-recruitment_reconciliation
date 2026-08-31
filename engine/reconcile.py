"""Pure deterministic reconciliation logic.

This module deliberately has no pandas, openpyxl, Streamlit, or filesystem
dependency. Workbook concerns live in :mod:`engine.io_excel`.
"""

from __future__ import annotations

from dataclasses import dataclass
from html import escape
import re
from typing import Iterable, Mapping, Sequence


OFFER_STAGES = {"offer", "offer approval"}
ACTIVE = {"open", "active"}
SEVERITY_MAP = {"high": "HIGH", "medium": "MEDIUM", "low": "LOW"}

FINDING_COLUMNS = (
    "Requisition No",
    "Position Number",
    "Requisition Title",
    "Recruitment Stage",
    "Recruiter",
    "Hiring Manager",
    "Exception Type",
    "Field",
    "Requisition Value",
    "Position Value",
    "Severity",
    "Status",
)


@dataclass(frozen=True)
class FieldPair:
    req_col: str
    pos_col: str
    rule: str
    weight: str
    status: str
    notes: str = ""


@dataclass(frozen=True)
class ValueNormaliser:
    report: str
    column: str
    pattern: str
    replacement: str


@dataclass(frozen=True)
class ReconciliationConfig:
    field_map: tuple[FieldPair, ...]
    normalisers: tuple[ValueNormaliser, ...]

    @property
    def confirmed_pairs(self) -> tuple[FieldPair, ...]:
        return tuple(pair for pair in self.field_map if canon(pair.status) == "confirmed")


@dataclass(frozen=True)
class Digest:
    recruiter: str
    count: int
    high_count: int
    html: str


@dataclass(frozen=True)
class ReconciliationResult:
    findings: tuple[dict[str, str], ...]
    resolved: tuple[dict[str, str], ...]
    digests: dict[str, Digest]
    summary: dict[str, int | str]


def canon(value: object) -> str:
    """Return the Office Script-compatible canonical comparison value."""

    text = str(value or "")
    return re.sub(r"\s+", " ", text.strip()).lower()


def _as_text(value: object) -> str:
    return str(value or "")


def _normalise_rows(
    rows: Iterable[Mapping[str, object]],
    report: str,
    normalisers: Sequence[ValueNormaliser],
) -> list[dict[str, str]]:
    copied = [{key: _as_text(value) for key, value in row.items()} for row in rows]
    for normaliser in normalisers:
        if canon(normaliser.report) != report:
            continue
        for row in copied:
            if normaliser.column in row:
                row[normaliser.column] = re.sub(
                    normaliser.pattern,
                    normaliser.replacement,
                    row[normaliser.column],
                )
    return copied


def chain_key(finding: Mapping[str, object]) -> str:
    return "|".join(
        (
            _as_text(finding.get("Requisition No")),
            _as_text(finding.get("Field")),
            _as_text(finding.get("Exception Type")),
        )
    )


def _severity(weight: object, recruitment_stage: object) -> str:
    severity = SEVERITY_MAP.get(canon(weight), "MEDIUM")
    if canon(recruitment_stage) in OFFER_STAGES:
        return "HIGH"
    return severity


def _base_finding(req: Mapping[str, str]) -> dict[str, str]:
    return {
        "Requisition No": _as_text(req.get("Requisition No")),
        "Position Number": _as_text(req.get("Position Number")),
        "Requisition Title": _as_text(req.get("Requisition Title (BL)")),
        "Recruitment Stage": _as_text(req.get("Recruitment Stage")),
        "Recruiter": _as_text(req.get("Recruiter (R) Name")),
        "Hiring Manager": _as_text(req.get("Hiring Manager Name")),
    }


def _make_finding(
    req: Mapping[str, str],
    exception_type: str,
    field: str,
    req_value: object,
    pos_value: object,
    weight: object,
) -> dict[str, str]:
    finding = _base_finding(req)
    finding.update(
        {
            "Exception Type": exception_type,
            "Field": field,
            "Requisition Value": _as_text(req_value),
            "Position Value": _as_text(pos_value),
            "Severity": _severity(weight, req.get("Recruitment Stage")),
            "Status": "",
        }
    )
    return finding


def _digest_html(findings: Sequence[Mapping[str, str]]) -> str:
    headers = ("Severity", "Requisition", "Position", "Exception", "Field", "Values", "Status")
    head = "".join(f"<th>{escape(header)}</th>" for header in headers)
    body_rows: list[str] = []
    for finding in findings:
        values = (
            finding["Severity"],
            finding["Requisition No"],
            finding["Position Number"],
            finding["Exception Type"],
            finding["Field"],
            f'{finding["Requisition Value"]} ⟷ {finding["Position Value"]}',
            finding["Status"],
        )
        body_rows.append("<tr>" + "".join(f"<td>{escape(value)}</td>" for value in values) + "</tr>")
    return (
        '<table style="border-collapse:collapse;width:100%">'
        f"<thead><tr>{head}</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody>"
        "</table>"
    )


def _build_digests(findings: Sequence[dict[str, str]]) -> dict[str, Digest]:
    grouped: dict[str, list[dict[str, str]]] = {}
    for finding in findings:
        grouped.setdefault(finding["Recruiter"], []).append(finding)
    return {
        recruiter: Digest(
            recruiter=recruiter,
            count=len(items),
            high_count=sum(item["Severity"] == "HIGH" for item in items),
            html=_digest_html(items),
        )
        for recruiter, items in sorted(grouped.items())
    }


def reconcile(
    requisitions: Iterable[Mapping[str, object]],
    positions: Iterable[Mapping[str, object]],
    config: ReconciliationConfig,
    run_date: str,
    previous_findings: Iterable[Mapping[str, object]] = (),
    *,
    apply_normalisers: bool = True,
) -> ReconciliationResult:
    """Reconcile one day's extracts and assign chained statuses."""

    confirmed_pairs = config.confirmed_pairs
    unsupported = sorted({pair.rule for pair in confirmed_pairs if canon(pair.rule) != "exact"})
    if unsupported:
        raise ValueError(f"Unsupported comparison rule(s): {', '.join(unsupported)}")

    normalisers = config.normalisers if apply_normalisers else ()
    req_rows = _normalise_rows(requisitions, "requisition", normalisers)
    pos_rows = _normalise_rows(positions, "position", normalisers)
    active_requisitions = [row for row in req_rows if canon(row.get("Current Status")) in ACTIVE]

    position_index: dict[str, dict[str, str]] = {}
    duplicate_keys: set[str] = set()
    for position in pos_rows:
        key = canon(position.get("Position Number"))
        if key in position_index:
            duplicate_keys.add(key)
        else:
            position_index[key] = position

    current: list[dict[str, str]] = []
    for req in active_requisitions:
        position_number = req.get("Position Number", "")
        key = canon(position_number)
        if not key:
            current.append(
                _make_finding(req, "MISSING KEY", "Position Number", position_number, "", "high")
            )
            continue
        if key in duplicate_keys:
            current.append(
                _make_finding(
                    req,
                    "DUPLICATE POSITION",
                    "Position Number",
                    position_number,
                    "(multiple position rows)",
                    "high",
                )
            )
            continue
        position = position_index.get(key)
        if position is None:
            current.append(
                _make_finding(
                    req,
                    "ORPHAN REQUISITION (no position row)",
                    "Position Number",
                    position_number,
                    "",
                    "high",
                )
            )
            continue

        for pair in confirmed_pairs:
            req_value = req.get(pair.req_col, "")
            pos_value = position.get(pair.pos_col, "")
            req_canon = canon(req_value)
            pos_canon = canon(pos_value)
            if req_canon == pos_canon:
                continue
            if not pos_canon:
                exception_type = "MISSING IN POSITION REPORT"
            elif not req_canon:
                exception_type = "MISSING IN REQUISITION REPORT"
            else:
                exception_type = "MISMATCH"
            current.append(
                _make_finding(
                    req,
                    exception_type,
                    f"{pair.req_col} <> {pair.pos_col}",
                    req_value,
                    pos_value,
                    pair.weight,
                )
            )

    previous_by_key = {chain_key(finding): finding for finding in previous_findings}
    current_keys: set[str] = set()
    for finding in current:
        key = chain_key(finding)
        current_keys.add(key)
        finding["Status"] = "Recurring" if key in previous_by_key else "New"

    resolved = tuple(
        {
            "Requisition No": _as_text(finding.get("Requisition No")),
            "Field": _as_text(finding.get("Field")),
            "Exception Type": _as_text(finding.get("Exception Type")),
            "Status": "Resolved",
        }
        for key, finding in previous_by_key.items()
        if key not in current_keys
    )
    findings = tuple(current)
    summary: dict[str, int | str] = {
        "run_date": run_date,
        "active_count": len(active_requisitions),
        "pairs_checked": len(confirmed_pairs),
        "exceptions": len(findings),
        "high": sum(finding["Severity"] == "HIGH" for finding in findings),
        "new": sum(finding["Status"] == "New" for finding in findings),
        "recurring": sum(finding["Status"] == "Recurring" for finding in findings),
        "resolved": len(resolved),
    }
    return ReconciliationResult(
        findings=findings,
        resolved=resolved,
        digests=_build_digests(findings),
        summary=summary,
    )
