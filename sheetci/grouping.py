"""Collapse per-cell findings into pattern groups and score them.

The scanner reports one finding per cell so that JSON consumers keep full detail.
A formula copied down forty rows is one problem, though, not forty, so grouping
happens here and the risk score is computed over groups.
"""

import re
from typing import Any, Dict, List, Optional, Tuple

from sheetci.formula import normalize_formula

SEVERITY_WEIGHTS = {"critical": 30, "warning": 10, "info": 3}
MAX_RISK_SCORE = 100
FAIL_THRESHOLD = 70
MAX_RENDERED_RUNS = 3

_CELL_ADDRESS = re.compile(r"^\$?([A-Z]{1,3})\$?([0-9]+)$", re.IGNORECASE)


def _cell_parts(address: Optional[str]) -> Tuple[Optional[str], Optional[int]]:
    match = _CELL_ADDRESS.match(address or "")
    if not match:
        return None, None
    return match.group(1).upper(), int(match.group(2))


def format_cell_range(cells: List[str]) -> str:
    """Summarise a list of cell addresses as contiguous runs, e.g. `I2:I41`."""
    if not cells:
        return "Sheet-Level"

    runs: List[List] = []  # [column, first_row, last_row, literal_or_None]
    for address in cells:
        column, row = _cell_parts(address)
        if column is None:
            runs.append([None, None, None, address])
            continue
        if runs and runs[-1][0] == column and runs[-1][2] == row - 1:
            runs[-1][2] = row
        else:
            runs.append([column, row, row, None])

    rendered = []
    for column, first, last, literal in runs[:MAX_RENDERED_RUNS]:
        if literal is not None:
            rendered.append(literal)
        elif first == last:
            rendered.append(f"{column}{first}")
        else:
            rendered.append(f"{column}{first}:{column}{last}")

    remaining = len(runs) - MAX_RENDERED_RUNS
    if remaining > 0:
        rendered.append(f"(+{remaining} more)")
    return ", ".join(rendered)


def group_findings(findings: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Collapse findings that share a sheet, rule, and normalised formula pattern.

    Findings without a cell address or without a formula (sheet-level rules such
    as HIDDEN_SHEET) are passed through untouched, one group each.
    """
    grouped: Dict[Any, Dict[str, Any]] = {}
    ordered: List[Dict[str, Any]] = []

    for index, finding in enumerate(findings):
        address = finding.get("cell_address")
        formula = finding.get("formula")
        _, row = _cell_parts(address)

        if not address or not formula or row is None:
            key: Any = ("__ungrouped__", index)
        else:
            key = (
                finding["sheet_name"],
                finding["rule_id"],
                normalize_formula(formula, row),
            )

        existing = grouped.get(key)
        if existing is None:
            group = dict(finding)
            group["cells"] = [address] if address else []
            group["occurrences"] = 1
            grouped[key] = group
            ordered.append(group)
        else:
            existing["cells"].append(address)
            existing["occurrences"] += 1

    for group in ordered:
        group["cell_range"] = format_cell_range(group["cells"])
    return ordered


def score_findings(groups: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Compute risk score and pass/fail status over groups, not occurrences."""
    counts = {"critical": 0, "warning": 0, "info": 0}
    risk_score = 0

    for group in groups:
        severity = group["severity"]
        if severity in counts:
            counts[severity] += 1
            risk_score += SEVERITY_WEIGHTS[severity]

    risk_score = min(risk_score, MAX_RISK_SCORE)
    is_pass = counts["critical"] == 0 and risk_score < FAIL_THRESHOLD

    return {
        "risk_score": risk_score,
        "status": "PASS" if is_pass else "FAIL",
        "counts": counts,
    }
