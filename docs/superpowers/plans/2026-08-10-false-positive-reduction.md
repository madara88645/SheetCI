# False-Positive Reduction Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make SheetCI pass a healthy workbook by aggregating repeated findings, re-tiering `HARDCODED_NUMBER` to info, and correcting four detectors that fire on ordinary spreadsheet idioms.

**Architecture:** Formula-text analysis moves out of `scanner.py` into a new pure-function module `sheetci/formula.py` (no openpyxl dependency, fully unit-testable). A second new module `sheetci/grouping.py` collapses per-cell findings into pattern groups and computes the risk score over groups. `scanner.py` keeps detection and openpyxl access, delegating text analysis and aggregation to those two modules.

**Tech Stack:** Python ≥3.11, openpyxl, typer, jinja2, pytest. Package manager is `uv`.

## Global Constraints

- Python `>=3.11` (`pyproject.toml`). No new runtime dependencies.
- All repository text — code comments, docstrings, commit messages, README — in English.
- Tests run with `uv run pytest` from the repository root.
- Work happens on branch `fix/reduce-false-positives`. Never commit to `main`, never push without explicit approval.
- The spec is `docs/superpowers/specs/2026-08-10-false-positive-reduction-design.md`. Its acceptance criteria are the definition of done.
- `examples/broken-commission-model.xlsx` must keep scanning to `FAIL` with `BROKEN_REF`, `SELF_REFERENCE`, and `CACHED_ERROR` present. Any task that breaks this is wrong.

## File Structure

| File | Responsibility |
| --- | --- |
| `sheetci/formula.py` | **New.** Pure formula-string analysis: reference normalisation, argument-aware numeric-literal extraction, structured-table-reference stripping, external-reference detection, column-aggregate detection. No openpyxl import. |
| `sheetci/grouping.py` | **New.** Collapses per-cell findings into groups keyed by normalised formula pattern; formats cell ranges; computes risk score and severity counts over groups. |
| `sheetci/scanner.py` | **Modified.** Keeps openpyxl access and detector logic. Imports text helpers from `formula.py`, re-exports `normalize_formula` and `extract_hardcoded_numbers` for backwards compatibility. Emits a `formula` field on cell-level findings. Delegates scoring to `grouping.py`. |
| `sheetci/reporters.py` | **Modified.** Renders occurrence counts and cell ranges. |
| `sheetci/cli.py` | **Modified.** Console summary renders occurrence counts and cell ranges. |
| `scripts/create_example_wb.py` | **Modified.** Adds `create_realistic_model`. |
| `tests/test_formula.py` | **New.** Unit tests for `formula.py`. |
| `tests/test_grouping.py` | **New.** Unit tests for `grouping.py`. |
| `tests/test_scanner.py` | **Modified.** Detector-level tests; two existing assertions updated. |
| `tests/test_cli.py` | **Modified.** End-to-end regression tests. |

---

### Task 1: Argument-aware numeric literal extraction

The single largest noise source. `extract_hardcoded_numbers` currently strips references with regexes and flags every remaining literal, so structural arguments like the `5` in `VLOOKUP(...,5,FALSE)` are reported as magic numbers. Replaced with a walker that knows which function and which argument position each literal sits in.

**Files:**
- Create: `sheetci/formula.py`
- Test: `tests/test_formula.py`

**Interfaces:**
- Consumes: nothing from earlier tasks.
- Produces:
  - `iter_numeric_literals(formula: str) -> Iterator[tuple[float, str | None, int]]` — yields `(value, innermost_function_name_upper_or_None, one_based_arg_index)`. Arg index is `0` when the literal is at top level.
  - `extract_hardcoded_numbers(formula: str) -> List[float]` — same signature and semantics as the existing `sheetci.scanner.extract_hardcoded_numbers`, plus structural-argument suppression.
  - `IGNORED_CONSTANTS: set[float]`, `STRUCTURAL_ARGS: dict[str, set[int]]`

- [ ] **Step 1: Write the failing test**

Create `tests/test_formula.py`:

```python
import pytest
from sheetci.formula import extract_hardcoded_numbers, iter_numeric_literals


def test_plain_constant_is_reported():
    assert extract_hardcoded_numbers("=A1*0.08") == [0.08]


def test_harmless_constants_are_ignored():
    assert extract_hardcoded_numbers("=SUM(A1:A10, 1, 0, -2)") == []


def test_numbers_inside_identifiers_are_not_literals():
    # DEC2HEX and the cell reference A1 must not contribute numbers.
    assert extract_hardcoded_numbers('=IF(DEC2HEX(A1)="1A", 100, 2)') == [100.0]


def test_vlookup_column_index_is_structural():
    assert extract_hardcoded_numbers('=VLOOKUP("Rep 1",Sales!A2:I41,5,FALSE)') == []


def test_eomonth_month_count_is_structural():
    assert extract_hardcoded_numbers("=EOMONTH(TODAY(),12)") == []


def test_match_type_is_structural():
    assert extract_hardcoded_numbers('=INDEX(Sales!G2:G41,MATCH("Rep 3",Sales!A2:A41,0))') == []


def test_innermost_function_wins_for_nesting():
    # 1.15 is ROUND argument 1 (reported); 2 is ROUND argument 2 (structural).
    assert extract_hardcoded_numbers("=ROUND(B2*1.15,2)") == [1.15]


def test_comparison_thresholds_are_reported():
    formula = '=IF(C2>150,"High",IF(C2>120,"Mid","Low"))'
    assert extract_hardcoded_numbers(formula) == [150.0, 120.0]


def test_commas_inside_strings_do_not_shift_argument_index():
    # The "a,b" string must not push 7 into argument position 3 of VLOOKUP.
    assert extract_hardcoded_numbers('=VLOOKUP("a,b",R,7,FALSE)') == []


def test_absolute_references_are_not_literals():
    assert extract_hardcoded_numbers("=E2*Assumptions!$B$2") == []


def test_row_range_is_not_a_literal():
    assert extract_hardcoded_numbers("=SUM(5:10)") == []


def test_malformed_formula_does_not_raise():
    assert extract_hardcoded_numbers('=SUM(A1,"unterminated') == []
    assert extract_hardcoded_numbers("=((((") == []


def test_iter_reports_function_and_argument_position():
    literals = list(iter_numeric_literals("=VLOOKUP(A1,R,5,FALSE)"))
    assert literals == [(5.0, "VLOOKUP", 3)]


def test_iter_reports_none_function_at_top_level():
    assert list(iter_numeric_literals("=B3/B2*100")) == [(100.0, None, 0)]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_formula.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'sheetci.formula'`

- [ ] **Step 3: Write minimal implementation**

Create `sheetci/formula.py`:

```python
"""Pure formula-string analysis helpers.

This module deliberately has no openpyxl dependency: everything here operates on
formula text alone, so it can be unit-tested without building workbooks.
"""

import re
from typing import Dict, Iterator, List, Optional, Set, Tuple

# Constants that are never worth reporting as magic numbers.
IGNORED_CONSTANTS: Set[float] = {0.0, 1.0, 2.0}

# Argument positions (1-based) whose numeric literals are structural parts of the
# function call rather than business constants. A VLOOKUP column index or an
# EOMONTH month count says nothing about the model's assumptions.
STRUCTURAL_ARGS: Dict[str, Set[int]] = {
    "VLOOKUP": {3},
    "HLOOKUP": {3},
    "MATCH": {3},
    "INDEX": {2, 3},
    "ROUND": {2},
    "ROUNDUP": {2},
    "ROUNDDOWN": {2},
    "EOMONTH": {2},
    "EDATE": {2},
    "OFFSET": {2, 3, 4, 5},
    "LARGE": {2},
    "SMALL": {2},
    "SUBTOTAL": {1},
    "LEFT": {2},
    "RIGHT": {2},
    "MID": {2, 3},
    "WEEKDAY": {2},
}

_IDENT_START = set("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_$")
_IDENT_CHARS = set("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_.$")


def _skip_quoted(formula: str, i: int, quote: str) -> int:
    """Return the index just past the quoted run starting at `i`.

    Excel escapes a quote by doubling it, so `""` inside a string does not end it.
    An unterminated string consumes the rest of the formula rather than raising.
    """
    i += 1
    while i < len(formula):
        if formula[i] == quote:
            if i + 1 < len(formula) and formula[i + 1] == quote:
                i += 2
                continue
            return i + 1
        i += 1
    return i


def iter_numeric_literals(formula: str) -> Iterator[Tuple[float, Optional[str], int]]:
    """Yield (value, innermost function name, 1-based argument index) per literal.

    The function name is upper-cased, or None when the literal sits outside any
    call. The argument index is 0 in that case.

    A digit run only counts as a literal when it does not continue an identifier,
    which is what keeps cell references (A1, $B$10) and function names containing
    digits (LOG10, DEC2HEX) from being mistaken for constants.
    """
    stack: List[List] = []  # [function_name, arg_index]
    last_ident: Optional[str] = None
    i, n = 0, len(formula)

    while i < n:
        ch = formula[i]

        if ch in ('"', "'"):
            i = _skip_quoted(formula, i, ch)
            last_ident = None
            continue

        if ch in _IDENT_START:
            start = i
            while i < n and formula[i] in _IDENT_CHARS:
                i += 1
            last_ident = formula[start:i]
            continue

        if ch == "(":
            stack.append([(last_ident or "").upper(), 1])
            last_ident = None
            i += 1
            continue

        if ch == ")":
            if stack:
                stack.pop()
            last_ident = None
            i += 1
            continue

        if ch in (",", ";"):
            if stack:
                stack[-1][1] += 1
            last_ident = None
            i += 1
            continue

        if ch.isdigit() or (ch == "." and i + 1 < n and formula[i + 1].isdigit()):
            start = i
            while i < n and (formula[i].isdigit() or formula[i] == "."):
                i += 1
            text = formula[start:i]

            # A whole-row range such as 5:10 is structural, not a constant.
            if i < n and formula[i] == ":":
                j = i + 1
                while j < n and formula[j].isdigit():
                    j += 1
                if j > i + 1:
                    i = j
                    last_ident = None
                    continue

            try:
                value = float(text)
            except ValueError:
                last_ident = None
                continue

            if stack:
                yield value, stack[-1][0] or None, stack[-1][1]
            else:
                yield value, None, 0
            last_ident = None
            continue

        last_ident = None
        i += 1


def extract_hardcoded_numbers(formula: str) -> List[float]:
    """Return business constants embedded in a formula, in order of appearance."""
    results: List[float] = []
    for value, func, arg_index in iter_numeric_literals(formula):
        if value in IGNORED_CONSTANTS:
            continue
        if func and arg_index in STRUCTURAL_ARGS.get(func, frozenset()):
            continue
        results.append(value)
    return results
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_formula.py -v`
Expected: PASS, 14 tests.

- [ ] **Step 5: Commit**

```bash
git add sheetci/formula.py tests/test_formula.py
git commit -m "feat(formula): extract constants with function-argument awareness"
```

---

### Task 2: Structured table references vs real external links

`EXTERNAL_LINK` currently flags any formula containing `[`, which matches Excel structured table references such as `=SUM(Table1[Amount])`.

**Files:**
- Modify: `sheetci/formula.py`
- Test: `tests/test_formula.py`

**Interfaces:**
- Consumes: `sheetci/formula.py` from Task 1.
- Produces:
  - `strip_table_references(formula: str) -> str`
  - `has_external_reference(formula: str) -> bool`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_formula.py`:

```python
from sheetci.formula import has_external_reference, strip_table_references


def test_simple_table_reference_is_not_external():
    assert has_external_reference("=SUM(Table1[Amount])") is False


def test_nested_table_reference_is_not_external():
    assert has_external_reference("=SUM(Table1[[#Headers],[Amount]])") is False


def test_external_workbook_link_is_detected():
    assert has_external_reference("=[Budget.xlsx]Sheet1!B2") is True


def test_quoted_path_external_link_is_detected():
    assert has_external_reference("='C:\\\\models\\\\[q3.xlsx]Sheet1'!A1") is True


def test_url_is_detected():
    assert has_external_reference('=WEBSERVICE("https://example.com/rates")') is True


def test_plain_formula_is_not_external():
    assert has_external_reference("=A1*B1") is False


def test_strip_leaves_surrounding_text():
    assert strip_table_references("=SUM(Table1[Amount])").strip() == "=SUM( )".strip()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_formula.py -k external -v`
Expected: FAIL — `ImportError: cannot import name 'has_external_reference'`

- [ ] **Step 3: Write minimal implementation**

Append to `sheetci/formula.py`:

```python
# A structured table reference is an identifier immediately followed by a bracket
# group: Table1[Amount], Table1[[#Headers],[Amount]]. The inner alternation uses
# `+` rather than `*` so the outer quantifier can never loop on an empty match.
_TABLE_REF = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*\[(?:[^\[\]]+|\[[^\[\]]*\])*\]")

_EXTERNAL_MARKERS = ("[", ".xlsx", ".xls", "http://", "https://")


def strip_table_references(formula: str) -> str:
    """Remove structured table references so they cannot look like external links."""
    previous = None
    current = formula
    while previous != current:
        previous = current
        current = _TABLE_REF.sub(" ", current)
    return current


def has_external_reference(formula: str) -> bool:
    """True when the formula points at another workbook or a URL."""
    remaining = strip_table_references(formula).lower()
    return any(marker in remaining for marker in _EXTERNAL_MARKERS)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_formula.py -v`
Expected: PASS, 21 tests.

- [ ] **Step 5: Commit**

```bash
git add sheetci/formula.py tests/test_formula.py
git commit -m "feat(formula): tell structured table refs apart from external links"
```

---

### Task 3: Column-aggregate detection for totals rows

`INCONSISTENT_FORMULA` flags the `=SUM(E2:E41)` totals row under column `E` as deviating from its own column. Position is deliberately not part of the rule, so a totals row placed above the data behaves the same.

**Files:**
- Modify: `sheetci/formula.py`
- Test: `tests/test_formula.py`

**Interfaces:**
- Consumes: `sheetci/formula.py` from Tasks 1-2.
- Produces: `is_column_aggregate(formula: str, column_letter: str) -> bool`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_formula.py`:

```python
from sheetci.formula import is_column_aggregate


def test_sum_over_own_column_is_an_aggregate():
    assert is_column_aggregate("=SUM(E2:E41)", "E") is True


def test_whole_column_sum_is_an_aggregate():
    assert is_column_aggregate("=SUM(E:E)", "E") is True


def test_aggregate_over_a_different_column_does_not_count():
    assert is_column_aggregate("=SUM(D2:D41)", "E") is False


def test_non_aggregate_function_does_not_count():
    assert is_column_aggregate("=IF(E2>1,E3,E4)", "E") is False


def test_arithmetic_formula_does_not_count():
    assert is_column_aggregate("=E2*E3", "E") is False


def test_subtotal_counts_as_an_aggregate():
    assert is_column_aggregate("=SUBTOTAL(9,E2:E41)", "E") is True


def test_column_letter_is_case_insensitive():
    assert is_column_aggregate("=sum(e2:e41)", "E") is True
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_formula.py -k aggregate -v`
Expected: FAIL — `ImportError: cannot import name 'is_column_aggregate'`

- [ ] **Step 3: Write minimal implementation**

Append to `sheetci/formula.py`:

```python
_AGGREGATE_FUNCTIONS = {
    "SUM",
    "AVERAGE",
    "COUNT",
    "COUNTA",
    "MIN",
    "MAX",
    "MEDIAN",
    "SUBTOTAL",
}

_OUTER_FUNCTION = re.compile(r"^=\s*([A-Za-z_][A-Za-z0-9_.]*)\s*\(")


def is_column_aggregate(formula: str, column_letter: str) -> bool:
    """True when the formula aggregates a range inside its own column.

    This is the totals-row shape: a column of `=A2*B2` formulas with `=SUM(E2:E41)`
    underneath is normal spreadsheet practice, not an inconsistency.
    """
    match = _OUTER_FUNCTION.match(formula)
    if not match or match.group(1).upper() not in _AGGREGATE_FUNCTIONS:
        return False

    col = re.escape(column_letter.upper())
    cell_range = re.compile(rf"\$?{col}\$?\d+\s*:\s*\$?{col}\$?\d+", re.IGNORECASE)
    whole_column = re.compile(
        rf"(?<![A-Za-z0-9_$]){col}\s*:\s*{col}(?![A-Za-z0-9_])", re.IGNORECASE
    )
    return bool(cell_range.search(formula) or whole_column.search(formula))
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_formula.py -v`
Expected: PASS, 28 tests.

- [ ] **Step 5: Commit**

```bash
git add sheetci/formula.py tests/test_formula.py
git commit -m "feat(formula): detect aggregate formulas over their own column"
```

---

### Task 4: Finding aggregation and group-based scoring

Moves `normalize_formula` into `formula.py` (so `grouping.py` can use it without importing `scanner.py`, which would be circular), then builds the grouping layer.

**Files:**
- Modify: `sheetci/formula.py` (move `normalize_formula` and `REF_PATTERN` in)
- Create: `sheetci/grouping.py`
- Test: `tests/test_grouping.py`, `tests/test_formula.py`

**Interfaces:**
- Consumes: `sheetci/formula.py` from Tasks 1-3.
- Produces:
  - `sheetci.formula.REF_PATTERN`, `sheetci.formula.normalize_formula(formula: str, cell_row: int) -> str` (moved verbatim from `scanner.py`)
  - `sheetci.grouping.format_cell_range(cells: List[str]) -> str`
  - `sheetci.grouping.group_findings(findings: List[Dict]) -> List[Dict]` — each returned finding gains `occurrences: int`, `cells: List[str]`, `cell_range: str`; input order of first appearance is preserved
  - `sheetci.grouping.score_findings(groups: List[Dict]) -> Dict` — returns `{"risk_score": int, "status": "PASS"|"FAIL", "counts": {"critical": int, "warning": int, "info": int}}`

- [ ] **Step 1: Write the failing test**

Create `tests/test_grouping.py`:

```python
from sheetci.grouping import format_cell_range, group_findings, score_findings


def _finding(sheet, cell, rule, severity, formula):
    return {
        "sheet_name": sheet,
        "cell_address": cell,
        "rule_id": rule,
        "severity": severity,
        "formula": formula,
        "explanation": f"issue in {cell}",
        "suggested_action": "fix it",
    }


def test_contiguous_cells_render_as_a_range():
    assert format_cell_range(["I2", "I3", "I4"]) == "I2:I4"


def test_single_cell_renders_alone():
    assert format_cell_range(["B7"]) == "B7"


def test_gaps_render_as_separate_runs():
    assert format_cell_range(["I2", "I3", "I7"]) == "I2:I3, I7"


def test_copied_down_column_collapses_to_one_group():
    findings = [
        _finding("Sales", f"I{row}", "HARDCODED_NUMBER", "info", f"=IF(C{row}>150,1,0)")
        for row in range(2, 42)
    ]
    groups = group_findings(findings)
    assert len(groups) == 1
    assert groups[0]["occurrences"] == 40
    assert groups[0]["cell_range"] == "I2:I41"
    assert groups[0]["cells"][0] == "I2"
    assert groups[0]["cells"][-1] == "I41"


def test_different_patterns_stay_separate():
    findings = [
        _finding("Sales", "I2", "HARDCODED_NUMBER", "info", "=C2*150"),
        _finding("Sales", "J2", "HARDCODED_NUMBER", "info", "=C2*999"),
    ]
    assert len(group_findings(findings)) == 2


def test_same_pattern_on_different_sheets_stays_separate():
    findings = [
        _finding("Sales", "I2", "HARDCODED_NUMBER", "info", "=C2*150"),
        _finding("Ops", "I2", "HARDCODED_NUMBER", "info", "=C2*150"),
    ]
    assert len(group_findings(findings)) == 2


def test_same_pattern_different_rules_stays_separate():
    findings = [
        _finding("Sales", "I2", "HARDCODED_NUMBER", "info", "=C2*150"),
        _finding("Sales", "I2", "INCONSISTENT_FORMULA", "warning", "=C2*150"),
    ]
    assert len(group_findings(findings)) == 2


def test_sheet_level_findings_are_never_grouped():
    sheet_level = {
        "sheet_name": "Hidden",
        "cell_address": None,
        "rule_id": "HIDDEN_SHEET",
        "severity": "warning",
        "formula": None,
        "explanation": "hidden",
        "suggested_action": "check",
    }
    groups = group_findings([dict(sheet_level), dict(sheet_level)])
    assert len(groups) == 2
    assert groups[0]["occurrences"] == 1
    assert groups[0]["cells"] == []
    assert groups[0]["cell_range"] == "Sheet-Level"


def test_score_counts_groups_not_occurrences():
    findings = [
        _finding("Sales", f"I{row}", "HARDCODED_NUMBER", "info", f"=C{row}*150")
        for row in range(2, 42)
    ]
    summary = score_findings(group_findings(findings))
    assert summary["risk_score"] == 3
    assert summary["counts"]["info"] == 1
    assert summary["status"] == "PASS"


def test_any_critical_forces_fail():
    groups = group_findings([_finding("S", "A1", "BROKEN_REF", "critical", "=#REF!")])
    summary = score_findings(groups)
    assert summary["risk_score"] == 30
    assert summary["status"] == "FAIL"


def test_score_is_capped_at_100():
    findings = [
        _finding("S", f"A{row}", "BROKEN_REF", "critical", f"=#REF!+{row}00")
        for row in range(1, 9)
    ]
    summary = score_findings(group_findings(findings))
    assert summary["risk_score"] == 100
```

Append to `tests/test_formula.py`:

```python
from sheetci.formula import normalize_formula


def test_normalize_formula_moved_to_formula_module():
    assert normalize_formula("=A2*B2", 2) == "=A[0]*B[0]"
    assert normalize_formula("=$A$2*B$2", 2) == "=$A$2*B$2"
    assert normalize_formula("=A4+A6", 5) == "=A[-1]+A[1]"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_grouping.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'sheetci.grouping'`

- [ ] **Step 3a: Move `normalize_formula` into `formula.py`**

Cut `REF_PATTERN` (currently `sheetci/scanner.py:20-22`) and `normalize_formula` (currently `sheetci/scanner.py:35-53`) out of `scanner.py` and paste them into `sheetci/formula.py`, unchanged, directly below the `_IDENT_CHARS` definition:

```python
# Reference pattern for cell references (e.g., A1, $B$10, C$5)
# Ensures it is not followed by alphanumeric chars or a parenthesis (which indicates a function call like LOG10() or DEC2HEX())
REF_PATTERN = re.compile(r'(\$?)([A-Z]{1,3})(\$?)([0-9]+)(?![A-Z0-9_]|\s*\()', re.IGNORECASE)


def normalize_formula(formula: str, cell_row: int) -> str:
    """
    Normalizes a formula relative to the cell's row.
    Relative row references are converted to relative offsets (e.g. A2 -> A[0] if cell_row=2).
    """
    def replace_ref(match):
        col_abs = match.group(1)
        col = match.group(2)
        row_abs = match.group(3)
        row_str = match.group(4)

        if row_abs == '$':
            # Absolute row reference - keep as is
            return match.group(0)
        else:
            offset = int(row_str) - cell_row
            return f"{col_abs}{col}[{offset}]"

    return REF_PATTERN.sub(replace_ref, formula)
```

- [ ] **Step 3b: Write `grouping.py`**

Create `sheetci/grouping.py`:

```python
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
```

- [ ] **Step 3c: Keep `scanner.py` importable**

`scanner.py` still references `normalize_formula` and `REF_PATTERN`, and `tests/test_scanner.py` imports both `normalize_formula` and `extract_hardcoded_numbers` from `sheetci.scanner`. Replace the deleted definitions with a re-export at the top of `sheetci/scanner.py`, immediately after the existing `import openpyxl` line:

```python
from sheetci.formula import (  # noqa: F401  (re-exported for backwards compatibility)
    REF_PATTERN,
    extract_hardcoded_numbers,
    normalize_formula,
)
```

Then delete the old `extract_hardcoded_numbers` body (currently `sheetci/scanner.py:55-89`) — Task 1's version supersedes it.

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_grouping.py tests/test_formula.py -v`
Expected: PASS, 40 tests (29 in `test_formula.py`, 11 in `test_grouping.py`).

Run: `uv run pytest -q`
Expected: PASS. `test_scanner.py` and `test_cli.py` still pass at this point because grouping is not yet wired into the scanner. If `test_hardcoded_number_detector` now fails, stop — that means Task 1's suppression changed a detector result unexpectedly; investigate before continuing.

- [ ] **Step 5: Commit**

```bash
git add sheetci/formula.py sheetci/grouping.py sheetci/scanner.py tests/test_grouping.py tests/test_formula.py
git commit -m "feat(grouping): collapse repeated findings and score over groups"
```

---

### Task 5: Correct the SELF_REFERENCE detector and re-tier HARDCODED_NUMBER

Two scanner-local corrections. `SELF_REFERENCE` builds `\b\$?{col}\$?{row}\b` and searches the whole formula without checking sheet qualification, so `=Sheet2!C5` in cell `C5` is reported as circular. `HARDCODED_NUMBER` drops from `warning` to `info`.

**Files:**
- Modify: `sheetci/scanner.py:175-187` (self-reference block), `sheetci/scanner.py:166-173` (hardcoded severity)
- Test: `tests/test_scanner.py`

**Interfaces:**
- Consumes: nothing new.
- Produces: no new public names. `RULE_HARDCODED_NUMBER` findings now carry `severity == "info"`.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_scanner.py`:

```python
def test_sheet_qualified_reference_is_not_self_reference(tmp_path):
    wb_path = tmp_path / "cross_sheet.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Main"
    wb.create_sheet("Sheet2")
    ws["C5"] = "=Sheet2!C5"
    wb.save(wb_path)

    res = WorkbookScanner(str(wb_path)).scan()

    findings = [f for f in res["findings"] if f["rule_id"] == RULE_SELF_REFERENCE]
    assert findings == []


def test_sheet_qualified_absolute_reference_is_not_self_reference(tmp_path):
    wb_path = tmp_path / "cross_sheet_absolute.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Main"
    wb.create_sheet("Sheet2")
    ws["C5"] = "=Sheet2!$C$5"
    wb.save(wb_path)

    res = WorkbookScanner(str(wb_path)).scan()

    findings = [f for f in res["findings"] if f["rule_id"] == RULE_SELF_REFERENCE]
    assert findings == []


def test_absolute_self_reference_is_still_detected(tmp_path):
    wb_path = tmp_path / "absolute_self_ref.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["C5"] = "=$C$5+1"
    wb.save(wb_path)

    res = WorkbookScanner(str(wb_path)).scan()

    findings = [f for f in res["findings"] if f["rule_id"] == RULE_SELF_REFERENCE]
    assert len(findings) == 1
    assert findings[0]["cell_address"] == "C5"


def test_range_ending_on_own_cell_is_still_self_reference(tmp_path):
    wb_path = tmp_path / "range_self_ref.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["C5"] = "=SUM(A1:C5)"
    wb.save(wb_path)

    res = WorkbookScanner(str(wb_path)).scan()

    findings = [f for f in res["findings"] if f["rule_id"] == RULE_SELF_REFERENCE]
    assert len(findings) == 1
    assert findings[0]["cell_address"] == "C5"
```

And change the two existing assertions that this task invalidates:

- `tests/test_scanner.py:94` — in `test_hardcoded_number_detector`, change
  `assert findings[0]["severity"] == "warning"` to `assert findings[0]["severity"] == "info"`.
- `tests/test_scanner.py:190` — in `test_risk_score_calculation`, the scored findings become
  one critical (`=A1+1`, +30), one warning (external link, +10), and one info
  (`=C2*0.08`, +3). Change `assert res["metadata"]["risk_score"] == 50` to
  `assert res["metadata"]["risk_score"] == 43`, and update the comment block above it
  (`tests/test_scanner.py:180-184`) to read:

```python
    # 1 critical (self ref, +30), 1 warning (external link, +10), 1 info (hardcoded, +3)
    ws["A1"] = "=A1+1"  # critical: +30
    ws["B2"] = "=C2*0.08"  # info: +3
    ws["D2"] = "=[Other.xlsx]Sheet1!A1"  # warning: +10
    # Total risk score should be 30 + 10 + 3 = 43
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_scanner.py -v`
Expected: FAIL — `test_sheet_qualified_reference_is_not_self_reference` finds 1 self-reference instead of 0; `test_hardcoded_number_detector` and `test_risk_score_calculation` fail on the new expectations.

- [ ] **Step 3: Write minimal implementation**

In `sheetci/scanner.py`, change the `HARDCODED_NUMBER` severity from `SEV_WARNING` to `SEV_INFO`:

```python
                        # 4. HARDCODED_NUMBER check
                        hardcoded_nums = extract_hardcoded_numbers(formula_str)
                        if hardcoded_nums:
                            nums_str = ", ".join(str(n) for n in hardcoded_nums)
                            self.findings.append({
                                "sheet_name": sheet_name,
                                "cell_address": cell_address,
                                "rule_id": RULE_HARDCODED_NUMBER,
                                "severity": SEV_INFO,
                                "explanation": f"Formula contains hardcoded numeric constants: {formula_str} (constants: {nums_str})",
                                "suggested_action": "Move hardcoded constants to input cells or parameters to make the model dynamic."
                            })
```

Replace the self-reference pattern (currently `sheetci/scanner.py:178`) so a sheet-qualified
reference no longer matches:

```python
                        # 5. SELF_REFERENCE check
                        # The negative lookbehind keeps `=Sheet2!C5` in cell C5 from
                        # matching: that points at another sheet, not at this cell.
                        # A range endpoint such as `=SUM(A1:C5)` in C5 still matches,
                        # because that genuinely is circular.
                        col_letter = cell.column_letter
                        row_num = cell.row
                        self_ref_pattern = re.compile(
                            rf"(?<![A-Za-z0-9_!$])\$?{col_letter}\$?{row_num}\b",
                            re.IGNORECASE,
                        )
                        if self_ref_pattern.search(formula_str):
```

The lookbehind excludes `$` as well as `!` and identifier characters, and the pattern body
keeps its own optional `\$?`. That combination is what makes all four reference styles
behave: in `=$C$5` the match starts at the `$` (preceded by `=`, allowed), while in
`=Sheet2!$C$5` the `$` position is blocked by the preceding `!` and the `C` position is
blocked by the preceding `$`, so neither offset matches. Excluding letters and digits is
what keeps `AC5` and `C50` from matching for cell `C5`.

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_scanner.py -v`
Expected: PASS.

Run: `uv run pytest -q`
Expected: PASS. `test_cli.py::test_scan_broken_model` still expects `100/100` — the broken
model has three criticals (90) plus warnings, so it stays capped at 100 even with
`HARDCODED_NUMBER` demoted. If it fails, stop and report the actual score before changing
the assertion.

- [ ] **Step 5: Commit**

```bash
git add sheetci/scanner.py tests/test_scanner.py
git commit -m "fix(scanner): ignore sheet-qualified refs, demote hardcoded numbers to info"
```

---

### Task 6: Wire the formula helpers and grouping into the scanner

Replaces the inline `[`-substring external-link check and adds the totals-row exemption, then makes `scan()` return grouped findings with group-based scoring.

**Files:**
- Modify: `sheetci/scanner.py:150-160` (external link), `sheetci/scanner.py:208-240` (inconsistent formula), `sheetci/scanner.py:242-285` (scoring and return)
- Test: `tests/test_scanner.py`

**Interfaces:**
- Consumes: `sheetci.formula.has_external_reference`, `sheetci.formula.is_column_aggregate`, `sheetci.grouping.group_findings`, `sheetci.grouping.score_findings`.
- Produces: `WorkbookScanner.scan()` returns `{"metadata": {...}, "findings": [...]}` where each finding carries `occurrences`, `cells`, `cell_range`, and `formula`; `metadata["total_findings"]` and `metadata["counts"]` count groups.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_scanner.py`:

```python
def test_structured_table_reference_is_not_an_external_link(tmp_path):
    wb_path = tmp_path / "table_ref.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "=SUM(Table1[Amount])"
    wb.save(wb_path)

    res = WorkbookScanner(str(wb_path)).scan()

    findings = [f for f in res["findings"] if f["rule_id"] == RULE_EXTERNAL_LINK]
    assert findings == []


def test_totals_row_is_not_an_inconsistent_formula(tmp_path):
    wb_path = tmp_path / "totals_row.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in range(2, 8):
        ws[f"A{r}"] = r
        ws[f"B{r}"] = 10
        ws[f"C{r}"] = f"=A{r}*B{r}"
    ws["C8"] = "=SUM(C2:C7)"
    wb.save(wb_path)

    res = WorkbookScanner(str(wb_path)).scan()

    findings = [f for f in res["findings"] if f["rule_id"] == RULE_INCONSISTENT_FORMULA]
    assert findings == []


def test_repeated_pattern_is_reported_once(tmp_path):
    wb_path = tmp_path / "repeated.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in range(2, 12):
        ws[f"A{r}"] = r
        ws[f"B{r}"] = f"=A{r}*7.5"
    wb.save(wb_path)

    res = WorkbookScanner(str(wb_path)).scan()

    findings = [f for f in res["findings"] if f["rule_id"] == RULE_HARDCODED_NUMBER]
    assert len(findings) == 1
    assert findings[0]["occurrences"] == 10
    assert findings[0]["cell_range"] == "B2:B11"
    assert res["metadata"]["total_findings"] == 1
    assert res["metadata"]["counts"]["info"] == 1
    assert res["metadata"]["risk_score"] == 3
    assert res["metadata"]["status"] == "PASS"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_scanner.py -k "table_reference or totals_row or repeated_pattern" -v`
Expected: FAIL — external link found, inconsistent formula found, and `KeyError: 'occurrences'`.

- [ ] **Step 3: Write minimal implementation**

Add to the import block at the top of `sheetci/scanner.py`:

```python
from sheetci.formula import has_external_reference, is_column_aggregate
from sheetci.grouping import group_findings, score_findings
```

Record the formula on every cell-level finding. In each of the five `self.findings.append({...})`
calls inside the cell loop (`BROKEN_REF`, `EXTERNAL_LINK`, `HARDCODED_NUMBER`,
`SELF_REFERENCE`, `CACHED_ERROR`), add `"formula": formula_str,` after the
`"cell_address"` entry. In the `INCONSISTENT_FORMULA` append, add `"formula": f_str,`.
In the `HIDDEN_SHEET` append, add `"formula": None,`.

Replace the external-link condition:

```python
                        # 2. EXTERNAL_LINK check
                        if has_external_reference(formula_str):
```

Add the totals-row exemption inside the inconsistency loop, replacing the
`if norm != majority_pattern:` body's opening:

```python
                    for (cell, f_str), norm in zip(cells_info, normalized_list):
                        if norm == majority_pattern:
                            continue
                        # A totals row under (or over) the column is normal practice.
                        if is_column_aggregate(f_str, col_letter):
                            continue
                        self.findings.append({
```

Replace the whole scoring block (`sheetci/scanner.py:242-285`, from the
`# Calculate risk score` comment through the final `return`) with:

```python
        # Collapse repeated patterns before scoring: one formula copied down forty
        # rows is one problem, not forty.
        groups = group_findings(self.findings)
        summary = score_findings(groups)

        self.metadata = {
            "workbook_name": openpyxl.utils.escape.unescape(self.filepath.split("/")[-1]),
            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_sheets": total_sheets,
            "total_formulas": total_formulas,
            "total_findings": len(groups),
            "risk_score": summary["risk_score"],
            "status": summary["status"],
            "counts": summary["counts"],
        }

        return {
            "metadata": self.metadata,
            "findings": groups,
        }
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_scanner.py -v`
Expected: PASS.

Run: `uv run pytest -q`
Expected: PASS. Two assertions in `tests/test_cli.py` may now need attention:
`test_scan_json_output` asserts `len(data["findings"]) == 8`. The broken model's eight
findings all have distinct patterns, so it should still be 8. If it is not, print the
grouped findings and confirm which two collapsed before touching the assertion — an
unexpected collapse means the group key is too coarse.

- [ ] **Step 5: Commit**

```bash
git add sheetci/scanner.py tests/test_scanner.py
git commit -m "feat(scanner): adopt formula helpers and group findings before scoring"
```

---

### Task 7: Render occurrence counts in all three reporters

**Files:**
- Modify: `sheetci/cli.py:72-77`, `sheetci/reporters.py:31-38`, `sheetci/reporters.py:337-339`
- Test: `tests/test_cli.py`

**Interfaces:**
- Consumes: grouped findings from Task 6 (`occurrences`, `cell_range`).
- Produces: no new public names.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_cli.py`:

```python
def test_console_output_shows_occurrence_counts(tmp_path):
    import openpyxl

    wb_path = tmp_path / "repeated.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in range(2, 12):
        ws[f"A{r}"] = r
        ws[f"B{r}"] = f"=A{r}*7.5"
    wb.save(wb_path)

    result = runner.invoke(app, ["scan", str(wb_path)])

    assert result.exit_code == 0
    assert "B2:B11" in result.stdout
    assert "x10" in result.stdout


def test_markdown_report_shows_occurrence_counts(tmp_path):
    md_report = tmp_path / "report.md"
    result = runner.invoke(app, [
        "scan",
        "examples/broken-commission-model.xlsx",
        "--out", str(md_report),
    ])
    assert result.exit_code == 1
    content = md_report.read_text(encoding="utf-8")
    assert "**Location**:" in content
    assert "Occurrences" in content
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_cli.py -k "occurrence" -v`
Expected: FAIL — `assert "x10" in result.stdout` fails; `"Occurrences"` missing from Markdown.

- [ ] **Step 3: Write minimal implementation**

In `sheetci/cli.py`, replace the findings loop:

```python
        if findings:
            typer.echo("\nFindings Detail:")
            for f in findings:
                location = f.get("cell_range") or "[Sheet-Level]"
                occurrences = f.get("occurrences", 1)
                count = f" x{occurrences}" if occurrences > 1 else ""
                typer.echo(f"- {f['severity'].upper()} ({f['rule_id']}) in {f['sheet_name']} [{location}]{count}: {f['explanation']}")
            typer.echo("=" * 60)
```

In `sheetci/reporters.py`, replace `render_finding`:

```python
    def render_finding(finding: Dict[str, Any]) -> str:
        location = finding.get("cell_range") or "Sheet-Level"
        occurrences = finding.get("occurrences", 1)
        return (
            f"### {finding['rule_id']} ({finding['severity'].upper()})\n"
            f"- **Location**: Sheet `{finding['sheet_name']}`, Cells `{location}`\n"
            f"- **Occurrences**: {occurrences}\n"
            f"- **Issue**: {finding['explanation']}\n"
            f"- **Action**: {finding['suggested_action']}\n"
        )
```

In `HTML_TEMPLATE`, replace the `finding-meta` div:

```html
                <div class="finding-meta" style="margin-bottom: 0.5rem;">
                    Location: Sheet <strong>{{ f.sheet_name }}</strong>{% if f.cell_range and f.cell_range != 'Sheet-Level' %}, Cells <strong>{{ f.cell_range }}</strong>{% endif %}{% if f.occurrences and f.occurrences > 1 %} &middot; <strong>{{ f.occurrences }} occurrences</strong>{% endif %}
                </div>
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run pytest -q`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add sheetci/cli.py sheetci/reporters.py tests/test_cli.py
git commit -m "feat(reporters): show occurrence counts and cell ranges"
```

---

### Task 8: Realistic example workbook, end-to-end regression, version bump

The acceptance gate. Adds the healthy workbook that exposed the problem as a committed
fixture so the regression cannot silently return.

**Files:**
- Modify: `scripts/create_example_wb.py`
- Create: `examples/realistic-model.xlsx` (generated, committed)
- Modify: `tests/test_cli.py`, `pyproject.toml`, `sheetci/__init__.py`, `README.md`

**Interfaces:**
- Consumes: everything from Tasks 1-7.
- Produces: `scripts/create_example_wb.py::create_realistic_model(filepath: str) -> None`

- [ ] **Step 1: Add the workbook generator**

In `scripts/create_example_wb.py`, add above the `if __name__ == "__main__":` block:

```python
def create_realistic_model(filepath: str):
    """A healthy model written the way a competent analyst writes one.

    Nothing here is a defect: rates live on a dedicated assumptions sheet and are
    referenced absolutely, currency is rounded, and each numeric column has a
    totals row. SheetCI must pass this file.
    """
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Assumptions"
    ws["A1"] = "Parameter"
    ws["B1"] = "Value"
    ws["A2"] = "VAT rate"
    ws["B2"] = 0.20
    ws["A3"] = "Commission rate"
    ws["B3"] = 0.075
    ws["A4"] = "Discount tier"
    ws["B4"] = 0.05

    ws2 = wb.create_sheet("Sales")
    headers = ["Rep", "Region", "Units", "Unit Price", "Gross", "VAT", "Net", "Commission", "Band"]
    for i, header in enumerate(headers, start=1):
        ws2.cell(row=1, column=i, value=header)

    for r in range(2, 42):
        ws2.cell(row=r, column=1, value=f"Rep {r - 1}")
        ws2.cell(row=r, column=2, value="EMEA" if r % 2 else "AMER")
        ws2.cell(row=r, column=3, value=100 + r)
        ws2.cell(row=r, column=4, value=19.99)
        ws2.cell(row=r, column=5, value=f"=C{r}*D{r}")
        ws2.cell(row=r, column=6, value=f"=E{r}*Assumptions!$B$2")
        ws2.cell(row=r, column=7, value=f"=E{r}-F{r}")
        ws2.cell(row=r, column=8, value=f"=ROUND(G{r}*Assumptions!$B$3,2)")
        ws2.cell(row=r, column=9, value=f'=IF(C{r}>150,"High",IF(C{r}>120,"Mid","Low"))')

    ws2["E42"] = "=SUM(E2:E41)"
    ws2["G42"] = "=SUM(G2:G41)"
    ws2["H42"] = "=SUM(H2:H41)"

    ws3 = wb.create_sheet("Summary")
    ws3["A1"] = "Metric"
    ws3["B1"] = "Value"
    ws3["A2"] = "Total Gross"
    ws3["B2"] = "=Sales!E42"
    ws3["A3"] = "Total Commission"
    ws3["B3"] = "=Sales!H42"
    ws3["A4"] = "Avg Unit Price"
    ws3["B4"] = "=AVERAGE(Sales!D2:D41)"
    ws3["B5"] = '=VLOOKUP("Rep 1",Sales!A2:I41,5,FALSE)'
    ws3["B6"] = '=INDEX(Sales!G2:G41,MATCH("Rep 3",Sales!A2:A41,0))'
    ws3["B7"] = "=B3/B2*100"
    ws3["B8"] = "=EOMONTH(TODAY(),12)"

    os.makedirs(os.path.dirname(filepath), exist_ok=True)
    wb.save(filepath)
    print(f"Created realistic model: {filepath}")
```

And extend the entry point:

```python
if __name__ == "__main__":
    create_broken_model("examples/broken-commission-model.xlsx")
    create_clean_model("examples/clean-model.xlsx")
    create_realistic_model("examples/realistic-model.xlsx")
```

- [ ] **Step 2: Generate the workbook**

Run: `uv run python scripts/create_example_wb.py`
Expected: three "Created ..." lines, and `examples/realistic-model.xlsx` exists.

- [ ] **Step 3: Write the regression tests**

Append to `tests/test_cli.py`:

```python
def test_realistic_model_passes():
    """A healthy workbook must not fail. This is the regression that started this work."""
    result = runner.invoke(app, ["scan", "examples/realistic-model.xlsx", "--json"])
    assert result.exit_code == 0
    data = json.loads(result.stdout)
    assert data["metadata"]["status"] == "PASS"
    assert data["metadata"]["counts"]["critical"] == 0
    assert data["metadata"]["counts"]["warning"] == 0
    assert len(data["findings"]) <= 5


def test_broken_model_still_detects_criticals():
    """Noise reduction must not cost detection power."""
    result = runner.invoke(app, ["scan", "examples/broken-commission-model.xlsx", "--json"])
    assert result.exit_code == 1
    data = json.loads(result.stdout)
    assert data["metadata"]["status"] == "FAIL"
    rules = {f["rule_id"] for f in data["findings"]}
    assert "BROKEN_REF" in rules
    assert "SELF_REFERENCE" in rules
    assert "CACHED_ERROR" in rules
```

- [ ] **Step 4: Run the full suite**

Run: `uv run pytest -q`
Expected: PASS, all tests.

Run: `uv run sheetci scan examples/realistic-model.xlsx`
Expected: `Status: PASS`, exit code 0, at most 5 findings. The expected shape is 2 findings
and a risk score of 6 — one grouped `HARDCODED_NUMBER` for `Sales!I2:I41` and one for
`Summary!B7`. If the count is higher, do not relax the assertion; report which extra
findings appeared.

- [ ] **Step 5: Bump the version and document the example**

In `pyproject.toml` change `version = "0.1.0"` to `version = "0.2.0"`.
In `sheetci/__init__.py` change `__version__ = "0.1.0"` to `__version__ = "0.2.0"`.

In `README.md`, add below the existing Demo section:

```markdown
### Healthy workbook

`examples/realistic-model.xlsx` is a well-built commission model: rates live on an
assumptions sheet, currency is rounded, and each column has a totals row. It is included so
the opposite case is testable too.

```bash
sheetci scan examples/realistic-model.xlsx
```

Expected result:

- Status: `PASS`
- Exit code: `0`

Findings that repeat down a column are reported once, with an occurrence count and a cell
range, rather than once per row.
```

- [ ] **Step 6: Commit**

```bash
git add scripts/create_example_wb.py examples/realistic-model.xlsx tests/test_cli.py pyproject.toml sheetci/__init__.py README.md
git commit -m "test: add healthy-workbook regression and bump to 0.2.0"
```

---

## Verification

After Task 8, confirm the spec's acceptance criteria directly:

```bash
uv run pytest -q
uv run sheetci scan examples/realistic-model.xlsx
uv run sheetci scan examples/clean-model.xlsx
uv run sheetci scan examples/broken-commission-model.xlsx --out report.md --html report.html
```

Expected: suite green; realistic and clean models exit `0` with status `PASS`; broken model
exits `1` with status `FAIL` and `BROKEN_REF`, `SELF_REFERENCE`, `CACHED_ERROR` present.

`report.md` and `report.html` are gitignored — do not commit them.
