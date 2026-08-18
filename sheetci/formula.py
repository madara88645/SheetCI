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


def _innermost_call(stack: List[List]) -> Tuple[Optional[str], int]:
    """Return the nearest enclosing function call and the argument index within it.

    A parenthesis with no identifier in front of it groups an expression rather than
    opening a call, so those frames are skipped: in `=ROUND((B2*1.15),2)` the 1.15
    still belongs to ROUND argument 1. Returns (None, 0) when nothing encloses it.
    """
    for name, arg_index in reversed(stack):
        if name:
            return name, arg_index
    return None, 0


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

            func_name, arg_index = _innermost_call(stack)
            yield value, func_name, arg_index
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
