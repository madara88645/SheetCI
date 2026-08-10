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
