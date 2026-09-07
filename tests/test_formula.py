import pytest
from sheetci.formula import (
    extract_hardcoded_numbers,
    has_external_reference,
    iter_numeric_literals,
    strip_text_literals,
)


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


from sheetci.formula import normalize_formula


def test_normalize_formula_moved_to_formula_module():
    assert normalize_formula("=A2*B2", 2) == "=A[0]*B[0]"
    assert normalize_formula("=$A$2*B$2", 2) == "=$A$2*B$2"
    assert normalize_formula("=A4+A6", 5) == "=A[-1]+A[1]"


def test_grouping_parentheses_do_not_hide_the_enclosing_function():
    # The inner parens group an expression, they are not a call, so 1.15 is still
    # ROUND argument 1 rather than being attributed to an anonymous frame.
    assert list(iter_numeric_literals("=ROUND((B2*1.15),2)")) == [
        (1.15, "ROUND", 1),
        (2.0, "ROUND", 2),
    ]


def test_structural_suppression_survives_grouping_parentheses():
    # The column index is structural whether or not it is wrapped in parens.
    assert extract_hardcoded_numbers("=VLOOKUP(A1,R,(5),FALSE)") == []


def test_top_level_grouping_parentheses_report_no_function():
    # With no enclosing call the function name is None, so the index must be 0.
    assert list(iter_numeric_literals("=(A1*100)")) == [(100.0, None, 0)]


def test_text_literal_punctuation_is_not_an_external_link():
    # A custom number format or a bracketed word inside a text literal is not a
    # workbook path, so it must not raise EXTERNAL_LINK.
    assert has_external_reference('=TEXT(A1,"[$-409]#,##0.00")') is False
    assert has_external_reference('=IF(A1>0,"[draft]","ok")') is False
    assert has_external_reference('=CONCATENATE("archive.xlsx is stale")') is False


def test_real_external_references_are_still_detected():
    # Workbook paths live outside text literals (single quotes wrap the path),
    # and a URL stays detectable even though it is written as a text literal.
    assert has_external_reference("='C:\\Models\\[budget.xlsx]Sheet1'!A1") is True
    assert has_external_reference("=[1]Sheet1!B2") is True
    assert has_external_reference('=HYPERLINK("https://example.com","x")') is True


def test_strip_text_literals_keeps_sheet_quoting():
    assert strip_text_literals('=A1&"text"') == "=A1& "
    # A doubled quote escapes a quote inside the literal, it does not end it.
    assert strip_text_literals('=IF(A1="a""b","[x]","")&B1') == "=IF(A1= , , )&B1"


@pytest.mark.parametrize("formula", [
    '=INDIRECT("[budget.xlsx]Sheet1!A1")',
    '=indirect ("[1]Sheet1!A1")',
    '=SUM(INDIRECT("[budget.xlsx]Sheet1!"&A1))',
    '=INDIRECT(CONCAT("[budget.xlsx]", "Sheet1!A1"))',
    '=INDIRECT("[budget.xlsx]Sheet1!R1C1",FALSE)',
])
def test_indirect_external_workbook_literals_are_detected(formula):
    assert has_external_reference(formula)


@pytest.mark.parametrize("formula", [
    '=INDIRECT("Sheet1!A1")',
    '=IF(A1,"[draft]",INDIRECT("Sheet1!A1"))',
    '=CONCAT("INDIRECT(","archive.xlsx")',
    '=INDIRECT("Sheet1!A1",IF(A1="archive.xlsx",TRUE,FALSE))',
])
def test_indirect_does_not_turn_unrelated_text_into_external_links(formula):
    assert not has_external_reference(formula)
