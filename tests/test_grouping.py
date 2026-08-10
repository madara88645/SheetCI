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
