from sheetci.reporters import generate_html


def _result(sheet_name: str, explanation: str) -> dict:
    return {
        "metadata": {
            "workbook_name": "wb.xlsx",
            "timestamp": "2026-01-01 00:00:00",
            "total_sheets": 1,
            "total_formulas": 1,
            "total_findings": 1,
            "risk_score": 3,
            "status": "PASS",
            "counts": {"critical": 0, "warning": 0, "info": 1},
        },
        "findings": [{
            "sheet_name": sheet_name,
            "cell_address": "B1",
            "cell_range": "B1",
            "occurrences": 1,
            "formula": "=A1",
            "rule_id": "HARDCODED_NUMBER",
            "severity": "info",
            "explanation": explanation,
            "suggested_action": "Move constants to input cells.",
        }],
    }


def test_html_report_escapes_workbook_content():
    # Sheet names and formulas come from an untrusted workbook: they must render
    # as text, not as markup, in a report that gets shared with reviewers.
    html = generate_html(_result(
        "<img src=x onerror=alert(1)>",
        '=IF(A1<3,"<b>low</b>",A1*1.75)',
    ))

    assert "<img src=x onerror=alert(1)>" not in html
    assert "&lt;img src=x onerror=alert(1)&gt;" in html
    assert "<b>low</b>" not in html
    assert "&lt;b&gt;low&lt;/b&gt;" in html
