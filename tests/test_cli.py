import json
from pathlib import Path
from typer.testing import CliRunner
from sheetci.cli import app

runner = CliRunner()

def test_scan_broken_model():
    # Scanning broken model should return exit code 1
    result = runner.invoke(app, ["scan", "examples/broken-commission-model.xlsx"])
    assert result.exit_code == 1
    assert "SheetCI Scan Results" in result.stdout
    assert "Risk Score:          100/100" in result.stdout
    assert "Status:              FAIL" in result.stdout

def test_scan_clean_model():
    # Scanning clean model should return exit code 0
    result = runner.invoke(app, ["scan", "examples/clean-model.xlsx"])
    assert result.exit_code == 0
    assert "SheetCI Scan Results" in result.stdout
    assert "Risk Score:          0/100" in result.stdout
    assert "Status:              PASS" in result.stdout

def test_scan_json_output():
    result = runner.invoke(app, ["scan", "examples/broken-commission-model.xlsx", "--json"])
    assert result.exit_code == 1
    # Verify valid JSON
    data = json.loads(result.stdout)
    assert data["metadata"]["workbook_name"] == "broken-commission-model.xlsx"
    assert data["metadata"]["status"] == "FAIL"
    assert len(data["findings"]) == 8

def test_reports_generation(tmp_path):
    md_report = tmp_path / "report.md"
    html_report = tmp_path / "report.html"
    
    result = runner.invoke(app, [
        "scan", 
        "examples/broken-commission-model.xlsx", 
        "--out", str(md_report),
        "--html", str(html_report)
    ])
    
    assert result.exit_code == 1
    assert md_report.exists()
    assert html_report.exists()
    
    md_content = md_report.read_text(encoding="utf-8")
    assert "# SheetCI Audit Report" in md_content
    assert "**Status**: **FAIL**" in md_content
    
    html_content = html_report.read_text(encoding="utf-8")
    assert "<title>SheetCI Audit Report" in html_content
    assert 'id="sheetci-report"' in html_content
