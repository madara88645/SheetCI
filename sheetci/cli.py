import sys
from pathlib import Path
from typing import Optional
import typer

from sheetci import __version__
from sheetci.scanner import WorkbookScanner
from sheetci.reporters import generate_markdown, generate_html, generate_json

app = typer.Typer(help="SheetCI: CI/test/lint for Excel files.")


def _version_callback(value: bool) -> None:
    if value:
        typer.echo(f"SheetCI version {__version__}")
        raise typer.Exit()


@app.callback()
def main(
    version: bool = typer.Option(
        False,
        "--version",
        "-V",
        help="Show the SheetCI version and exit.",
        callback=_version_callback,
        is_eager=True,
    )
):
    """SheetCI: CI/test/lint for Excel files."""


@app.command(name="version")
def version_command():
    """
    Print the version of SheetCI.
    """
    typer.echo(f"SheetCI version {__version__}")

def _ensure_parent(path: Path) -> None:
    """Create the report's parent directory, so `--out reports/x.md` just works."""
    parent = path.parent
    if str(parent):
        parent.mkdir(parents=True, exist_ok=True)


@app.command(name="scan")
def scan(
    file: Path = typer.Argument(..., help="Path to the Excel (.xlsx) file to scan.", exists=True, file_okay=True, dir_okay=False, readable=True),
    out: Optional[Path] = typer.Option(None, "--out", "-o", help="Path to save the Markdown report."),
    html: Optional[Path] = typer.Option(None, "--html", "-t", help="Path to save the HTML report."),
    json_flag: bool = typer.Option(False, "--json", "-j", help="Print the JSON summary to stdout.")
):
    """
    Scans an Excel workbook for formula errors, dependencies, and inconsistencies.
    """
    try:
        scanner = WorkbookScanner(str(file))
        result = scanner.scan()
    except Exception as e:
        typer.echo(f"Error scanning workbook: {e}", err=True)
        raise typer.Exit(code=2)
        
    metadata = result["metadata"]
    findings = result["findings"]
    
    # Generate requested reports
    if out:
        try:
            md_content = generate_markdown(result)
            _ensure_parent(out)
            out.write_text(md_content, encoding="utf-8")
        except Exception as e:
            typer.echo(f"Error saving Markdown report: {e}", err=True)
            raise typer.Exit(code=2)
            
    if html:
        try:
            html_content = generate_html(result)
            _ensure_parent(html)
            html.write_text(html_content, encoding="utf-8")
        except Exception as e:
            typer.echo(f"Error saving HTML report: {e}", err=True)
            raise typer.Exit(code=2)
            
    # Output JSON summary to stdout if requested
    if json_flag:
        typer.echo(generate_json(result))
    else:
        # Default terminal output: Print a clean console summary
        typer.echo("=" * 60)
        typer.echo(f"SheetCI Scan Results for: {metadata['workbook_name']}")
        typer.echo(f"Timestamp: {metadata['timestamp']}")
        typer.echo("-" * 60)
        typer.echo(f"Total Sheets:        {metadata['total_sheets']}")
        typer.echo(f"Total Formula Cells: {metadata['total_formulas']}")
        typer.echo(f"Total Findings:      {metadata['total_findings']}")
        typer.echo(f"Risk Score:          {metadata['risk_score']}/100")
        typer.echo(f"Status:              {metadata['status']}")
        typer.echo("=" * 60)
        
        if findings:
            typer.echo("\nFindings Detail:")
            for f in findings:
                location = f.get("cell_range") or "Sheet-Level"
                occurrences = f.get("occurrences", 1)
                count = f" x{occurrences}" if occurrences > 1 else ""
                typer.echo(f"- {f['severity'].upper()} ({f['rule_id']}) in {f['sheet_name']} [{location}]{count}: {f['explanation']}")
            typer.echo("=" * 60)

    # Determine exit code based on risk threshold
    if metadata["status"] == "FAIL":
        raise typer.Exit(code=1)
    else:
        raise typer.Exit(code=0)

if __name__ == "__main__":
    app()
