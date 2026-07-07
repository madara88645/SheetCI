import json
from typing import Dict, Any, List
from jinja2 import Template

def generate_markdown(result: Dict[str, Any]) -> str:
    metadata = result["metadata"]
    findings = result["findings"]
    
    md = []
    md.append(f"# SheetCI Audit Report: `{metadata['workbook_name']}`")
    md.append("")
    md.append(f"**Scan Timestamp**: {metadata['timestamp']}")
    md.append(f"**Status**: **{metadata['status']}**")
    md.append(f"**Risk Score**: {metadata['risk_score']}/100")
    md.append("")
    md.append("## Summary Metadata")
    md.append(f"- **Total Sheets**: {metadata['total_sheets']}")
    md.append(f"- **Total Formula Cells**: {metadata['total_formulas']}")
    md.append(f"- **Total Findings**: {metadata['total_findings']} (Critical: {metadata['counts']['critical']}, Warning: {metadata['counts']['warning']}, Info: {metadata['counts']['info']})")
    md.append("")
    
    if not findings:
        md.append("### 🎉 No issues detected. This workbook is clean!")
        return "\n".join(md)
        
    # Group findings by severity
    criticals = [f for f in findings if f["severity"] == "critical"]
    warnings = [f for f in findings if f["severity"] == "warning"]
    infos = [f for f in findings if f["severity"] == "info"]
    
    def render_finding(finding: Dict[str, Any]) -> str:
        addr = finding["cell_address"] or "Sheet-Level"
        return (
            f"### {finding['rule_id']} ({finding['severity'].upper()})\n"
            f"- **Location**: Sheet `{finding['sheet_name']}`, Cell `{addr}`\n"
            f"- **Issue**: {finding['explanation']}\n"
            f"- **Action**: {finding['suggested_action']}\n"
        )
        
    if criticals:
        md.append("## 🔴 Critical Findings")
        md.append("")
        for f in criticals:
            md.append(render_finding(f))
            
    if warnings:
        md.append("## 🟡 Warning Findings")
        md.append("")
        for f in warnings:
            md.append(render_finding(f))
            
    if infos:
        md.append("## 🔵 Info Findings")
        md.append("")
        for f in infos:
            md.append(render_finding(f))
            
    return "\n".join(md)


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="SheetCI Excel spreadsheet audit report for {{ metadata.workbook_name }}">
    <title>SheetCI Audit Report - {{ metadata.workbook_name }}</title>
    <style>
        :root {
            --bg-color: #0f172a;
            --card-bg: #1e293b;
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
            --border-color: #334155;
            --primary: #6366f1;
            --success: #10b981;
            --danger: #ef4444;
            --warning: #f59e0b;
            --info: #06b6d4;
        }
        
        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }
        
        body {
            font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--bg-color);
            color: var(--text-main);
            line-height: 1.6;
            padding: 2rem 1rem;
        }
        
        .container {
            max-width: 1000px;
            margin: 0 auto;
        }
        
        header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 1.5rem;
            margin-bottom: 2rem;
            flex-wrap: wrap;
            gap: 1rem;
        }
        
        h1 {
            font-size: 2.25rem;
            font-weight: 800;
            background: linear-gradient(135deg, #a5b4fc, #6366f1);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        
        .badge {
            padding: 0.5rem 1rem;
            border-radius: 9999px;
            font-weight: 700;
            text-transform: uppercase;
            font-size: 0.875rem;
            letter-spacing: 0.05em;
            display: inline-block;
        }
        
        .badge-pass {
            background-color: rgba(16, 185, 129, 0.15);
            color: var(--success);
            border: 1px solid var(--success);
        }
        
        .badge-fail {
            background-color: rgba(239, 68, 68, 0.15);
            color: var(--danger);
            border: 1px solid var(--danger);
        }
        
        .grid-stats {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 1.5rem;
            margin-bottom: 2.5rem;
        }
        
        .stat-card {
            background-color: var(--card-bg);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 1.5rem;
            text-align: center;
            transition: transform 0.2s, border-color 0.2s;
        }
        
        .stat-card:hover {
            transform: translateY(-2px);
            border-color: var(--primary);
        }
        
        .stat-card h3 {
            font-size: 0.875rem;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 0.5rem;
        }
        
        .stat-card p {
            font-size: 2rem;
            font-weight: 700;
        }
        
        .risk-score {
            font-size: 2.5rem;
            font-weight: 800;
        }
        
        .risk-low { color: var(--success); }
        .risk-medium { color: var(--warning); }
        .risk-high { color: var(--danger); }
        
        .section-title {
            font-size: 1.5rem;
            margin-bottom: 1.5rem;
            padding-bottom: 0.5rem;
            border-bottom: 2px solid var(--border-color);
        }
        
        .findings-list {
            display: flex;
            flex-direction: column;
            gap: 1rem;
        }
        
        .finding-card {
            background-color: var(--card-bg);
            border-left: 6px solid var(--primary);
            border-radius: 0 12px 12px 0;
            padding: 1.5rem;
            border-top: 1px solid var(--border-color);
            border-right: 1px solid var(--border-color);
            border-bottom: 1px solid var(--border-color);
        }
        
        .finding-card.severity-critical {
            border-left-color: var(--danger);
        }
        .finding-card.severity-warning {
            border-left-color: var(--warning);
        }
        .finding-card.severity-info {
            border-left-color: var(--info);
        }
        
        .finding-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 0.75rem;
            flex-wrap: wrap;
            gap: 0.5rem;
        }
        
        .finding-title {
            font-weight: 700;
            font-size: 1.125rem;
        }
        
        .finding-meta {
            font-size: 0.875rem;
            color: var(--text-muted);
        }
        
        .finding-detail {
            background: rgba(15, 23, 42, 0.4);
            border-radius: 6px;
            padding: 0.75rem;
            margin-bottom: 0.75rem;
            font-family: monospace;
            font-size: 0.9rem;
            border: 1px solid rgba(255, 255, 255, 0.05);
            word-break: break-all;
        }
        
        .finding-action {
            font-size: 0.9rem;
            background: rgba(99, 102, 241, 0.08);
            border: 1px dashed rgba(99, 102, 241, 0.3);
            border-radius: 6px;
            padding: 0.75rem;
        }
        
        .finding-action strong {
            color: #a5b4fc;
        }
        
        .empty-state {
            text-align: center;
            padding: 4rem 2rem;
            background-color: var(--card-bg);
            border: 1px dashed var(--border-color);
            border-radius: 12px;
        }
        
        .empty-state h2 {
            color: var(--success);
            margin-bottom: 1rem;
        }
        
        footer {
            margin-top: 4rem;
            text-align: center;
            font-size: 0.875rem;
            color: var(--text-muted);
            border-top: 1px solid var(--border-color);
            padding-top: 1.5rem;
        }
    </style>
</head>
<body>
    <div class="container" id="sheetci-report">
        <header>
            <div>
                <h1>SheetCI Audit Report</h1>
                <p style="color: var(--text-muted); font-size: 0.95rem; margin-top: 0.25rem;">Workbook: <strong>{{ metadata.workbook_name }}</strong></p>
            </div>
            <div>
                {% if metadata.status == 'PASS' %}
                <span class="badge badge-pass" id="status-badge">PASS</span>
                {% else %}
                <span class="badge badge-fail" id="status-badge">FAIL</span>
                {% endif %}
            </div>
        </header>
        
        <div class="grid-stats">
            <div class="stat-card">
                <h3>Risk Score</h3>
                <p class="risk-score {% if metadata.risk_score < 30 %}risk-low{% elif metadata.risk_score < 70 %}risk-medium{% else %}risk-high{% endif %}">
                    {{ metadata.risk_score }}<span style="font-size: 1.25rem; font-weight:400; color:var(--text-muted);">/100</span>
                </p>
            </div>
            <div class="stat-card">
                <h3>Total Sheets</h3>
                <p>{{ metadata.total_sheets }}</p>
            </div>
            <div class="stat-card">
                <h3>Total Formula Cells</h3>
                <p>{{ metadata.total_formulas }}</p>
            </div>
            <div class="stat-card">
                <h3>Total Findings</h3>
                <p>{{ metadata.total_findings }}</p>
            </div>
        </div>
        
        {% if not findings %}
        <div class="empty-state">
            <h2>🎉 No Issues Detected</h2>
            <p>This workbook is completely clean. All checks passed successfully.</p>
        </div>
        {% else %}
        
        <h2 class="section-title">Audit Findings ({{ findings|length }})</h2>
        <div class="findings-list">
            {% for f in findings %}
            <div class="finding-card severity-{{ f.severity }}">
                <div class="finding-header">
                    <span class="finding-title">{{ f.rule_id }}</span>
                    <span class="badge {% if f.severity == 'critical' %}badge-fail{% else %}badge-pass{% endif %}" style="padding: 0.25rem 0.5rem; font-size: 0.75rem;">
                        {{ f.severity }}
                    </span>
                </div>
                <div class="finding-meta" style="margin-bottom: 0.5rem;">
                    Location: Sheet <strong>{{ f.sheet_name }}</strong>{% if f.cell_address %}, Cell <strong>{{ f.cell_address }}</strong>{% endif %}
                </div>
                <div class="finding-detail">
                    {{ f.explanation }}
                </div>
                <div class="finding-action">
                    <strong>Suggested Action:</strong> {{ f.suggested_action }}
                </div>
            </div>
            {% endfor %}
        </div>
        {% endif %}
        
        <footer>
            <p>Generated by SheetCI on {{ metadata.timestamp }}</p>
        </footer>
    </div>
</body>
</html>
"""

def generate_html(result: Dict[str, Any]) -> str:
    template = Template(HTML_TEMPLATE)
    return template.render(
        metadata=result["metadata"],
        findings=result["findings"]
    )

def generate_json(result: Dict[str, Any]) -> str:
    return json.dumps(result, indent=2)
