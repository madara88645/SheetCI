# SheetCI Initial Launch Plan

SheetCI is still an MVP. The first launch goal is not paid acquisition. The goal is to test whether developers, analysts, and spreadsheet-heavy teams understand the problem and want to try the CLI.

## Positioning

One-liner:

> SheetCI is a local-first CI guardrail for Excel workbooks. It scans `.xlsx` files for risky formulas, hidden sheets, external links, cached errors, and formula inconsistencies, then generates Markdown/HTML audit reports.

Short pitch:

> Business-critical spreadsheets often ship without tests. SheetCI brings a small CI-style safety check to Excel workbooks: scan locally, detect risky formula patterns, and produce an audit report before a workbook reaches finance, ops, or a pull request.

What to avoid claiming:

- Do not say SheetCI calculates every Excel formula.
- Do not say it automatically repairs workbooks.
- Do not say it replaces Excel, spreadsheet review, audit, or finance QA.
- Do not imply files are uploaded anywhere.

## Target Audience

- Developers who version spreadsheets in repos
- Finance/RevOps/ops analysts who maintain recurring Excel models
- Data quality and analytics engineers
- Founders/operators using Excel for pricing, budgets, or commissions
- Open-source users interested in local-first business tooling

## Launch Channels

### Day 1

- GitHub profile / pinned repo
- X / Twitter
- LinkedIn
- Relevant Discord/Slack communities where self-promotion is allowed

### Day 2-3

- Hacker News `Show HN`
- Reddit communities only if rules allow project sharing
- Indie Hackers or small builder communities

### Later

- Product Hunt after there is a stronger demo page or GitHub Action use case
- Short technical blog post: "CI for Excel files"
- Example video/GIF showing broken workbook -> audit report

## Copy-Paste Posts

### X / Twitter

```text
I built SheetCI: a local-first CLI that scans Excel workbooks for risky formulas and generates audit reports.

It catches things like:
- #REF! formulas
- direct self-references
- hidden sheets
- external workbook links
- hardcoded numbers
- inconsistent formula patterns

Think "CI for spreadsheets".

Repo: https://github.com/madara88645/SheetCI
```

### LinkedIn

```text
I released a small open-source MVP called SheetCI.

The idea is simple: many business-critical Excel files contain important logic, but they rarely get the same checks as software code.

SheetCI is a local-first CLI that scans `.xlsx` workbooks and produces Markdown/HTML audit reports for risky spreadsheet patterns:

- broken references
- self-referencing formulas
- cached calculation errors
- hidden sheets
- external workbook links
- hardcoded values inside formulas
- inconsistent formulas in repeated columns

It does not upload files or use an LLM. The MVP is intentionally deterministic and local.

GitHub: https://github.com/madara88645/SheetCI
```

### Hacker News / Show HN

```text
Show HN: SheetCI - local-first lint and audit reports for Excel workbooks

I built an MVP CLI tool that scans `.xlsx` files for risky spreadsheet patterns and generates Markdown/HTML audit reports.

It detects broken references, self-references, hidden sheets, external workbook links, cached formula errors, hardcoded numbers, and simple inconsistent formula patterns.

The current version is intentionally small and deterministic: no cloud upload, no LLM API, no formula repair, no full Excel calculation engine.

Repo: https://github.com/madara88645/SheetCI
```

### GitHub Repo Description

```text
Local-first CI guardrail for Excel workbooks: scan .xlsx files for risky formulas and generate audit reports.
```

## Feedback Questions

Ask early users:

1. Would you use this in a repo or only as a local CLI?
2. Which detector is most useful: broken refs, hidden sheets, hardcoded numbers, external links, or formula inconsistency?
3. What would make this trustworthy enough for a finance/ops workflow?
4. Would GitHub PR comments be useful?
5. What spreadsheet patterns create the most real-world pain for you?

## Success Signals

Minimum useful signal:

- 10+ GitHub stars from real users
- 3+ people try the CLI
- 2+ concrete detector or workflow suggestions
- 1 real workbook class/example identified by a user

Strong signal:

- Someone asks for GitHub Action PR comments
- Someone asks for config/ignore rules
- Someone shares a real spreadsheet QA story
- Someone opens an issue with a sample workbook pattern

## Next Build After Launch

Build only after feedback:

1. GitHub Action PR comment mode
2. `sheetci.toml` config for ignored constants/rules
3. Workbook diff/regression mode
4. Better formula parsing with fewer false positives

Avoid building add-ons, dashboards, repair automation, or LLM patching until there is clear pull from users.
