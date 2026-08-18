# SheetCI Design: False-Positive Reduction

- **Date**: 2026-08-10
- **Status**: Approved, ready for implementation planning
- **Scope**: Detector accuracy and finding aggregation. No new detectors.

## Problem

SheetCI fails healthy workbooks. This was measured, not assumed.

A synthetic but realistic sales/commission model was built with nothing wrong in it:
40 data rows, rates stored in a dedicated `Assumptions` sheet and referenced absolutely,
`ROUND` on currency, a `VLOOKUP` and an `INDEX`/`MATCH` on the summary sheet, and a totals
row under each numeric column. This is how a competent analyst writes a model.

Scanning it with the current `main` produces:

```
Total Formula Cells: 210
Total Findings:       46
Risk Score:          100/100
Status:              FAIL

  HARDCODED_NUMBER:      43
  INCONSISTENT_FORMULA:   3
```

All 46 findings are false positives. They come from three causes:

1. **No aggregation.** One formula copied down 40 rows produces 40 identical findings.
   The `Sales!I` quarter-bucket column alone accounts for 40 of the 43 `HARDCODED_NUMBER`
   hits, all reporting the same two numbers.
2. **`HARDCODED_NUMBER` cannot tell a business constant from a structural one.**
   `extract_hardcoded_numbers` strips references and identifiers, then flags every
   remaining literal except `0`, `1`, `2`. So the `5` in `VLOOKUP(...,5,FALSE)` (a column
   index) and the `12` in `EOMONTH(...,12)` (a month count) are reported as magic numbers.
3. **Two rules fire on ordinary spreadsheet idioms.** `INCONSISTENT_FORMULA` flags the
   `=SUM(E2:E41)` totals row under a column as deviating from its own column.
   `EXTERNAL_LINK` flags any formula containing `[`, which includes Excel structured table
   references such as `=SUM(Table1[Amount])`.

The scoring model then guarantees failure: each warning adds 10 and the workbook fails at
70, so any real model crosses the threshold on noise alone. A user who ran SheetCI on their
own file saw a wall of meaningless warnings and a red FAIL. This is a sufficient
explanation for zero adoption on its own, independent of reach.

A third defect was found by reading the code but is not yet covered by a failing example:
`SELF_REFERENCE` builds the pattern `\b\$?{col}\$?{row}\b` and searches the whole formula
without checking sheet qualification, so `=Sheet2!C5` placed in cell `C5` of another sheet
is reported as a circular reference.

## Goals

- A healthy workbook scans to **PASS** with a small number of defensible findings.
- The intentionally broken example workbook **still FAILs**, with its critical findings intact.
- Repeated instances of one underlying issue are reported once.

## Non-goals

- No `sheetci.toml` configuration file. Configuration is the natural follow-up, but tuning
  knobs cannot substitute for correct defaults: nobody configures a tool before it has been
  useful once.
- No new detectors, no formula evaluation engine, no repair, no LLM.
- No GitHub Action or PR-comment mode. That work is only worth doing on top of an engine
  that does not spam.

## Design

### 1. Finding aggregation

The scanner keeps emitting one finding per cell. Aggregation happens in a separate layer
between scanning and reporting, so per-cell detail stays available to JSON consumers and
the grouping logic stays independently testable.

- **Group key**: `(sheet_name, rule_id, normalized_formula)`, reusing the existing
  `normalize_formula` helper, which already rewrites relative row references to offsets and
  therefore makes a copied-down column collapse to a single pattern.
- **Sheet-level findings** (`HIDDEN_SHEET`, any finding with `cell_address is None`) are
  never grouped.
- **Group representation**: one finding object carrying `occurrences` (int) and `cells`
  (the full ordered list of cell addresses) plus a human-readable `cell_range` summarising
  contiguous runs, e.g. `I2:I41`. Reports render the range and the count; the full cell
  list stays in JSON.

**Scoring**: the risk score is computed over groups, not occurrences. Forty repetitions of
one pattern contribute one finding's weight. Without this the score is unchanged and the
aggregation is cosmetic. Occurrence count deliberately does not amplify the score — a
volume multiplier would reintroduce exactly the threshold-tuning problem this work exists
to remove.

### 2. Severity re-tiering

`HARDCODED_NUMBER` moves from `warning` (+10) to `info` (+3).

A constant embedded in a formula is a maintainability smell worth surfacing, not evidence
that the workbook is wrong. At `info` it can no longer drive a FAIL on its own, which is
the intended behavioural change: a model with many embedded constants should be reported,
not blocked.

`INCONSISTENT_FORMULA` stays at `warning`. A single formula deviating from its column is
the classic "someone typed over a copied-down cell" defect and is the most valuable signal
the tool produces. It is made more precise below, not weakened.

### 3. Argument-aware constant extraction

`extract_hardcoded_numbers` is replaced with a small left-to-right formula walker that
tracks, for each numeric literal, the **innermost** enclosing function name and the
literal's 1-based argument position within that function. Literals in
structurally-determined argument positions are suppressed; everything else is reported as
today.

Innermost is what makes nesting behave: in `=ROUND(x*1.15,2)` the `2` is `ROUND` argument 2
and is suppressed, while `1.15` sits in argument 1 and is reported. In
`=VLOOKUP(a,r,MATCH("x",B1:Z1,0),FALSE)` the `0` is `MATCH` argument 3, suppressed by the
`MATCH` rule rather than the `VLOOKUP` one.

The walker must track paren depth, maintain a stack of `(function_name, arg_index)`, treat
commas only at the current depth as argument separators, and ignore commas and parens
inside string literals.

Suppressed positions:

| Function                        | Argument(s) |
| ------------------------------- | ----------- |
| `VLOOKUP`, `HLOOKUP`            | 3           |
| `MATCH`                         | 3           |
| `INDEX`                         | 2, 3        |
| `ROUND`, `ROUNDUP`, `ROUNDDOWN` | 2           |
| `EOMONTH`, `EDATE`              | 2           |
| `OFFSET`                        | 2, 3, 4, 5  |
| `LARGE`, `SMALL`                | 2           |
| `SUBTOTAL`                      | 1           |
| `LEFT`, `RIGHT`, `MID`          | 2, 3        |
| `WEEKDAY`                       | 2           |

Comparison thresholds are **not** suppressed: the `150` and `120` in
`=IF(C2>150,"High",IF(C2>120,"Mid","Low"))` are genuine business constants and remain
reportable — once, at `info`, thanks to aggregation and re-tiering.

This is the highest-risk item in the change. Malformed formulas, deep nesting, and commas
inside quoted strings all have to be handled without raising. Test weight concentrates here.

### 4. Targeted false-positive fixes

**Totals rows (`INCONSISTENT_FORMULA`).** A deviating cell is skipped when its outermost
function is an aggregate (`SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `MIN`, `MAX`, `MEDIAN`,
`SUBTOTAL`) *and* it references a range within its own column. Position is not considered,
so a totals row above the data is handled identically.

**Structured table references (`EXTERNAL_LINK`).** Structured references are stripped
before the external-link check runs. The stripping pattern handles both the simple
`Table1[Amount]` form and the nested `Table1[[#Headers],[Amount]]` form. The existing
`[`, `.xlsx`, `.xls`, `http://`, `https://` checks then run against what remains, so real
external workbook links (`[budget.xlsx]Sheet1!A1`) are still caught.

**Sheet-qualified references (`SELF_REFERENCE`).** The self-reference pattern gains a
negative lookbehind for `!` and for identifier characters, so `=Sheet2!C5` in cell `C5` no
longer matches. `=SUM(A1:C5)` in cell `C5` continues to match, because a range whose
endpoint is the cell itself is a genuine circular reference.

### 5. Example workbook and regression tests

The healthy model used to measure the problem is committed as
`examples/realistic-model.xlsx`, generated by an added function in the existing
`scripts/create_example_wb.py`. It serves as both regression fixture and README evidence
that a well-built workbook passes.

Two end-to-end regression tests anchor the work:

- `examples/realistic-model.xlsx` → status `PASS`, at most 5 findings.
- `examples/broken-commission-model.xlsx` → status `FAIL`, with its `BROKEN_REF`,
  `SELF_REFERENCE`, and `CACHED_ERROR` findings still present.

The second is not negotiable. Reducing noise while losing detection power would be
suppression, not correction.

Unit tests cover each fix independently: argument-position suppression (including nesting,
quoted commas, and malformed input), table-reference stripping versus real external links,
cross-sheet versus genuine self-reference, aggregate-over-own-column detection, and group
key construction with contiguous-range formatting.

## Data model changes

`generate_json` output gains `occurrences`, `cells`, and `cell_range` on grouped findings.
Ungrouped findings carry `occurrences: 1` and a single-entry `cells` list, so consumers see
one uniform shape. `metadata.total_findings` and the `metadata.counts` severity breakdown
both count groups rather than cells, matching the scoring model. This breaks the existing
JSON shape; acceptable at `0.1.0` with no known consumers. Version bumps to `0.2.0` in
`pyproject.toml` and `sheetci/__init__.py`.

Console, Markdown, and HTML reporters render one line per group with its occurrence count
and range instead of one line per cell.

## Acceptance criteria

1. `examples/realistic-model.xlsx` scans to `PASS` with ≤ 5 findings.
   Expected in practice: 2 findings, risk score 6 — one grouped `HARDCODED_NUMBER` for the
   `Sales!I2:I41` thresholds, one for `=B3/B2*100` on the summary sheet.
2. `examples/broken-commission-model.xlsx` still scans to `FAIL` and still reports
   `BROKEN_REF`, `SELF_REFERENCE`, and `CACHED_ERROR`.
3. `VLOOKUP(...,5,FALSE)` and `EOMONTH(...,12)` produce no `HARDCODED_NUMBER` finding.
4. `=SUM(Table1[Amount])` produces no `EXTERNAL_LINK` finding;
   `=[budget.xlsx]Sheet1!A1` still does.
5. `=Sheet2!C5` in cell `C5` produces no `SELF_REFERENCE` finding;
   `=SUM(A1:C5)` in cell `C5` still does.
6. A totals row `=SUM(E2:E41)` under column `E` produces no `INCONSISTENT_FORMULA` finding.
7. Full suite green, CI green.

## Follow-up, explicitly out of this change

- `sheetci.toml` for rule enable/disable and score weights, plus a `--fail-on` threshold flag.
- GitHub Action with PR-comment output.
- README rewrite showing the healthy-file PASS alongside the broken-file FAIL.
