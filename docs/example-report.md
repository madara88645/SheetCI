# SheetCI Audit Report: `broken-commission-model.xlsx`

**Scan Timestamp**: 2026-07-07 13:31:38
**Status**: **FAIL**
**Risk Score**: 100/100

## Summary Metadata
- **Total Sheets**: 2
- **Total Formula Cells**: 11
- **Total Findings**: 8 (Critical: 3, Warning: 5, Info: 0)

## 🔴 Critical Findings

### SELF_REFERENCE (CRITICAL)
- **Location**: Sheet `Commissions`, Cell `A1`
- **Issue**: Formula contains a circular self-reference to its own cell: =A1+1
- **Action**: Refactor the formula to avoid referencing its own cell coordinate.

### BROKEN_REF (CRITICAL)
- **Location**: Sheet `Commissions`, Cell `E2`
- **Issue**: Formula contains broken reference (#REF!): =B2+#REF!
- **Action**: Fix the formula reference pointing to a deleted cell or range.

### CACHED_ERROR (CRITICAL)
- **Location**: Sheet `Commissions`, Cell `H2`
- **Issue**: Cell has a cached calculation error value: #DIV/0!
- **Action**: Review the cell's inputs or formula structure to resolve the evaluation error.

## 🟡 Warning Findings

### HARDCODED_NUMBER (WARNING)
- **Location**: Sheet `Commissions`, Cell `F2`
- **Issue**: Formula contains hardcoded numeric constants: =B2*0.08 (constants: 0.08)
- **Action**: Move hardcoded constants to input cells or parameters to make the model dynamic.

### EXTERNAL_LINK (WARNING)
- **Location**: Sheet `Commissions`, Cell `G2`
- **Issue**: Formula references an external workbook or URL: =[ExternalSource.xlsx]Sheet1!A1
- **Action**: Ensure the external link is accessible, secure, and intended.

### HARDCODED_NUMBER (WARNING)
- **Location**: Sheet `Commissions`, Cell `D6`
- **Issue**: Formula contains hardcoded numeric constants: =B6*C6+100 (constants: 100.0)
- **Action**: Move hardcoded constants to input cells or parameters to make the model dynamic.

### INCONSISTENT_FORMULA (WARNING)
- **Location**: Sheet `Commissions`, Cell `D6`
- **Issue**: Formula is inconsistent with neighboring cells in column D: =B6*C6+100 (expected pattern similar to: =B[0]*C[0])
- **Action**: Check if the formula was modified intentionally or copied down incorrectly.

### HIDDEN_SHEET (WARNING)
- **Location**: Sheet `HiddenArchive`, Cell `Sheet-Level`
- **Issue**: Worksheet 'HiddenArchive' is hidden or very hidden.
- **Action**: Verify if the hidden sheet contains deprecated calculations or sensitive data.
