# Excel Formula Auditor

A production-oriented Microsoft Excel Office Script that scans workbook formulas and creates a structured, filterable audit report without changing the source worksheets.

## Features

- Detects formula errors such as `#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, and `#SPILL!`
- Flags external workbook links
- Identifies volatile functions, including `OFFSET`, `INDIRECT`, `NOW`, `TODAY`, and `RAND`
- Finds hard-coded numeric constants in formulas
- Flags direct same-sheet references to blank cells
- Detects formula patterns that differ from adjacent cells
- Identifies hidden worksheets
- Supports an optional complete formula inventory
- Classifies findings as Critical, High, Medium, Low, or Info
- Produces a formatted Excel table with recommendations and conditional formatting
- Limits findings to protect performance on large workbooks

## Installation

1. Open a workbook in Excel where Office Scripts is available.
2. On the **Automate** tab, select **New Script**.
3. Delete the generated sample code.
4. Copy the complete contents of [`src/FormulaAuditor.ts`](src/FormulaAuditor.ts).
5. Paste the code into the Office Scripts editor.
6. Save it as **Formula Auditor**.
7. Select **Run**.

The script creates a new worksheet named **Formula Audit**. If that name already exists, it creates **Formula Audit 2**, **Formula Audit 3**, and so on.

## Parameters

The `main` function exposes two optional parameters:

```typescript
function main(
  workbook: ExcelScript.Workbook,
  includeFormulaInventory: boolean = false,
  maxFindings: number = 10000
): string
```

- `includeFormulaInventory`: When `true`, every formula is included as an informational result.
- `maxFindings`: Maximum number of findings. Values are constrained to a range of 100 to 50,000.

## Audit rules

| Rule | Default severity | Description |
|---|---:|---|
| `FORMULA_ERROR` | Critical | Formula displays a recognized Excel error value. |
| `EXTERNAL_LINK` | High | Formula appears to reference another workbook. |
| `INCONSISTENT_FORMULA` | High | Formula pattern differs from neighboring formulas. |
| `VOLATILE_FUNCTION` | Medium | Formula uses a volatile function. |
| `HARDCODED_NUMBER` | Medium | Formula contains a numeric constant other than 0 or 1. |
| `BLANK_REFERENCE` | Low | Formula directly references a blank cell on the same sheet. |
| `HIDDEN_SHEET` | Info | Worksheet is hidden. |
| `FORMULA` | Info | Optional formula inventory entry. |

## Output columns

The report includes Severity, Rule, Sheet, Cell, Formula, Displayed value, Referenced cells, Detected items, Finding, and Recommendation.

## Important limitations

- Results are audit candidates, not proof that a workbook is incorrect.
- Blank-reference checks cover direct same-sheet A1 references only.
- Formula consistency is based on normalized adjacent formulas and can produce false positives in irregular layouts.
- Cross-sheet dependency graphs, structured references, named formulas, LAMBDA internals, and dynamic-array lineage are not fully traced.
- Numeric constants may be intentional and should be reviewed in context.
- This project is an Office Script. A persistent task pane, custom ribbon, centralized add-in deployment, or AppSource package requires an Office Add-in built with Office.js.

## Privacy and safety

The script runs against the open workbook. It writes results only to a new report worksheet and does not intentionally modify source worksheets. Review formulas before sharing audit reports because formulas may contain workbook names, sheet names, business assumptions, and other sensitive information.

## Repository structure

```text
excel-formula-auditor/
├── .github/
│   └── ISSUE_TEMPLATE/
├── docs/
│   ├── RULES.md
│   └── TROUBLESHOOTING.md
├── src/
│   └── FormulaAuditor.ts
├── CHANGELOG.md
├── CODE_OF_CONDUCT.md
├── CONTRIBUTING.md
├── LICENSE
├── README.md
└── SECURITY.md
```

## Contributing

Bug reports and focused pull requests are welcome. Read [`CONTRIBUTING.md`](CONTRIBUTING.md) before submitting changes.

## License

Released under the MIT License. See [`LICENSE`](LICENSE).
