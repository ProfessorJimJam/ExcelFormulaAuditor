# Contributing

Thank you for improving Excel Formula Auditor.

## Before submitting an issue

- Confirm the issue is reproducible in the Excel Office Scripts editor.
- Remove confidential workbook data.
- Include the exact compiler or runtime message.
- Include a minimal example formula and expected result.
- State whether the workbook is being used in Excel for the web or another supported Excel client.

## Pull requests

1. Fork the repository and create a focused branch.
2. Keep source changes in `src/FormulaAuditor.ts`.
3. Avoid changes to source worksheets. The auditor should remain read-only except for its report sheet.
4. Describe new audit rules, expected false positives, and performance implications.
5. Update `README.md`, `docs/RULES.md`, and `CHANGELOG.md` when behavior changes.
6. Test with blank workbooks, formula-heavy workbooks, hidden sheets, formula errors, and existing audit sheets.

## Coding guidelines

- Use explicit Office Scripts types.
- Avoid `any`.
- Prefer bulk range reads and writes over cell-by-cell workbook calls.
- Cap potentially large outputs.
- Treat findings as review candidates rather than automatic corrections.
- Do not transmit workbook content to external services.
