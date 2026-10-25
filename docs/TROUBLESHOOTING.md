# Troubleshooting

## `addCustom is not a function`

Use the current `src/FormulaAuditor.ts`. Version 1.1 creates the conditional format with `range.addConditionalFormat(ExcelScript.ConditionalFormatType.custom)` and then configures the custom rule.

## The script creates many findings

Run with `includeFormulaInventory` set to `false`, review Critical and High findings first, and lower `maxFindings` for an initial pass.

## A formula is incorrectly flagged as inconsistent

The check compares normalized formulas with immediate neighbors. Irregular layouts, subtotal rows, transition periods, and exception formulas can be legitimate differences.

## A hard-coded number is intentional

Treat the result as a review prompt. The auditor does not alter or automatically replace any formula.

## The report sheet name already exists

The script creates a unique name such as `Formula Audit 2`.
