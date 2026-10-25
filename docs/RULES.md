# Audit Rule Reference

## FORMULA_ERROR

Detects displayed values beginning with recognized Excel error tokens. Review precedents, names, types, and function arguments.

## EXTERNAL_LINK

Detects formulas containing bracketed workbook references. Confirm ownership, availability, refresh behavior, and portability.

## VOLATILE_FUNCTION

Detects configured volatile functions. Volatility is not automatically wrong, but repeated use can affect recalculation behavior.

## HARDCODED_NUMBER

Detects numeric constants in formulas after quoted strings, cell references, scientific notation, zero, and one are excluded. Review whether the constant represents a business assumption that belongs in a labeled input cell.

## BLANK_REFERENCE

Detects direct same-sheet A1 references that point to blank cells inside the used range. Cross-sheet references and ranges are not fully traced.

## INCONSISTENT_FORMULA

Normalizes A1 cell references and compares the formula with its immediate vertical and horizontal neighbors. Irregular models may create false positives.

## HIDDEN_SHEET

Records worksheets that are not visible. This is informational and does not imply an error.
