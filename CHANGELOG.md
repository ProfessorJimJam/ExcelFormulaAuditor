# Changelog

All notable changes to this project are documented here.

## 1.1.0 - 2026-10-25

### Fixed

- Replaced the invalid `getConditionalFormats().addCustom(...)` calls with `Range.addConditionalFormat(ExcelScript.ConditionalFormatType.custom)`.
- Configured custom rules through `getCustom().getRule().setFormula(...)`.

### Included

- Formula-error, external-link, volatile-function, hard-coded-number, blank-reference, inconsistent-formula, and hidden-sheet checks.
- Optional formula inventory.
- Formatted report worksheet with severity highlighting.

## 1.0.0 - 2026-10-25

- Initial release.
