/**
 * Formula Auditor for Excel (Office Scripts)
 * Read-only against source sheets. Writes findings to a dedicated report sheet.
 */

type Severity = "Critical" | "High" | "Medium" | "Low" | "Info";
type AuditRow = [string, string, string, string, string, string, string, string, string, string];

interface AuditOptions {
  reportSheetName: string;
  includeFormulaInventory: boolean;
  flagBlankReferences: boolean;
  flagHardcodedNumbers: boolean;
  flagVolatileFunctions: boolean;
  flagExternalLinks: boolean;
  flagInconsistentFormulas: boolean;
  flagHiddenSheets: boolean;
  maxFindings: number;
}

function main(
  workbook: ExcelScript.Workbook,
  includeFormulaInventory: boolean = false,
  maxFindings: number = 10000
): string {
  const options: AuditOptions = {
    reportSheetName: "Formula Audit",
    includeFormulaInventory,
    flagBlankReferences: true,
    flagHardcodedNumbers: true,
    flagVolatileFunctions: true,
    flagExternalLinks: true,
    flagInconsistentFormulas: true,
    flagHiddenSheets: true,
    maxFindings: Math.max(100, Math.min(maxFindings, 50000))
  };

  const findings: AuditRow[] = [];
  const reportName = uniqueReportName(workbook, options.reportSheetName);
  const sheets = workbook.getWorksheets();

  for (const sheet of sheets) {
    if (sheet.getName().startsWith(options.reportSheetName)) continue;
    auditWorksheet(sheet, findings, options);
    if (findings.length >= options.maxFindings) break;
  }

  if (findings.length >= options.maxFindings) {
    addFinding(
      findings,
      options,
      "High",
      "AUDIT_LIMIT",
      "Workbook",
      "",
      "",
      "",
      "",
      "",
      `Audit stopped after ${options.maxFindings} findings. Increase maxFindings and rerun if required.`,
      "Review the current results first, then rerun with a higher limit."
    );
  }

  writeReport(workbook, reportName, findings);
  return `Formula audit complete. ${findings.length} finding(s) written to '${reportName}'.`;
}

function auditWorksheet(
  sheet: ExcelScript.Worksheet,
  findings: AuditRow[],
  options: AuditOptions
): void {
  const sheetName = sheet.getName();
  const used = sheet.getUsedRange(true);

  if (
    options.flagHiddenSheets &&
    sheet.getVisibility() !== ExcelScript.SheetVisibility.visible
  ) {
    addFinding(
      findings,
      options,
      "Info",
      "HIDDEN_SHEET",
      sheetName,
      "",
      "",
      "",
      "",
      "",
      "Worksheet is hidden.",
      "Confirm that the sheet should remain hidden and that formulas do not conceal key assumptions."
    );
  }

  if (!used) return;

  const formulas = used.getFormulas();
  const values = used.getValues();
  const texts = used.getTexts();
  const startRow = used.getRowIndex();
  const startCol = used.getColumnIndex();
  const rowCount = used.getRowCount();
  const colCount = used.getColumnCount();

  const blankMap: boolean[][] = [];
  for (let r = 0; r < rowCount; r++) {
    blankMap[r] = [];
    for (let c = 0; c < colCount; c++) {
      blankMap[r][c] = values[r][c] === "" || values[r][c] === null;
    }
  }

  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      if (findings.length >= options.maxFindings) return;

      const raw = formulas[r][c];
      if (typeof raw !== "string" || !raw.startsWith("=")) continue;

      const absoluteRow = startRow + r;
      const absoluteCol = startCol + c;
      const cell = sheet.getCell(absoluteRow, absoluteCol);
      const address = cell.getAddress().split("!").pop() || cell.getAddress();
      const displayed = texts[r][c] || "";
      const formula = raw;

      if (options.includeFormulaInventory) {
        addFinding(
          findings,
          options,
          "Info",
          "FORMULA",
          sheetName,
          address,
          formula,
          displayed,
          "",
          "",
          "Formula inventory entry.",
          "No action required."
        );
      }

      if (isErrorText(displayed)) {
        addFinding(
          findings,
          options,
          "Critical",
          "FORMULA_ERROR",
          sheetName,
          address,
          formula,
          displayed,
          "",
          "",
          `Formula evaluates to ${displayed}.`,
          "Inspect precedents, function arguments, data types, and missing names or ranges."
        );
      }

      if (options.flagExternalLinks && /\[[^\]]+\]/.test(formula)) {
        addFinding(
          findings,
          options,
          "High",
          "EXTERNAL_LINK",
          sheetName,
          address,
          formula,
          displayed,
          "",
          "",
          "Formula appears to reference another workbook.",
          "Verify the source workbook, link path, refresh behavior, and ownership."
        );
      }

      if (options.flagVolatileFunctions) {
        const volatile = findFunctions(formula, [
          "NOW",
          "TODAY",
          "RAND",
          "RANDBETWEEN",
          "OFFSET",
          "INDIRECT",
          "CELL",
          "INFO"
        ]);

        if (volatile.length > 0) {
          addFinding(
            findings,
            options,
            "Medium",
            "VOLATILE_FUNCTION",
            sheetName,
            address,
            formula,
            displayed,
            "",
            volatile.join(", "),
            `Formula uses volatile function(s): ${volatile.join(", ")}.`,
            "Confirm volatility is necessary; excessive use can slow recalculation."
          );
        }
      }

      if (options.flagHardcodedNumbers) {
        const constants = extractHardcodedNumbers(formula);
        if (constants.length > 0) {
          addFinding(
            findings,
            options,
            "Medium",
            "HARDCODED_NUMBER",
            sheetName,
            address,
            formula,
            displayed,
            "",
            constants.join(", "),
            `Formula contains numeric constant(s): ${constants.join(", ")}.`,
            "Move business assumptions to clearly labeled input cells where practical."
          );
        }
      }

      if (options.flagBlankReferences) {
        const blankRefs = findBlankReferences(
          formula,
          blankMap,
          startRow,
          startCol,
          rowCount,
          colCount
        );

        if (blankRefs.length > 0) {
          addFinding(
            findings,
            options,
            "Low",
            "BLANK_REFERENCE",
            sheetName,
            address,
            formula,
            displayed,
            blankRefs.join(", "),
            "",
            `Formula directly references blank cell(s): ${blankRefs.join(", ")}.`,
            "Confirm blanks are intentional and handled explicitly."
          );
        }
      }

      if (options.flagInconsistentFormulas) {
        const neighbors = getNeighborFormulaSignatures(formulas, r, c);
        const signature = normalizeFormula(formula);

        if (neighbors.length >= 2) {
          const matching = neighbors.filter((n) => n === signature).length;
          if (matching === 0) {
            addFinding(
              findings,
              options,
              "High",
              "INCONSISTENT_FORMULA",
              sheetName,
              address,
              formula,
              displayed,
              "",
              "",
              "Formula pattern differs from adjacent formulas.",
              "Compare this cell with neighboring rows and columns; it may be an overwritten or mis-copied formula."
            );
          }
        }
      }
    }
  }
}

function writeReport(
  workbook: ExcelScript.Workbook,
  name: string,
  findings: AuditRow[]
): void {
  const ws = workbook.addWorksheet(name);
  ws.setTabColor("#C00000");

  ws.getRange("A1:J1").merge(false);
  ws.getRange("A1").setValue("Excel Formula Audit Report");
  ws.getRange("A1").getFormat().getFont().setBold(true);
  ws.getRange("A1").getFormat().getFont().setSize(16);
  ws.getRange("A1").getFormat().getFill().setColor("#1F4E78");
  ws.getRange("A1").getFormat().getFont().setColor("#FFFFFF");

  const summary: (string | number)[][] = [
    ["Generated", new Date().toISOString()],
    ["Findings", findings.length],
    ["Critical", findings.filter((x) => x[0] === "Critical").length],
    ["High", findings.filter((x) => x[0] === "High").length],
    ["Medium", findings.filter((x) => x[0] === "Medium").length],
    ["Low", findings.filter((x) => x[0] === "Low").length],
    ["Scope", "All used cells in non-audit worksheets"],
    ["Mode", "Read-only against source sheets"]
  ];

  ws.getRangeByIndexes(2, 0, summary.length, 2).setValues(summary);
  ws.getRange("A3:A10").getFormat().getFont().setBold(true);

  const headers: string[][] = [[
    "Severity",
    "Rule",
    "Sheet",
    "Cell",
    "Formula",
    "Displayed value",
    "Referenced cells",
    "Detected items",
    "Finding",
    "Recommendation"
  ]];

  const headerRow = 12;
  ws.getRangeByIndexes(headerRow, 0, 1, 10).setValues(headers);

  const rows: AuditRow[] =
    findings.length > 0
      ? findings
      : [[
          "Info",
          "NO_FINDINGS",
          "Workbook",
          "",
          "",
          "",
          "",
          "",
          "No issues were detected by the configured rules.",
          "Review key formulas manually before relying on the workbook."
        ]];

  ws.getRangeByIndexes(headerRow + 1, 0, rows.length, 10).setValues(rows);

  const tableRange = ws.getRangeByIndexes(headerRow, 0, rows.length + 1, 10);
  const table = ws.addTable(tableRange, true);
  table.setName(safeTableName(name));
  table.setPredefinedTableStyle("TableStyleMedium2");

  const dataRange = ws.getRangeByIndexes(headerRow + 1, 0, rows.length, 10);
  dataRange.getFormat().setWrapText(true);
  dataRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);

  const severityRange = ws.getRangeByIndexes(headerRow + 1, 0, rows.length, 1);
  addSeverityConditionalFormat(severityRange, '=A14="Critical"', "#F4CCCC");
  addSeverityConditionalFormat(severityRange, '=A14="High"', "#FCE5CD");
  addSeverityConditionalFormat(severityRange, '=A14="Medium"', "#FFF2CC");

  ws.getRange("A:J").getFormat().autofitColumns();
  ws.getRange("A:A").getFormat().setColumnWidth(85);
  ws.getRange("B:B").getFormat().setColumnWidth(145);
  ws.getRange("C:D").getFormat().setColumnWidth(105);
  ws.getRange("E:E").getFormat().setColumnWidth(300);
  ws.getRange("F:H").getFormat().setColumnWidth(135);
  ws.getRange("I:J").getFormat().setColumnWidth(260);
  ws.getFreezePanes().freezeRows(headerRow + 1);
  ws.activate();
}

function addSeverityConditionalFormat(
  range: ExcelScript.Range,
  formula: string,
  fillColor: string
): void {
  const conditionalFormat = range.addConditionalFormat(
    ExcelScript.ConditionalFormatType.custom
  );
  const customFormat = conditionalFormat.getCustom();
  customFormat.getRule().setFormula(formula);
  customFormat.getFormat().getFill().setColor(fillColor);
}

function addFinding(
  rows: AuditRow[],
  options: AuditOptions,
  severity: Severity,
  rule: string,
  sheet: string,
  cell: string,
  formula: string,
  value: string,
  refs: string,
  items: string,
  finding: string,
  recommendation: string
): void {
  if (rows.length >= options.maxFindings) return;
  rows.push([
    severity,
    rule,
    sheet,
    cell,
    formula,
    value,
    refs,
    items,
    finding,
    recommendation
  ]);
}

function isErrorText(text: string): boolean {
  return /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|FIELD!|BLOCKED!|UNKNOWN!|CONNECT!)/i.test(
    text.trim()
  );
}

function findFunctions(formula: string, names: string[]): string[] {
  const upper = formula.toUpperCase();
  return names.filter((name) =>
    new RegExp("(^|[^A-Z0-9_.])" + name + "\\s*\\(", "i").test(upper)
  );
}

function extractHardcodedNumbers(formula: string): string[] {
  let cleaned = formula.replace(/"(?:[^"]|"")*"/g, "");
  cleaned = cleaned.replace(/\$?[A-Z]{1,3}\$?\d+/gi, "");
  cleaned = cleaned.replace(/\b\d+(?:\.\d+)?[Ee][+-]?\d+\b/g, "");

  const matches =
    cleaned.match(/(^|[^A-Z_])(-?\d+(?:\.\d+)?%?)(?![A-Z_])/gi) || [];

  const values = matches
    .map((x) => (x.match(/-?\d+(?:\.\d+)?%?/) || [""])[0])
    .filter((x) => x !== "" && x !== "0" && x !== "1");

  return Array.from(new Set(values)).slice(0, 10);
}

function findBlankReferences(
  formula: string,
  blankMap: boolean[][],
  startRow: number,
  startCol: number,
  rowCount: number,
  colCount: number
): string[] {
  if (formula.includes("!") || formula.includes("[")) return [];

  const refs = formula.match(/\$?[A-Z]{1,3}\$?\d+/gi) || [];
  const blanks: string[] = [];

  for (const ref of Array.from(new Set(refs))) {
    const rc = a1ToIndexes(ref);
    const r = rc[0] - startRow;
    const c = rc[1] - startCol;

    if (
      r >= 0 &&
      c >= 0 &&
      r < rowCount &&
      c < colCount &&
      blankMap[r][c]
    ) {
      blanks.push(ref);
    }
  }

  return blanks.slice(0, 20);
}

function a1ToIndexes(address: string): [number, number] {
  const clean = address.replace(/\$/g, "").toUpperCase();
  const match = clean.match(/^([A-Z]+)(\d+)$/);
  if (!match) return [-1, -1];

  let col = 0;
  for (const ch of match[1]) {
    col = col * 26 + ch.charCodeAt(0) - 64;
  }

  return [parseInt(match[2], 10) - 1, col - 1];
}

function normalizeFormula(formula: string): string {
  return formula
    .toUpperCase()
    .replace(/\$?[A-Z]{1,3}\$?\d+/g, "REF")
    .replace(/\s+/g, "");
}

function getNeighborFormulaSignatures(
  formulas: (string | number | boolean)[][],
  r: number,
  c: number
): string[] {
  const candidates: [number, number][] = [
    [r - 1, c],
    [r + 1, c],
    [r, c - 1],
    [r, c + 1]
  ];

  const result: string[] = [];

  for (const position of candidates) {
    if (
      position[0] < 0 ||
      position[1] < 0 ||
      position[0] >= formulas.length ||
      position[1] >= formulas[0].length
    ) {
      continue;
    }

    const value = formulas[position[0]][position[1]];
    if (typeof value === "string" && value.startsWith("=")) {
      result.push(normalizeFormula(value));
    }
  }

  return result;
}

function uniqueReportName(
  workbook: ExcelScript.Workbook,
  base: string
): string {
  const names = new Set(
    workbook.getWorksheets().map((sheet) => sheet.getName().toLowerCase())
  );

  if (!names.has(base.toLowerCase())) return base;

  let i = 2;
  while (names.has(`${base} ${i}`.toLowerCase())) i++;
  return `${base} ${i}`;
}

function safeTableName(sheetName: string): string {
  return (
    "FormulaAudit_" + sheetName.replace(/[^A-Za-z0-9_]/g, "_")
  ).substring(0, 240);
}
