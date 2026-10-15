import { loadWorkbook, workbookTraversalBudget, worksheetRows } from "@/lib/excel/workbook";

import { parseGcashRow } from "./row-parser";
import { annotationColumns } from "./annotations";
import type { GcashParseIssue, GcashParseResult, GcashParsedRow } from "./types";

export type { GcashParseIssue, GcashParseResult, GcashParsedRow } from "./types";

function headerlessRepeatedAmountColumns(rows: readonly string[][]): readonly number[] | null {
  const transactions = rows.filter((row) => /^\d{4}-\d{2}-\d{2}/.test(String(row[0] ?? "").trim()) && String(row[2] ?? "").trim() && String(row[6] ?? "").trim());
  if (transactions.length === 0) return null;

  const width = Math.max(...transactions.map((row) => row.length));
  for (let column = 0; column < width - 1; column++) {
    const repeatedAmountCount = transactions.filter((row) => {
      const value = String(row[column] ?? "").trim();
      return /^\d[\d,]*\.\d{2}$/.test(value) && value === String(row[column + 1] ?? "").trim();
    }).length;
    if (repeatedAmountCount > transactions.length / 2) return [column, column + 1];
  }
  return null;
}

export async function parseGcashWorkbook(buffer: Buffer): Promise<GcashParseResult> {
  const workbook = await loadWorkbook(buffer);
  const budget = workbookTraversalBudget(workbook);
  const rows: GcashParsedRow[] = [];
  const errors: GcashParseIssue[] = [];
  const warnings: GcashParseIssue[] = [];
  const sheetsProcessed: string[] = [];

  for (const sheet of workbook.worksheets) {
    const { raw: rawRows, formatted: formattedRows } = worksheetRows(sheet, budget);
    if (rawRows.length === 0) continue;
    sheetsProcessed.push(sheet.name);
    const sourceHeaders = formattedRows.find((row) => String(row[0] ?? "").trim().toUpperCase().startsWith("DATE AND TIME")) ?? [];
    const annotations = annotationColumns(formattedRows);
    const excluded = new Set([...annotations.customer, ...annotations.conflict]);
    const labeledAmounts = sourceHeaders.flatMap((value, column) => /^(amount|debit|credit|debit amount|credit amount)$/i.test(String(value ?? "").trim()) ? [column] : []);
    const headerlessAmounts = labeledAmounts.length === 0 ? headerlessRepeatedAmountColumns(formattedRows) : null;
    // Retain legacy fixed-layout support when financial labels are unavailable,
    // but never consume known application annotations as financial amounts.
    const legacyUnlabeled = [8, 10].filter((column) => !String(sourceHeaders[column] ?? "").trim());
    const amountColumns = [...new Set(labeledAmounts.length ? [...labeledAmounts, ...legacyUnlabeled] : headerlessAmounts ?? [8, 10])]
      .filter((column) => !excluded.has(column));

    for (let index = 0; index < rawRows.length; index++) {
      const rawRow = rawRows[index];
      if (!rawRow) continue;
      const formattedRow = formattedRows[index] ?? [];
      const rowIndex = index + 1;
      const result = parseGcashRow(rawRow, formattedRow, amountColumns, headerlessAmounts !== null);

      switch (result.kind) {
        case "skip":
          break;
        case "unsupported":
          warnings.push({ rowIndex, sheetName: sheet.name, field: "row", message: result.message });
          break;
        case "error":
          errors.push({ rowIndex, sheetName: sheet.name, field: result.field, message: result.message });
          break;
        case "transaction":
          rows.push({ rowIndex, sheetName: sheet.name, ...result.data, raw: {
            ...result.data.raw,
            export_source: { headers: sourceHeaders, cells: formattedRow, reference_column: "col0" in result.data.raw ? null : 6 },
          } });
          if (result.data.warning) {
            warnings.push({
              rowIndex,
              sheetName: sheet.name,
              field: "referenceNumber",
              message: result.data.warning,
            });
          }
          break;
      }
    }
  }

  return {
    rows,
    errors,
    warnings,
    totalRows: rows.length + errors.length + warnings.filter((issue) => issue.field === "row").length,
    validCount: rows.length,
    errorCount: errors.length,
    warningCount: warnings.length,
    sheetsProcessed,
  };
}
