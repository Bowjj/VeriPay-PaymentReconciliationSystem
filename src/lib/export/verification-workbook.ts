import ExcelJS from "exceljs";
import { z } from "zod";
import { proofHref } from "@/lib/verification/review";
import { textCell } from "./cells";
import { type ExportData, type ExportPayment, ExportLimitError, MAX_EXPORT_CELLS, MAX_EXPORT_ROWS } from "./export-data";
import { type ExportKind, type SourceWorkbook } from "./source-workbook";
import { annotateGcashSheet, countAnnotationConflicts, existingCustomerHeaderRow } from "./gcash-annotations";

export const PAYMENT_ANNOTATIONS = ["Verification Status", "Verification Note"];
export const STATUS_COLORS: Record<string, { fill: string; text: string }> = {
  VERIFIED: { fill: "FFDCFCE7", text: "FF166534" },
  NEEDS_REVIEW: { fill: "FFFEE2E2", text: "FF991B1B" },
  CASH: { fill: "FFDBEAFE", text: "FF1E40AF" },
  BANK: { fill: "FFFEF9C3", text: "FF854D0E" },
  UNKNOWN: { fill: "FFFFEDD5", text: "FF9A3412" },
};
const envelopeSchema = z.object({ filename: z.string().optional(), sheet_name: z.string().optional(), source_index: z.number().int().nonnegative().optional(), source: z.record(z.string(), z.unknown()).optional() });
const gcashCellsSchema = z.object({ headers: z.array(z.unknown()), cells: z.array(z.unknown()) });

export function exportFilename(data: ExportData, kind: ExportKind, sourceIndex = 0) {
  const sourceFilename = kind === "gcash"
    ? data.gcash.find((row) => {
      const envelope = envelopeSchema.safeParse(row.raw_data);
      return envelope.success && envelope.data.source_index === sourceIndex && !!envelope.data.filename;
    })
    : undefined;
  const original = kind === "payments" ? data.run.payment_filename : envelopeSchema.parse(sourceFilename?.raw_data ?? {}).filename ?? data.run.gcash_filename;
  const fallback = kind === "payments" ? "Payment Records" : "GCash";
  const base = (original ?? fallback).replace(/^.*[\\/]/, "").replace(/\.(xlsx?|xlsm)$/i, "")
    .normalize("NFKC").replace(/[^A-Za-z0-9 _().-]/g, "_").replace(/^[ .]+|[ .]+$/g, "").slice(0, 100) || fallback;
  const safe = /^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)/i.test(base) ? `_${base}` : base;
  return `${safe} - ${kind === "payments" ? "Verified" : "Matched"}.xlsx`;
}

export function verificationNote(p: ExportPayment) {
  if (p.manual_status) return `Manually marked ${p.effective_status === "VERIFIED" ? "verified" : "needs review"}.${p.manual_note ? ` ${p.manual_note}` : ""}`;
  const notes: Record<string, string> = {
    REFERENCE_MATCH: "Exact reference found", REFERENCE_NOT_FOUND: "Reference not found",
    MISSING_REFERENCE: "Missing reference", DUPLICATE_REFERENCE: "Duplicate reference",
    DUPLICATE_PAYMENT_REFERENCE: "Duplicate payment reference", CASH_PAYMENT: "Cash payment",
    BANK_MANUAL_VERIFICATION: "Bank payment", UNKNOWN_PAYMENT_METHOD: "Unknown payment method",
  };
  return notes[p.automated_reason] ?? "Review required";
}

function validateSnapshot(data: ExportData) {
  if (data.payments.length + data.gcash.length > MAX_EXPORT_ROWS) throw new ExportLimitError("Export supports at most 20,000 combined source rows.");
  if ([...data.payments, ...data.gcash].some((r) => r.business_id !== data.run.business_id || r.verification_run_id !== data.run.id)) throw new Error("Export scope mismatch.");
  if (new Set(data.payments.map((p) => p.id)).size !== data.payments.length || new Set(data.gcash.map((g) => g.id)).size !== data.gcash.length) throw new Error("Duplicate export source IDs.");
  const counts = { total: data.payments.length, verified: 0, needs_review: 0, cash: 0, bank: 0 };
  for (const p of data.payments) {
    if (p.effective_status === "VERIFIED") counts.verified++;
    else if (p.effective_status === "NEEDS_REVIEW") counts.needs_review++;
    else if (p.effective_status === "CASH") counts.cash++;
    else if (p.effective_status === "BANK") counts.bank++;
  }
  if (Object.entries(counts).some(([key, value]) => Reflect.get(data.run, key) !== value)) throw new Error("Export snapshot counts are inconsistent.");
  const byId = new Map(data.gcash.map((g) => [g.id, g]));
  const customers = new Map<string, string>();
  for (const p of data.payments) {
    if (!p.gcash_transaction_id) continue;
    const g = byId.get(p.gcash_transaction_id);
    if (!g || g.matched_payment_id !== p.id || g.matched_customer !== p.customer || p.automated_status !== "VERIFIED"
      || p.automated_reason !== "REFERENCE_MATCH" || customers.has(g.id)) throw new Error("Inconsistent persisted relationship.");
    customers.set(g.id, p.customer);
  }
  for (const g of data.gcash) if ((g.matched_payment_id || g.matched_customer) && !customers.has(g.id)) throw new Error("Inconsistent persisted relationship.");
  return customers;
}

function sourceValue(cell: ExcelJS.Cell, value: unknown) {
  // XLSX string cells are explicitly strings, never executable formulas. Original
  // text is kept verbatim; newly written annotations use textCell/safeText instead.
  if (typeof value === "number" || typeof value === "boolean" || typeof value === "string" || value instanceof Date) cell.value = value;
  else if (value === null || value === undefined) cell.value = null;
  else cell.value = JSON.stringify(value);
}
function sanitizeWorkbook(book: ExcelJS.Workbook) {
  let cells = 0;
  for (const sheet of book.worksheets) {
    cells += sheet.rowCount * (sheet.columnCount + 2);
    if (cells > MAX_EXPORT_CELLS || sheet.columnCount > 16382) throw new ExportLimitError("Workbook exceeds the 500,000-cell export limit.");
    sheet.eachRow((row) => row.eachCell((cell) => {
      if (cell.type === ExcelJS.ValueType.Formula) sourceValue(cell, cell.result ?? null);
      else if (cell.hyperlink && !proofHref(cell.hyperlink)) cell.value = cell.text;
      // Keep original string data (including formula-looking strings) as text.
    }));
  }
}
async function loadSource(source: SourceWorkbook) {
  const bytes = Buffer.from(source.base64, "base64");
  if (!bytes.length || bytes.length > 10 * 1024 * 1024) throw new ExportLimitError("Invalid retained source size.");
  const book = new ExcelJS.Workbook();
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new Error("Retained source workbook is not an XLSX file.");
  await book.xlsx.load(new Uint8Array(bytes).buffer);
  sanitizeWorkbook(book);
  return book;
}
function addHeaders(sheet: ExcelJS.Worksheet, row: number, column: number, labels: string[]) {
  labels.forEach((label, i) => {
    const cell = sheet.getCell(row, column + i);
    if (cell.isMerged || cell.value !== null) throw new Error("Annotation would overwrite source content.");
    textCell(cell, label);
    cell.font = { bold: true, color: { argb: "FF243746" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8EEF3" } };
    cell.alignment = { wrapText: true, vertical: "middle" };
    sheet.getColumn(column + i).width = label === "Verification Note" ? 48 : 25;
  });
}
function annotatePayment(sheet: ExcelJS.Worksheet, row: number, column: number, payment: ExportPayment) {
  const cell = sheet.getCell(row, column);
  textCell(cell, payment.effective_status.replaceAll("_", " "));
  const color = STATUS_COLORS[payment.effective_status] ?? STATUS_COLORS["UNKNOWN"];
  if (!color) throw new Error("Missing status palette");
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color.fill } };
  cell.font = { bold: true, color: { argb: color.text } };
  cell.alignment = { wrapText: true, vertical: "top" };
  textCell(sheet.getCell(row, column + 1), verificationNote(payment));
  sheet.getCell(row, column + 1).alignment = { wrapText: true, vertical: "top" };
}

// Older runs never retained original files. Explicitly recovered sheets contain
// only stored fields/rows, not invented original headers, formatting or content.
function recoveredWorkbook(data: ExportData, kind: ExportKind, customers: Map<string, string>) {
  const book = new ExcelJS.Workbook();
  book.description = "Recovered from retained import data. Original workbook structure/formatting and unimported content are unavailable. Re-upload originals for source-preserving exports.";
  if (kind === "payments") {
    const sheet = book.addWorksheet("Recovered Payment Records");
    const rows = [...data.payments].sort((a, b) => a.row_index - b.row_index);
    const sources = rows.map((p) => {
      const parsed = envelopeSchema.safeParse(p.raw_data);
      return parsed.success && parsed.data.source && Object.keys(parsed.data.source).length ? parsed.data.source : {
        Customer: p.customer, Account: p.account, "Billing Period": p.billing_period, Amount: p.amount_decimal,
        Method: p.method, "Reference #": p.reference_number, "Payment Date": p.payment_date, Notes: p.notes,
        "Paid By": p.paid_by, "Received By": p.received_by, Photo: p.photo_url, "Created At": p.created_at_source,
      };
    });
    const columns = [...new Set(sources.flatMap((s) => Object.keys(s)))];
    columns.forEach((key, i) => textCell(sheet.getCell(1, i + 1), key));
    addHeaders(sheet, 1, columns.length + 1, PAYMENT_ANNOTATIONS);
    rows.forEach((p, i) => {
      columns.forEach((key, j) => sourceValue(sheet.getCell(i + 2, j + 1), /^(reference #|reference_number)$/i.test(key) ? p.reference_number : sources[i]?.[key]));
      annotatePayment(sheet, i + 2, columns.length + 1, p);
    });
  } else {
    const groups = new Map<string, typeof data.gcash>();
    for (const row of [...data.gcash].sort((a, b) => a.source_order - b.source_order)) {
      const envelope = envelopeSchema.safeParse(row.raw_data);
      const name = envelope.success ? envelope.data.sheet_name ?? "Recovered GCash" : "Recovered GCash";
      groups.set(name, [...(groups.get(name) ?? []), row]);
    }
    for (const [name, rows] of groups) {
      const sheet = book.addWorksheet(name.replace(/[:\\/?*\[\]]/g, "_").slice(0, 31));
      const cells = rows.map((g) => {
        const envelope = envelopeSchema.safeParse(g.raw_data);
        const source = envelope.success ? envelope.data.source : undefined;
        const preserved = gcashCellsSchema.safeParse(source?.["export_source"]);
        return preserved.success ? preserved.data : { headers: Object.keys(source ?? { Date: "", Description: "", Reference: "", Amount: "" }), cells: source ? Object.entries(source).map(([key, value]) => key === "ref" ? g.reference_number : value) : [g.transaction_date, g.description, g.reference_number, g.amount_decimal] };
      });
      const width = Math.max(1, ...cells.map((r) => Math.max(r.headers.length, r.cells.length)));
      for (let i = 0; i < width; i++) sourceValue(sheet.getCell(1, i + 1), cells[0]?.headers[i] ?? "");
      const matches = new Map<number, string>();
      rows.forEach((g, i) => {
        cells[i]?.cells.forEach((value, j) => sourceValue(sheet.getCell(i + 2, j + 1), value));
        const customer = customers.get(g.id);
        if (customer !== undefined) matches.set(i + 2, customer);
      });
      annotateGcashSheet(sheet, 1, width, matches);
    }
  }
  sanitizeWorkbook(book);
  return book;
}

export async function buildOperationalWorkbook(data: ExportData, kind: ExportKind, sourceIndex = 0): Promise<ExcelJS.Workbook> {
  const customers = validateSnapshot(data);
  const sourceValue = kind === "gcash" ? data.sources?.gcashSources : undefined;
  const source: SourceWorkbook | undefined = kind === "gcash" && sourceValue
    ? sourceValue.find((entry) => entry.sourceIndex === sourceIndex)?.source
    : data.sources?.[kind];
  if (!source) return recoveredWorkbook(data, kind, customers);
  const book = await loadSource(source);
  const seen = new Set<string>();
  const layouts = new Map<string, { sheet: ExcelJS.Worksheet; column: number; offset: number; header: number; matches: Map<number, string> }>();
  const first = source.sheets[0];
  for (const meta of source.sheets) {
    const sheetRows = kind === "payments" ? (meta === first ? data.payments : []) : data.gcash.filter((g) => {
      const envelope = envelopeSchema.parse(g.raw_data);
      return envelope.sheet_name === meta.name && (!Array.isArray(sourceValue) || (envelope.source_index ?? 0) === sourceIndex);
    });
    const hasRecords = kind === "payments" ? meta === first : sheetRows.length > 0;
    if (!hasRecords) continue;
    const sheet = book.getWorksheet(meta.name);
    if (!sheet) throw new Error("Missing retained source worksheet");
    const column = Math.max(sheet.columnCount, meta.lastColumn) + 1;
    const firstDataRow = Math.min(...sheetRows.map((row) => row.row_index));
    // A detected statement header may belong to a later section. Keep every
    // persisted source row mapped in place; label only the new far-right column
    // at row 1 when transactions precede that header. If row 1 is itself data,
    // insert a header row and apply the same explicit offset to every annotation.
    const annotationHeader = kind === "gcash" ? existingCustomerHeaderRow(sheet) ?? meta.headerRow : meta.headerRow;
    const lateGcashHeader = kind === "gcash" && annotationHeader !== null && annotationHeader >= firstDataRow;
    const shift = annotationHeader === null || (lateGcashHeader && firstDataRow === 1) ? 1 : 0;
    if (sheetRows.some((row) => !Number.isInteger(row.row_index) || row.row_index < 1)) throw new Error("Invalid original row provenance");
    if (shift) sheet.insertRow(1, []);
    const header = shift || lateGcashHeader ? 1 : annotationHeader ?? 1;
    if (kind === "payments") addHeaders(sheet, header, column, PAYMENT_ANNOTATIONS);
    // Parsed row indexes are absolute Excel rows. rowOffset describes the used
    // range for source metadata, not an adjustment to persisted provenance.
    layouts.set(meta.name, { sheet, column, offset: shift, header, matches: new Map() });
  }
  const rows = kind === "payments" ? data.payments : data.gcash.filter((row) => !Array.isArray(sourceValue) || envelopeSchema.parse(row.raw_data).source_index === sourceIndex);
  for (const row of rows) {
    const envelope = envelopeSchema.parse(row.raw_data);
    const name = kind === "payments" ? first?.name : envelope.sheet_name;
    const layout = name ? layouts.get(name) : undefined;
    if (!layout) throw new Error("Missing original row provenance");
    const physicalRow = row.row_index + layout.offset;
    const mapping = `${name}:${physicalRow}`;
    if (physicalRow <= layout.header || physicalRow > layout.sheet.rowCount || seen.has(mapping)) throw new Error("Invalid or ambiguous original row provenance");
    seen.add(mapping);
    if (kind === "payments" && "effective_status" in row) annotatePayment(layout.sheet, physicalRow, layout.column, row);
    else {
      const customer = customers.get(row.id);
      if (customer !== undefined) layout.matches.set(physicalRow, customer);
    }
  }
  if (kind === "gcash") for (const layout of layouts.values()) annotateGcashSheet(layout.sheet, layout.header, layout.column - 1, layout.matches);
  return book;
}
export async function operationalWorkbookBuffer(data: ExportData, kind: ExportKind, sourceIndex = 0) {
  return (await operationalWorkbookExport(data, kind, sourceIndex)).buffer;
}
export async function operationalWorkbookExport(data: ExportData, kind: ExportKind, sourceIndex = 0) {
  const book = await buildOperationalWorkbook(data, kind, sourceIndex);
  return { buffer: new Uint8Array(await book.xlsx.writeBuffer()), annotationConflicts: kind === "gcash" ? countAnnotationConflicts(book) : 0 };
}
