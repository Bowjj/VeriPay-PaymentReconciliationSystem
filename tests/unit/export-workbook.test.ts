import ExcelJS from "exceljs";
import { beforeAll, describe, expect, test } from "vitest";
import { buildOperationalWorkbook, exportFilename, operationalWorkbookBuffer, STATUS_COLORS } from "@/lib/export/verification-workbook";
import { safeText, moneyCell } from "@/lib/export/cells";
import { captureSourceWorkbook } from "@/lib/export/source-workbook";
import { exportFixture, transactionId } from "../fixtures/export";
import { operationalFixture } from "../fixtures/operational-export";

let fixture: Awaited<ReturnType<typeof operationalFixture>>;
let payments: ExcelJS.Workbook;
let gcash: ExcelJS.Workbook;
function sheet(book: ExcelJS.Workbook, name: string) { const s = book.getWorksheet(name); if (!s) throw new Error(`Missing ${name}`); return s; }
function p(row: number, col: number) { return sheet(payments, "Payment Testing").getCell(row, col); }
beforeAll(async () => {
  fixture = await operationalFixture();
  payments = new ExcelJS.Workbook(); gcash = new ExcelJS.Workbook();
  await payments.xlsx.load((await operationalWorkbookBuffer(fixture.data, "payments")).buffer);
  await gcash.xlsx.load((await operationalWorkbookBuffer(fixture.data, "gcash")).buffer);
});
test("exactly two operational workbook kinds, retaining original worksheet names, not analytical sheets", () => {
  expect(payments.worksheets.map((s) => s.name)).toEqual(["Payment Testing", "Original Instructions"]);
  expect(gcash.worksheets.map((s) => s.name)).toEqual(["September", "October", "Cover"]);
});
test("Payment original row positions, blank rows, footer, merges and non-payment sheet remain", () => {
  expect(sheet(payments, "Payment Testing").rowCount).toBe(13);
  expect(p(12, 1).value).toBeNull(); expect(p(13, 1).value).toBe("Keep original footer");
  expect(p(13, 2).isMerged).toBe(true);
  expect(sheet(payments, "Original Instructions").getCell("B3").value).toBe("Non-payment content");
  expect(p(13, 13).value).toBeNull();
});
test("original columns stay in place; only two annotations added at the far right", () => {
  for (let col = 1; col <= 12; col++) expect(p(1, col).value).toEqual(fixture.paymentBook.worksheets[0]?.getCell(1, col).value);
  expect(p(1, 13).text).toBe("Verification Status"); expect(p(1, 14).text).toBe("Verification Note");
  expect(sheet(payments, "Payment Testing").columnCount).toBe(14);
});
test("original Payment values and value types are unchanged (including method, notes, dates and amounts)", () => {
  const original = fixture.paymentBook.worksheets[0]; if (!original) throw new Error("Missing source");
  for (let row = 1; row <= 10; row++) for (let col = 1; col <= 12; col++) {
    if (row === 4 && col === 11) continue; // unsafe hyperlink intentionally removed, text retained
    expect(p(row, col).value).toEqual(original.getCell(row, col).value);
  }
  expect(p(2, 4).value).toBe(1300); expect(p(3, 4).value).toBe("1000");
});
test.each([[2, "VERIFIED", "VERIFIED"], [3, "NEEDS REVIEW", "NEEDS_REVIEW"], [6, "CASH", "CASH"], [7, "BANK", "BANK"]] as const)("acceptance row %s has visible %s and the required color", (row, label, status) => {
  expect(p(row, 13).text).toBe(label);
  expect(p(row, 13).fill).toMatchObject({ fgColor: { argb: STATUS_COLORS[status]?.fill } });
  expect(p(row, 13).font.color).toEqual({ argb: STATUS_COLORS[status]?.text });
});
test("original cell fills are preserved; status coloring does not cover the row", () => {
  expect(p(2, 1).fill).toMatchObject({ fgColor: { argb: "FFE2E8F0" } });
  expect(p(2, 4).fill).not.toEqual(p(2, 13).fill);
});
test("source formatting, dimensions, frozen header and date format survive", () => {
  const s = sheet(payments, "Payment Testing");
  expect(s.getColumn(1).width).toBe(32); expect(s.getRow(1).height).toBe(30);
  expect(s.views[0]).toMatchObject({ state: "frozen", ySplit: 1 });
  expect(p(2, 4).numFmt).toBe("#,##0.00"); expect(p(2, 7).numFmt).toBe("dd mmm yyyy");
});
test.each([[2, "Exact reference found"], [3, "Reference not found"], [4, "Missing reference"], [5, "Duplicate reference"], [6, "Cash payment"], [7, "Bank payment"]] as const)("row %s has a human-readable note", (row, note) => expect(p(row, 14).text).toBe(note));
test("manual-only VERIFIED retains final status and note, with no fabricated GCash customer", () => {
  expect(p(8, 13).text).toBe("VERIFIED"); expect(p(8, 14).text).toContain("Manually marked verified. Confirmed manually with admin");
  for (const s of gcash.worksheets) s.eachRow((row) => expect(row.getCell(s.columnCount).text).not.toBe("Manual Only Customer"));
});
test("GCash original worksheets, row positions, values and extra columns remain", () => {
  for (const name of ["September", "October"]) {
    const output = sheet(gcash, name); const original = sheet(fixture.gcashBook, name);
    expect(output.rowCount).toBe(6);
    for (let row = 1; row <= 6; row++) for (let col = 1; col <= 12; col++) expect(output.getCell(row, col).value).toEqual(original.getCell(row, col).value);
    expect(output.columnCount).toBe(13); expect(output.getCell(1, 13).text).toBe("Matched Customer");
  }
  expect(sheet(gcash, "Cover").getCell("A1").text).toBe("Original statement cover");
});
test("1300 payment and 1299 GCash stay unchanged with Jake Tirana linked", () => {
  const s = sheet(gcash, "September");
  expect(p(2, 4).value).toBe(1300); expect(p(2, 13).text).toBe("VERIFIED");
  expect(s.getCell("I2").value).toBe(1299); expect(s.getCell("M2").text).toBe("Jake Tirana");
});
test("leading-zero references stay text in both files", () => {
  expect(p(2, 6).value).toBe("0045276500984"); expect(p(9, 6).value).toBe("0000203966985");
  expect(sheet(gcash, "September").getCell("G2").value).toBe("0045276500984");
  expect(sheet(gcash, "October").getCell("G2").value).toBe("0000203966985");
});
test("unmatched GCash rows and ambiguous duplicates remain with blank customers", () => {
  expect(sheet(gcash, "October").getCell("G3").text).toBe("UNMATCHED");
  for (const [name, row] of [["September", 3], ["September", 4], ["October", 3], ["October", 4]] as const) expect(sheet(gcash, name).getCell(row, 13).text).toBe("");
});
test("export uses persisted row mapping even when result arrays are shuffled", async () => {
  const data = structuredClone(fixture.data); data.payments.reverse(); data.gcash.reverse();
  const result = await buildOperationalWorkbook(data, "payments");
  expect(sheet(result, "Payment Testing").getCell("M2").text).toBe("VERIFIED");
  expect(sheet(result, "Payment Testing").getCell("M3").text).toBe("NEEDS REVIEW");
});
test("source hyperlink allowlist retains safe proof, strips unsafe link without losing text", () => {
  expect(p(2, 11).hyperlink).toBe("https://example.com/proof.png");
  expect(p(4, 11).hyperlink).toBeUndefined(); expect(p(4, 11).text).toBe("Unsafe receipt");
});
test("source literal formula-looking text is preserved as a non-formula string", () => {
  expect(p(2, 8).value).toBe("=original literal text"); expect(p(2, 8).formula).toBeUndefined();
  expect(p(10, 1).value).toBe('=HYPERLINK("https://example.com")'); expect(p(10, 1).formula).toBeUndefined();
});
test("original executable formulas are frozen to cached values", () => { expect(p(13, 8).value).toBe(2); expect(p(13, 8).formula).toBeUndefined(); });
test("new matched-customer annotation is formula-safe", async () => {
  const data = structuredClone(fixture.data); const payment = data.payments[0]; const transaction = data.gcash[0];
  if (!payment || !transaction) throw new Error("Missing fixture"); payment.customer = "=CMD()"; transaction.matched_customer = payment.customer;
  const book = await buildOperationalWorkbook(data, "gcash");
  expect(sheet(book, "September").getCell("M2").value).toBe("'=CMD()");
  expect(sheet(book, "September").getCell("M2").formula).toBeUndefined();
});
test.each(["=SUM(1,2)", "+CMD", "-1+2", "@SUM(A1)", "\t=HYPERLINK(\"x\")"])("annotation safeText neutralizes %s", (value) => expect(safeText(value)).toBe(`'${value}`));
test("filenames use sanitized original filenames with the requested suffix", () => {
  expect(exportFilename(fixture.data, "payments")).toBe("Payment Testing - Verified.xlsx");
  expect(exportFilename(fixture.data, "gcash")).toBe("Gcash Testing - Matched.xlsx");
  const multiSource = structuredClone(fixture.data);
  const [first, second] = multiSource.gcash;
  if (!first || !second) throw new Error("Missing GCash fixtures");
  if (typeof first.raw_data !== "object" || first.raw_data === null || typeof second.raw_data !== "object" || second.raw_data === null) throw new Error("Missing GCash provenance");
  first.raw_data = { ...first.raw_data, filename: "GCash Main.xlsx", source_index: 0 };
  second.raw_data = { ...second.raw_data, filename: "GCash Secondary.xlsx", source_index: 1 };
  expect(exportFilename(multiSource, "gcash", 1)).toBe("GCash Secondary - Matched.xlsx");
  const data = exportFixture(); data.run.payment_filename = '../bad\r\n"name.xlsx';
  expect(exportFilename(data, "payments")).toBe("bad___name - Verified.xlsx");
});
test("legacy fallback is explicitly recovered, with no fake analytical sheets", async () => {
  const data = exportFixture(); const book = await buildOperationalWorkbook(data, "payments");
  expect(book.worksheets.map((s) => s.name)).toEqual(["Recovered Payment Records"]);
  expect(book.description).toContain("Original workbook structure/formatting");
});
test("legacy GCash retained layouts remain separate and no customer is guessed", async () => {
  const book = await buildOperationalWorkbook(exportFixture(), "gcash");
  expect(book.worksheets.map((s) => s.name)).toEqual(["September", "October"]);
  expect(sheet(book, "September").getCell("M2").text).toBe("Jake Tirana");
  expect(sheet(book, "September").getCell("M3").text).toBe("");
});
describe.each(["business_id", "verification_run_id"])("scope guard %s", (key) => {
  test.each(["payments", "gcash"] as const)("rejects foreign %s rows", async (kind) => {
    const data = exportFixture(); const row = data[kind][0]; if (row) Reflect.set(row, key, "foreign");
    await expect(buildOperationalWorkbook(data, kind)).rejects.toThrow("scope mismatch");
  });
});
test("manual-only fabricated relationship fails closed", async () => {
  const data = exportFixture(); const payment = data.payments[6]; if (payment) payment.gcash_transaction_id = transactionId(5);
  await expect(buildOperationalWorkbook(data, "gcash")).rejects.toThrow("persisted relationship");
});
test("duplicate row provenance fails instead of overwriting annotations", async () => {
  const data = structuredClone(fixture.data); const row = data.payments[1]; if (row) row.row_index = 2;
  await expect(buildOperationalWorkbook(data, "payments")).rejects.toThrow("provenance");
});
test("retention captures original bytes exactly", async () => expect((await captureSourceWorkbook(fixture.paymentBytes, "payments")).base64).toBe(fixture.paymentBytes.toString("base64")));
test.each([7, 1])("GCash transactions starting at row %s before a late header retain source positions and matches", async (startRow) => {
  const data = structuredClone(fixture.data);
  const source = new ExcelJS.Workbook(); const s = source.addWorksheet("GCASH TESTING ");
  if (startRow > 1) s.getCell("A1").value = "Statement cover information";
  const positions = [startRow, 20, 35, 80, 150, 175];
  data.gcash.forEach((g, i) => {
    const row = positions[i]; if (!row) throw new Error("Missing position");
    g.row_index = row; g.raw_data = { sheet_name: s.name };
    s.getCell(row, 1).value = "2026-09-21 08:00 AM";
    s.getCell(row, 7).value = g.reference_number;
    s.getCell(row, 9).value = Number(g.amount_decimal);
  });
  s.getCell("A149").value = "DATE AND TIME"; s.getCell("G149").value = "REF NO";
  s.getCell("I149").value = "AMOUNT"; s.getCell("A180").value = "Original footer";
  data.sources = { gcash: await captureSourceWorkbook(Buffer.from(await source.xlsx.writeBuffer()), "gcash") };
  expect(data.sources.gcash?.sheets[0]?.headerRow).toBe(149);
  const result = new ExcelJS.Workbook(); await result.xlsx.load((await operationalWorkbookBuffer(data, "gcash")).buffer);
  const output = sheet(result, s.name); const shift = startRow === 1 ? 1 : 0;
  expect(output.getCell("J1").text).toBe("Matched Customer");
  for (let row = 1; row <= 180; row++) for (let col = 1; col <= 9; col++) expect(output.getCell(row + shift, col).value).toEqual(s.getCell(row, col).value);
  expect(output.getCell(startRow + shift, 7).text).toBe("0045276500984");
  expect(output.getCell(startRow + shift, 9).value).toBe(1299);
  expect(output.getCell(startRow + shift, 10).text).toBe("Jake Tirana");
  for (const row of [20, 35, 150, 175]) expect(output.getCell(row + shift, 10).text).toBe("");
  expect(output.rowCount).toBe(180 + shift);
});
test("source row offsets and a headerless GCash sheet are mapped without rematching", async () => {
  const data = structuredClone(fixture.data);
  data.payments = data.payments.slice(0, 1); data.gcash = data.gcash.slice(0, 1);
  Object.assign(data.run, { total: 1, verified: 1, needs_review: 0, cash: 0, bank: 0 });
  const sourceP = new ExcelJS.Workbook(); const p = sourceP.addWorksheet("Offset Payments");
  p.getRow(5).getCell(2).value = "Customer"; p.getRow(5).getCell(3).value = "Amount";
  p.getCell("B6").value = "Jake Tirana"; p.getCell("C6").value = 1300;
  const payment = data.payments[0]; if (payment) payment.row_index = 6;
  const sourceG = new ExcelJS.Workbook(); sourceG.addWorksheet("Headerless").addRow(["0045276500984", 1299]);
  const row = data.gcash[0]; if (row) { row.row_index = 1; row.raw_data = { sheet_name: "Headerless" }; }
  data.sources = {
    payments: await captureSourceWorkbook(Buffer.from(await sourceP.xlsx.writeBuffer()), "payments"),
    gcash: await captureSourceWorkbook(Buffer.from(await sourceG.xlsx.writeBuffer()), "gcash"),
  };
  const outputP = await buildOperationalWorkbook(data, "payments");
  expect(sheet(outputP, "Offset Payments").getCell("D5").text).toBe("Verification Status");
  expect(sheet(outputP, "Offset Payments").getCell("D6").text).toBe("VERIFIED");
  expect(sheet(outputP, "Offset Payments").getCell("C6").value).toBe(1300);
  const outputG = await buildOperationalWorkbook(data, "gcash");
  expect(sheet(outputG, "Headerless").getCell("C1").text).toBe("Matched Customer");
  expect(sheet(outputG, "Headerless").getCell("A2").text).toBe("0045276500984");
  expect(sheet(outputG, "Headerless").getCell("C2").text).toBe("Jake Tirana");
});
test("over-limit data is rejected, not truncated", async () => {
  const data = exportFixture(); data.payments = Array.from({ length: 20001 }, () => data.payments[0]).filter((p) => p !== undefined);
  await expect(buildOperationalWorkbook(data, "payments")).rejects.toThrow("20,000");
});
test("extreme amounts can be emitted exactly as text", () => {
  const cell = new ExcelJS.Workbook().addWorksheet("Test").getCell("A1"); moneyCell(cell, "90071992547409.91"); expect(cell.value).toBe("90071992547409.91");
});
test("new annotation text length limit is enforced", () => expect(() => safeText("x".repeat(32768))).toThrow("32,767"));
