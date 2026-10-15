// @vitest-environment node
import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";

import { workbookBuffer } from "@/../tests/fixtures/workbook";
import { parseGcashWorkbook } from "@/lib/gcash/parser";
import { captureSourceWorkbook } from "@/lib/export/source-workbook";

const header = ["Date and Time", null, "ACCOUNTS TO", null, null, null, "REF NO", null, "AMOUNT"];
const transaction = ["2026-08-20 08:20 AM", null, "Cash in via bank", null, null, null, "0000203966985", null, "1,000.00"];

describe("parseGcashWorkbook", () => {
  it("preserves worksheet order, source rows, and leading-zero references", async () => {
    const result = await parseGcashWorkbook(await workbookBuffer([
      { name: "September", rows: [header, transaction] },
      { name: "October", rows: [header, ["2026-08-21 08:20 AM", null, "Buy Load Transaction", null, null, null, "2040920724131", null, "40.00"]] },
    ]));
    expect(result.sheetsProcessed).toEqual(["September", "October"]);
    expect(result.rows.map((row) => [row.sheetName, row.rowIndex, row.referenceNumber, row.direction])).toEqual([["September", 2, "0000203966985", "incoming"], ["October", 2, "2040920724131", "outgoing"]]);
  });

  it("uses formatted numeric references and rejects scientific notation", async () => {
    const book = new ExcelJS.Workbook(); const sheet = book.addWorksheet("GCash"); sheet.addRow(header); sheet.addRow([...transaction.slice(0, 6), 203966985, null, "100"]); sheet.getCell("G2").numFmt = "0000000000000";
    const formatted = await parseGcashWorkbook(Buffer.from(await book.xlsx.writeBuffer()));
    sheet.getCell("G2").numFmt = "0.0000E+00";
    const scientific = await parseGcashWorkbook(Buffer.from(await book.xlsx.writeBuffer()));
    expect(formatted.rows[0]?.referenceNumber).toBe("0000203966985"); expect(scientific.errors[0]?.field).toBe("referenceNumber");
  });

  it("finds late headers, ignores repeated headers and preserves annotation columns", async () => {
    const result = await parseGcashWorkbook(await workbookBuffer([{ name: "Weekly", rows: [["Statement title"], header, transaction, header, ["2026-08-21 08:20 AM", null, "Cash in via bank", null, null, null, "1234567890123", null, "99.00", null, null, "Existing Customer"]] }]));
    expect(result.rows).toHaveLength(2); expect(result.rows.map((row) => row.rowIndex)).toEqual([3, 5]);
  });

  it("parses headerless concatenated rows and reports unsupported rows", async () => {
    const result = await parseGcashWorkbook(await workbookBuffer([{ name: "Statement", rows: [["2026-04-18 05:41 Transfer from 09331953428 to 09203778483 2039915534650 32000.00 161469.49"], ["not a transaction"]] }]));
    expect(result.rows[0]).toMatchObject({ referenceNumber: "2039915534650", amountCentavos: 3200000 });
    expect(result.warnings[0]?.field).toBe("row");
  });

  it("parses headerless exports with repeated references and duplicated amount columns", async () => {
    const row = (date: string, reference: string, amount: string) => [date, null, "Transfer from 09925251063 to 09190777226", null, null, null, reference, reference, reference, amount, amount];
    const creditOnlyRow = ["2026-09-20 01:21 AM", null, "Transfer from 09565374852 to 09190777226", null, null, null, "4045238564907", "4045238564907", "4045238564907", null, "200.00"];
    const result = await parseGcashWorkbook(await workbookBuffer([{ name: "AllPages", rows: [[], row("2026-09-19 10:01 PM", "9045235006755", "1,000.00"), row("2026-09-20 12:21 AM", "8045238564907", "300.00"), creditOnlyRow] }]));
    expect(result.rows.map((entry) => [entry.rowIndex, entry.referenceNumber, entry.amountCentavos])).toEqual([[2, "9045235006755", 100000], [3, "8045238564907", 30000], [4, "4045238564907", 20000]]);
    expect(result.errors).toHaveLength(0);
  });

  it("rejects aggregate cell budgets before traversing multiple individually valid sheets", async () => {
    const book = new ExcelJS.Workbook();
    for (const name of ["January", "February"]) {
      const sheet = book.addWorksheet(name);
      sheet.getCell(300_001, 1).value = name;
    }
    const bytes = Buffer.from(await book.xlsx.writeBuffer());
    await expect(parseGcashWorkbook(bytes)).rejects.toThrow("Workbook exceeds the 500,000-cell");
    await expect(captureSourceWorkbook(bytes, "gcash")).rejects.toThrow("Workbook exceeds the 500,000-cell");
  });
});
