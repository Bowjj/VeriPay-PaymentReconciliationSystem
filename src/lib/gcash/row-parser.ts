import { parseExcelIdentifier } from "@/lib/excel/identifier";
import { parseCentavos } from "@/lib/money";

import type { GcashDirection } from "./types";

export type ParsedGcashRowData = {
  readonly transactionDate: string;
  readonly description: string;
  readonly referenceNumber: string;
  readonly amountCentavos: number;
  readonly direction: GcashDirection;
  readonly raw: Record<string, unknown>;
  readonly warning?: string;
};

export type GcashRowParseResult =
  | { readonly kind: "transaction"; readonly data: ParsedGcashRowData }
  | { readonly kind: "skip" }
  | { readonly kind: "error"; readonly field: string; readonly message: string }
  | { readonly kind: "unsupported"; readonly message: string };

const CONCATENATED_ROW =
  /^(\d{4}-\d{2}-\d{2}\s+\d{1,2}:\d{2}(?:\s*(?:AM|PM))?)\s+(.+?)\s+(\d{8,15})\s+([\d,.]+)\s+([\d,.]+)$/;

function cellText(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function isMetadata(value: string): boolean {
  const upper = value.toUpperCase();
  return (
    !upper ||
    upper.includes("BALANCE") ||
    upper.includes("TOTAL DEBIT") ||
    upper.includes("TOTAL CREDIT") ||
    upper.startsWith("DATE AND TIME") ||
    upper.startsWith("ACCOUNT") ||
    upper.startsWith("REF NO")
  );
}

function detectDirection(description: string): GcashDirection {
  const lower = description.toLowerCase();
  if (
    lower.includes("buy load") ||
    lower.includes("sent gcash") ||
    lower.includes("payment to") ||
    lower.includes("pay to") ||
    lower.includes("cash out")
  ) {
    return "outgoing";
  }
  if (
    lower.includes("cash in") ||
    lower.includes("received gcash") ||
    lower.includes("payment received") ||
    lower.includes("deposit")
  ) {
    return "incoming";
  }
  return "unknown";
}

function parseConcatenated(value: string): GcashRowParseResult | null {
  const match = CONCATENATED_ROW.exec(value);
  if (!match) return null;
  const [, transactionDate, description, referenceNumber, amount] = match;
  if (!transactionDate || !description || !referenceNumber || !amount) return null;
  const parsedAmount = parseCentavos(amount);
  if (!parsedAmount.success) {
    return { kind: "error", field: "amount", message: parsedAmount.error };
  }
  return {
    kind: "transaction",
    data: {
      transactionDate,
      description,
      referenceNumber,
      amountCentavos: parsedAmount.centavos,
      direction: detectDirection(description),
      raw: { col0: value },
    },
  };
}

export function parseGcashRow(
  rawRow: readonly unknown[],
  formattedRow: readonly unknown[],
  amountColumns: readonly number[] = [8, 10],
  deduplicateMatchingAmounts = false,
): GcashRowParseResult {
  const firstCell = cellText(formattedRow[0]);
  if (isMetadata(firstCell)) return { kind: "skip" };

  const concatenated = parseConcatenated(firstCell);
  if (concatenated) return concatenated;

  const transactionDate = firstCell;
  const description = cellText(formattedRow[2]);
  if (!transactionDate && !description) return { kind: "skip" };
  if (!description || !/^\d{4}-\d{2}-\d{2}/.test(transactionDate)) {
    return { kind: "unsupported", message: "Unrecognized non-transaction row" };
  }

  const reference = parseExcelIdentifier(rawRow[6], cellText(formattedRow[6]));
  if (!reference.success) {
    return { kind: "error", field: "referenceNumber", message: reference.error };
  }
  if (!reference.value) {
    return { kind: "error", field: "referenceNumber", message: "Missing reference number" };
  }

  const amountValues = amountColumns.map((column) => cellText(formattedRow[column])).filter(Boolean);
  const distinctAmountValues = deduplicateMatchingAmounts ? [...new Set(amountValues)] : amountValues;
  if (distinctAmountValues.length !== 1) {
    return {
      kind: "error",
      field: "amount",
      message: distinctAmountValues.length === 0 ? "Missing transaction amount" : "Ambiguous debit and credit amounts",
    };
  }
  const parsedAmount = parseCentavos(distinctAmountValues[0] ?? "");
  if (!parsedAmount.success) {
    return { kind: "error", field: "amount", message: parsedAmount.error };
  }

  return {
    kind: "transaction",
    data: {
      transactionDate,
      description,
      referenceNumber: reference.value,
      amountCentavos: parsedAmount.centavos,
      direction: detectDirection(description),
      raw: {
        date: rawRow[0],
        description: rawRow[2],
        ref: rawRow[6],
        debit: rawRow[8],
        credit: rawRow[10],
      },
      ...(reference.warning ? { warning: reference.warning } : {}),
    },
  };
}
