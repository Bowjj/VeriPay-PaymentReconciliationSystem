import type { preparePaymentImport } from "@/lib/payments/import";
import type { prepareGcashImport } from "@/lib/gcash/import";
import { centavosToDecimal } from "@/lib/money";

export function paymentImportRows(prepared: Awaited<ReturnType<typeof preparePaymentImport>>, filename: string) {
  return prepared.valid.map((r) => ({
    customer: r.customer,
    account: r.account || null,
    billing_period: r.billingPeriod || null,
    amount: centavosToDecimal(r.amountCentavos),
    method: r.method,
    reference_number: r.referenceNumber || null,
    payment_date: r.paymentDate || null,
    notes: r.notes || null,
    paid_by: r.paidBy || null,
    received_by: r.receivedBy || null,
    photo_url: r.photoUrl || null,
    created_at_source: r.createdAt || null,
    row_index: r.rowIndex,
    raw_data: { filename, source: r.rawData },
  }));
}

export function gcashImportRows(prepared: Awaited<ReturnType<typeof prepareGcashImport>>, filename: string, sourceIndex = 0, sourceOrderOffset = 0, occurrenceCounts?: ReadonlyMap<string, number>) {
  return prepared.valid.map((r, sourceOrder) => ({
    transaction_date: r.transactionDate || null,
    description: r.description || null,
    reference_number: r.referenceNumber || null,
    amount: centavosToDecimal(r.amountCentavos),
    direction: r.direction,
    reference_occurrence_count: occurrenceCounts?.get(r.referenceNumber) ?? r.referenceOccurrenceCount,
    row_index: r.rowIndex,
    raw_data: { filename, sheet_name: r.sheetName, source_index: sourceIndex, source_order: sourceOrder + sourceOrderOffset, source: r.rawData },
  }));
}
