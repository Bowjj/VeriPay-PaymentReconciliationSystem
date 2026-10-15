"use server";

import { revalidatePath } from "next/cache";
import { getActiveWorkspace } from "@/lib/auth/workspace";
import { preparePaymentImport } from "@/lib/payments/import";
import { prepareGcashImport } from "@/lib/gcash/import";
import { createClient } from "@/lib/supabase/server";
import { paymentImportRows, gcashImportRows } from "@/lib/verification/import-rows";
import { reconcileVerificationRun } from "@/lib/verification/run-reconciliation";
import type { ReconciliationSummary } from "@/lib/verification/reconcile";
import { captureSourceWorkbook } from "@/lib/export/source-workbook";
import { guardSpreadsheetUpload } from "@/lib/excel/upload-guard";

export type VerifyUploadsResult =
  | { success: true; verificationRunId: string; paymentCount: number; gcashCount: number; summary: ReconciliationSummary }
  | { success: false; error: string; verificationRunId?: string };

export async function verifyUploadsAction(formData: FormData): Promise<VerifyUploadsResult> {
  const { workspace } = await getActiveWorkspace();
  if (!workspace || formData.get("workspaceId") !== workspace.id) {
    return { success: false, error: "The active workspace changed. Reload New Verification and select your files again." };
  }
  const payment = formData.get("payment");
  const gcash = formData.getAll("gcash");
  const sessionId = formData.get("sessionId");
  if (!(payment instanceof File) || gcash.length === 0 || !gcash.every((file): file is File => file instanceof File) ||
      typeof sessionId !== "string" || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(sessionId)) {
    return { success: false, error: "A payment file, at least one GCash file, and a verification session are required." };
  }
  const gcashFiles = gcash.filter((file): file is File => file instanceof File);
  const [paymentUpload, ...gcashUploads] = await Promise.all([guardSpreadsheetUpload(payment), ...gcashFiles.map(guardSpreadsheetUpload)]);
  if (!paymentUpload.success) return { success: false, error: paymentUpload.error };
  if (gcashUploads.some((upload) => !upload.success)) return { success: false, error: gcashUploads.find((upload) => !upload.success)?.error ?? "Could not read a GCash workbook." };
  const validGcashUploads = gcashUploads.filter((upload): upload is Extract<typeof upload, { success: true }> => upload.success);
  let payments;
  let transactions;
  try {
    payments = await preparePaymentImport(paymentUpload.bytes);
    transactions = await Promise.all(validGcashUploads.map((upload) => prepareGcashImport(upload.bytes)));
  } catch {
    return { success: false, error: "Could not read one or both Excel workbooks." };
  }
  if (!payments.valid.length || transactions.some((transaction) => !transaction.valid.length)) {
    return { success: false, error: "The payment file and every GCash file must contain valid rows. Review the validation messages." };
  }
  let paymentSource;
  let gcashSources;
  try {
    [paymentSource, ...gcashSources] = await Promise.all([
      captureSourceWorkbook(paymentUpload.bytes, "payments"),
      ...validGcashUploads.map((upload) => captureSourceWorkbook(upload.bytes, "gcash")),
    ]);
  } catch {
    return { success: false, error: "Could not preserve these source workbooks. Each workbook must fit within 500,000 cells." };
  }
  const supabase = await createClient();
  const occurrenceCounts = new Map<string, number>();
  for (const transaction of transactions) for (const row of transaction.valid) occurrenceCounts.set(row.referenceNumber, (occurrenceCounts.get(row.referenceNumber) ?? 0) + 1);
  let sourceOrder = 0;
  const gcashRows = transactions.flatMap((transaction, sourceIndex) => {
    const rows = gcashImportRows(transaction, gcashFiles[sourceIndex]?.name ?? "GCash", sourceIndex, sourceOrder, occurrenceCounts);
    sourceOrder += rows.length;
    return rows;
  });
  const { data, error } = await supabase.rpc("create_multi_gcash_import", {
    p_business_id: workspace.id,
    p_session_id: sessionId,
    p_payments: paymentImportRows(payments, payment.name),
    p_gcash: gcashRows,
    p_payment_source: paymentSource,
    p_gcash_sources: gcashSources,
  });
  if (error || typeof data !== "string") {
    return { success: false, error: "Could not save both datasets. Retry with the same files." };
  }
  try {
    const reconciliation = await reconcileVerificationRun(supabase, workspace.id, data);
    revalidatePath("/history");
    revalidatePath("/dashboard");
    return { success: true, verificationRunId: data, ...reconciliation };
  } catch {
    return { success: false, verificationRunId: data, error: "Both files are saved, but reconciliation has not completed. Retry Verify Payments to resume this same run." };
  }
}
