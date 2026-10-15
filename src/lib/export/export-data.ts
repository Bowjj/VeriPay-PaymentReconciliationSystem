import { z } from "zod";
import { gcashResultSchema, paymentDetailSchema, runSummarySchema } from "@/lib/verification/results-data";
import { sourceWorkbookSchema } from "./source-workbook";

export const exportPaymentSchema = paymentDetailSchema.safeExtend({ business_id: z.uuid(), verification_run_id: z.uuid() });
export const exportGcashSchema = gcashResultSchema.extend({ business_id: z.uuid(), verification_run_id: z.uuid(), matched_payment_id: z.uuid().nullable() });
export const exportDataSchema = z.object({
  workspaceName: z.string(), run: runSummarySchema, payments: z.array(exportPaymentSchema), gcash: z.array(exportGcashSchema),
  sources: z.object({ payments: sourceWorkbookSchema.optional(), gcash: sourceWorkbookSchema.optional(), gcashSources: z.array(z.object({ sourceIndex: z.number().int().nonnegative(), source: sourceWorkbookSchema })).optional() }).optional(),
});
export type ExportData = z.infer<typeof exportDataSchema>;
export type ExportPayment = z.infer<typeof exportPaymentSchema>;
export type ExportGcash = z.infer<typeof exportGcashSchema>;
export class ExportLimitError extends Error {}
export const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
export const MAX_EXPORT_ROWS = 20000;
export const MAX_EXPORT_CELLS = 500000;
