import { beforeEach, expect, test, vi } from "vitest";
import { workbookBuffer } from "../fixtures/workbook";
import { verifyUploadsAction } from "@/app/(app)/verification/new/verify-actions";

const { rpc, workspace } = vi.hoisted(() => ({
  rpc: vi.fn(),
  workspace: vi.fn(),
}));
vi.mock("@/lib/auth/workspace", () => ({ getActiveWorkspace: workspace }));
vi.mock("@/lib/supabase/server", () => ({ createClient: async () => ({ rpc }) }));
vi.mock("next/cache", () => ({ revalidatePath: vi.fn() }));

async function file(rows: unknown[][], name: string) {
  return new File([new Uint8Array(await workbookBuffer([{ name: "Sheet1", rows }]))], name);
}

async function inputs() {
  const form = new FormData();
  form.set("workspaceId", "active-workspace");
  form.set("sessionId", "10000000-0000-0000-0000-000000000001");
  form.set("payment", await file([
    ["Customer", "Amount", "Method", "Reference #"],
    ...Array.from({ length: 74 }, (_, i) => [`Customer ${i}`, "1,000.25", "gcash", `000${i}`]),
  ], "payments.xlsx"));
  form.set("gcash", await file([
    ["Date and Time", null, "ACCOUNTS TO", null, null, null, "REF NO", null, "AMOUNT"],
    ...Array.from({ length: 82 }, (_, i) => ["2026-08-20 08:00 AM", null, "Unspecified transaction", null, null, null, `000${i}`, null, "1,000.25"]),
  ], "gcash.xlsx"));
  return form;
}

beforeEach(() => {
  vi.clearAllMocks();
  workspace.mockResolvedValue({ workspace: { id: "active-workspace" } });
  rpc.mockReset();
  rpc.mockImplementation(async (name: string) => {
    if (name === "create_multi_gcash_import") return { data: "run-id", error: null };
    if (name === "get_reconciliation_inputs") return { data: {
      payments: Array.from({ length: 74 }, (_, i) => ({ id: `p${i}`, business_id: "active-workspace", verification_run_id: "run-id", method: "GCASH", reference_number: `000${i}` })),
      gcash: Array.from({ length: 82 }, (_, i) => ({ id: `g${i}`, business_id: "active-workspace", verification_run_id: "run-id", reference_number: `000${i}`, reference_occurrence_count: 1 })),
    }, error: null };
    return { data: { total: 74, verified: 74, needsReview: 0, cash: 0, bank: 0 }, error: null };
  });
});

test("reparses original files, atomically imports all rows, then reconciles persisted inputs", async () => {
  const form = await inputs();
  form.set("preview", JSON.stringify([{ amount: "999999", business_id: "forged" }]));
  const result = await verifyUploadsAction(form);
  expect(result).toMatchObject({ success: true, paymentCount: 74, gcashCount: 82 });
  expect(rpc).toHaveBeenCalledTimes(3);
  expect(rpc).toHaveBeenCalledWith("create_multi_gcash_import", expect.objectContaining({
    p_business_id: "active-workspace",
    p_payments: expect.arrayContaining([expect.objectContaining({ amount: "1000.25", reference_number: "0000" })]),
    p_gcash: expect.arrayContaining([expect.objectContaining({ amount: "1000.25", reference_number: "0000", direction: "unknown" })]),
  }));
  expect(rpc.mock.calls[0]?.[1].p_payments).toHaveLength(74);
  expect(rpc.mock.calls[0]?.[1].p_gcash).toHaveLength(82);
  expect(rpc.mock.calls[0]?.[1].p_payment_source).toMatchObject({ version: 1, sheets: [expect.objectContaining({ name: "Sheet1", rowOffset: 0, headerRow: 1 })] });
  const originalFile = form.get("payment");
  if (!(originalFile instanceof File)) throw new Error("Missing fixture file");
  expect(Buffer.from(rpc.mock.calls[0]?.[1].p_payment_source.base64, "base64")).toEqual(Buffer.from(await originalFile.arrayBuffer()));
  expect(rpc.mock.calls[2]?.[1].p_results).toHaveLength(74);
  expect(result).toMatchObject({ summary: { total: 74, verified: 74, needsReview: 0, cash: 0, bank: 0 } });
});

test("imports every GCash workbook with source provenance and cross-file duplicate counts", async () => {
  const form = await inputs();
  form.append("gcash", await file([
    ["Date and Time", null, "ACCOUNTS TO", null, null, null, "REF NO", null, "AMOUNT"],
    ["2026-08-21 08:00 AM", null, "Unspecified transaction", null, null, null, "0000", null, "100.00"],
  ], "gcash-second.xlsx"));
  const result = await verifyUploadsAction(form);
  expect(result).toMatchObject({ success: true });
  const args = rpc.mock.calls[0]?.[1];
  expect(args.p_gcash_sources).toHaveLength(2);
  expect(args.p_gcash).toEqual(expect.arrayContaining([
    expect.objectContaining({ reference_number: "0000", reference_occurrence_count: 2, raw_data: expect.objectContaining({ source_index: 0 }) }),
    expect.objectContaining({ reference_number: "0000", reference_occurrence_count: 2, raw_data: expect.objectContaining({ source_index: 1, source_order: 82 }) }),
  ]));
});

test("rejects a workspace changed in another tab or forged by the client", async () => {
  const form = await inputs();
  form.set("workspaceId", "other-business");
  expect(await verifyUploadsAction(form)).toMatchObject({ success: false });
  expect(rpc).not.toHaveBeenCalled();
});

test.each(["payment", "gcash"])("refuses persistence when %s file is missing", async (key) => {
  const form = await inputs();
  form.delete(key);
  expect(await verifyUploadsAction(form)).toMatchObject({ success: false });
  expect(rpc).not.toHaveBeenCalled();
});

test("rejects corrupted authoritative input despite a claimed valid preview", async () => {
  const form = await inputs();
  form.set("gcash", await file([["Unsupported"], ["Meaningful bad row"]], "bad.xlsx"));
  form.set("validCount", "82");
  expect(await verifyUploadsAction(form)).toMatchObject({ success: false });
  expect(rpc).not.toHaveBeenCalled();
});

test("failed atomic save is reported without claiming import success", async () => {
  rpc.mockResolvedValue({ data: null, error: { message: "database failure" } });
  expect(await verifyUploadsAction(await inputs())).toMatchObject({ success: false });
});

test("failed reconciliation leaves a resumable saved run and does not report completion", async () => {
  rpc.mockResolvedValueOnce({ data: "run-id", error: null });
  rpc.mockResolvedValueOnce({ data: null, error: { message: "read failed" } });
  expect(await verifyUploadsAction(await inputs())).toMatchObject({ success: false, verificationRunId: "run-id", error: expect.stringContaining("Retry Verify Payments") });
  expect(rpc).toHaveBeenCalledTimes(2);
});
