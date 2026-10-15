"use client";

import { useState } from "react";
import { Button } from "@/components/ui/button";

export function DownloadExcelButton({ runId, sourcePreserved, gcashSourceCount = 1 }: { runId: string; sourcePreserved: boolean; gcashSourceCount?: number }) {
  const [pending, setPending] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [warning, setWarning] = useState<string | null>(null);
  async function download(kind: "payments" | "gcash", sourceIndex = 0) {
    const key = `${kind}-${sourceIndex}`;
    setPending(key);
    setError(null);
    setWarning(null);
    try {
      const response = await fetch(`/verification/${encodeURIComponent(runId)}/export?file=${kind}&source=${sourceIndex}`, { cache: "no-store" });
      if (!response.ok) {
        const body: unknown = await response.json().catch(() => null);
        throw new Error(body && typeof body === "object" && "error" in body && typeof body.error === "string" ? body.error : "Export failed.");
      }
      if (!response.headers.get("content-type")?.includes("spreadsheetml.sheet")) throw new Error("Please sign in and reopen this verification before exporting.");
      const filename = /filename="([A-Za-z0-9 _().-]+\.xlsx)"/.exec(response.headers.get("content-disposition") ?? "")?.[1] ?? `${kind === "payments" ? "Payment Records - Verified" : "GCash - Matched"}.xlsx`;
      const url = URL.createObjectURL(await response.blob());
      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      document.body.append(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      const conflicts = Number(response.headers.get("X-GCash-Annotation-Conflicts") ?? "0");
      if (kind === "gcash" && Number.isSafeInteger(conflicts) && conflicts > 0) setWarning(`GCash downloaded with ${conflicts} annotation conflict${conflicts === 1 ? "" : "s"}. Review the Annotation Conflict column in the workbook. Historical names were preserved; verification results are unchanged.`);
    } catch (error) {
      setError(error instanceof Error ? error.message : "Could not download the report. Try again.");
    } finally { setPending(null); }
  }
  return <div className="space-y-2"><div className="flex flex-wrap gap-2">
    <Button variant="outline" className="min-h-11 px-4" onClick={() => download("payments")} disabled={!!pending}>{pending === "payments-0" ? "Preparing Payment Records…" : "Download Payment Records"}</Button>
    {Array.from({ length: gcashSourceCount }, (_, sourceIndex) => <Button key={sourceIndex} variant="outline" className="min-h-11 px-4" onClick={() => download("gcash", sourceIndex)} disabled={!!pending}>{pending === `gcash-${sourceIndex}` ? "Preparing GCash…" : gcashSourceCount === 1 ? "Download GCash" : `Download GCash ${sourceIndex + 1}`}</Button>)}
    </div>
    {!sourcePreserved && <p className="max-w-2xl text-xs text-muted-foreground">This older run has no saved original workbooks. Downloads contain retained imported rows only; original formatting and omitted content cannot be recovered. Upload the original files in a new verification for source-preserving downloads.</p>}
    {error && <p role="alert" className="text-sm text-destructive">{error}</p>}
    {warning && <p role="status" className="max-w-2xl rounded-lg border border-amber-500/40 bg-amber-50 p-3 text-sm text-amber-900 dark:bg-amber-950/40 dark:text-amber-200">{warning}</p>}
  </div>;
}
