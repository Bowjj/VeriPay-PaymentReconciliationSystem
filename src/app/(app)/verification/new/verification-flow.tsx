"use client";

import { useRef, useState } from "react";
import Link from "next/link";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { UploadDropzone } from "@/components/verification/upload-dropzone";
import { GcashUploadDropzone } from "@/components/verification/gcash-upload-dropzone";
import { PreviewTable } from "@/components/verification/preview-table";
import { GcashPreviewTable } from "@/components/verification/gcash-preview-table";
import { parseUploadAction, type UploadParseResult } from "./actions";
import { parseGcashUploadAction, type GcashUploadParseResult } from "./gcash-actions";
import { verifyUploadsAction } from "./verify-actions";

type ParseResult = UploadParseResult | GcashUploadParseResult;

function useUpload<T extends ParseResult>(parse: (form: FormData) => Promise<T>) {
  const [file, setFile] = useState<File | null>(null);
  const [result, setResult] = useState<T | null>(null);
  const [pending, setPending] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const version = useRef(0);

  function reset() {
    version.current++;
    setFile(null);
    setResult(null);
    setPending(false);
    setError(null);
  }

  async function select(file: File) {
    const current = ++version.current;
    setFile(file);
    setResult(null);
    setError(null);
    setPending(true);
    try {
      if (file.size > 10 * 1024 * 1024) throw new Error("Each file must be 10 MB or smaller.");
      const form = new FormData();
      form.set("file", file);
      const parsed = await parse(form);
      if (version.current === current) setResult(parsed);
    } catch {
      if (version.current === current) setError("Could not read this file. Use an Excel file up to 10 MB and try again.");
    } finally {
      if (version.current === current) setPending(false);
    }
  }

  return { file, result, pending, error, select, reset,
    ready: !!result?.success && (result.validCount ?? 0) > 0 };
}

function useGcashUploads() {
  const [uploads, setUploads] = useState<{ file: File; result: GcashUploadParseResult | null; pending: boolean; error: string | null }[]>([]);
  async function select(files: readonly File[]) {
    const additions = files.map((file) => ({ file, result: null, pending: true, error: null }));
    setUploads((current) => [...current, ...additions]);
    await Promise.all(additions.map(async (entry) => {
      try {
        if (entry.file.size > 10 * 1024 * 1024) throw new Error("Each file must be 10 MB or smaller.");
        const form = new FormData(); form.set("file", entry.file);
        const result = await parseGcashUploadAction(form);
        setUploads((current) => current.map((item) => item.file === entry.file ? { ...item, result, pending: false } : item));
      } catch {
        setUploads((current) => current.map((item) => item.file === entry.file ? { ...item, pending: false, error: "Could not read this file. Use an Excel file up to 10 MB and try again." } : item));
      }
    }));
  }
  return { uploads, select, remove: (file: File) => setUploads((current) => current.filter((item) => item.file !== file)), reset: () => setUploads([]), ready: uploads.length > 0 && uploads.every((item) => !!item.result?.success && (item.result.validCount ?? 0) > 0) };
}

function FileSummary({ result }: { result: ParseResult }) {
  return (
    <div className="space-y-3 text-sm">
      <dl className="grid grid-cols-3 gap-3 border-y py-3 tabular-nums">
        <div><dt className="text-muted-foreground">Found</dt><dd className="text-xl font-semibold">{result.totalRows ?? 0}</dd></div>
        <div><dt className="text-muted-foreground">Valid</dt><dd className="text-xl font-semibold">{result.validCount ?? 0}</dd></div>
        <div><dt className="text-muted-foreground">Attention</dt><dd className="text-xl font-semibold">{(result.errors?.length ?? 0) + (result.warnings?.length ?? 0)}</dd></div>
      </dl>
      {(!result.success || !result.validCount) && <p role="alert" className="text-destructive">No valid rows. Replace this file to continue.</p>}
      {!!((result.errors?.length ?? 0) + (result.warnings?.length ?? 0)) && (
        <details className="rounded-md border border-amber-300 p-3 dark:border-amber-700">
          <summary className="cursor-pointer font-medium">Validation messages</summary>
          <p className="my-2 text-muted-foreground">Invalid rows are excluded. Warnings are retained for review; only valid rows will be saved.</p>
          <ul className="max-h-48 space-y-1 overflow-auto text-xs">
            {[...(result.errors ?? []), ...(result.warnings ?? [])].map((issue, index) => (
              <li key={index}>Row {issue.rowIndex}: {issue.field} — {issue.message}</li>
            ))}
          </ul>
        </details>
      )}
    </div>
  );
}

export function VerificationFlow({ workspaceId }: { workspaceId: string }) {
  const payment = useUpload(parseUploadAction);
  const gcash = useGcashUploads();
  const [saving, setSaving] = useState(false);
  const [submitted, setSubmitted] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [saved, setSaved] = useState<Awaited<ReturnType<typeof verifyUploadsAction>> | null>(null);
  const sessionId = useRef<string | null>(null);
  const savingRef = useRef(false);

  async function verify() {
    if (!payment.ready || !gcash.ready || !payment.file || savingRef.current) return;
    savingRef.current = true;
    setSaving(true);
    setSubmitted(true);
    setError(null);
    sessionId.current ??= crypto.randomUUID();
    const form = new FormData();
    form.set("payment", payment.file);
    gcash.uploads.forEach(({ file }) => form.append("gcash", file));
    form.set("workspaceId", workspaceId);
    form.set("sessionId", sessionId.current);
    try {
      const result = await verifyUploadsAction(form);
      if (result.success) setSaved(result);
      else setError(result.error ?? "Could not save both files. Retry to continue.");
    } catch {
      setError("Connection interrupted. Retry to safely resume the same verification session.");
    } finally {
      savingRef.current = false;
      setSaving(false);
    }
  }

  function reset() {
    payment.reset();
    gcash.reset();
    sessionId.current = null;
    setSubmitted(false);
    setSaved(null);
    setError(null);
  }

  if (saved?.success) return (
    <Card><CardHeader><CardTitle>Verification Complete</CardTitle></CardHeader>
      <CardContent className="space-y-4">
        <p>{payment.file?.name} + {gcash.uploads.map(({ file }) => file.name).join(", ")}</p>
        <p>{saved.paymentCount} payment records and {saved.gcashCount} GCash transactions saved to one verification run.</p>
        {saved.summary && <dl aria-label="Reconciliation summary" className="grid grid-cols-2 gap-4 border-y py-4 sm:grid-cols-5">
          <div><dt className="text-sm text-muted-foreground">Total Payments</dt><dd className="text-2xl font-semibold tabular-nums">{saved.summary.total}</dd></div>
          <div><dt className="text-sm text-muted-foreground">Verified</dt><dd className="text-2xl font-semibold tabular-nums">{saved.summary.verified}</dd></div>
          <div><dt className="text-sm text-muted-foreground">Needs Review</dt><dd className="text-2xl font-semibold tabular-nums">{saved.summary.needsReview}</dd></div>
          <div><dt className="text-sm text-muted-foreground">Cash</dt><dd className="text-2xl font-semibold tabular-nums">{saved.summary.cash}</dd></div>
          <div><dt className="text-sm text-muted-foreground">Bank</dt><dd className="text-2xl font-semibold tabular-nums">{saved.summary.bank}</dd></div>
        </dl>}
        <p className="text-sm text-muted-foreground">Verified means one unique exact reference match. Amount differences do not affect verification. Payments needing review have no customer attached to a GCash transaction.</p>
        <p className="break-all text-xs text-muted-foreground">Run: {saved.verificationRunId}</p>
        <Link className="inline-flex rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground" href={`/verification/${saved.verificationRunId}`}>View Results</Link>
        <Button variant="outline" onClick={reset}>Start new verification</Button>
      </CardContent>
    </Card>
  );

  return (
    <div className="space-y-6">
      {saving && <div role="alertdialog" aria-modal="true" aria-labelledby="verification-loading-title" aria-describedby="verification-loading-description" className="fixed inset-0 z-[60] grid place-items-center bg-background/80 p-6 backdrop-blur-sm">
        <div className="w-full max-w-sm space-y-4 rounded-xl border bg-card p-6 text-center shadow-lg">
          <span aria-hidden="true" className="mx-auto block size-10 animate-spin rounded-full border-4 border-primary border-t-transparent" />
          <div className="space-y-2"><h2 id="verification-loading-title" className="text-lg font-semibold">Verifying payments</h2><p id="verification-loading-description" className="text-sm text-muted-foreground">Please wait while VeriPay saves your files and checks exact payment references. Do not close or navigate away from this page.</p></div>
        </div>
      </div>}
      <p className="text-sm text-muted-foreground">Select both files in any order. Preview and replace either file before you continue.</p>
      <div className="grid items-start gap-6 xl:grid-cols-2">
        <section aria-label="Payment Records" className="min-w-0">
          <Card><CardHeader><CardTitle>Payment Records</CardTitle></CardHeader>
            <CardContent className="space-y-4">
              {!payment.file ? <UploadDropzone onFileSelected={payment.select} isProcessing={submitted} /> : <>
                <p className="break-all font-medium">{payment.file.name}</p>
                {payment.pending && <p role="status">Reading payment records…</p>}
                {payment.error && <p role="alert">{payment.error}</p>}
                {payment.result && <>
                  <FileSummary result={payment.result} />
                  <p className="text-sm">GCash: {payment.result.summary?.gcashCount ?? 0} / Cash: {payment.result.summary?.cashCount ?? 0} / Bank: {payment.result.summary?.bankCount ?? 0} / Unknown: {payment.result.summary?.unknownMethodCount ?? 0}</p>
                  <details><summary className="cursor-pointer text-sm font-medium">Preview</summary>
                    <div className="mt-3"><PreviewTable rows={payment.result.preview ?? []} totalCount={payment.result.validCount ?? 0} /></div>
                  </details>
                </>}
                <Button variant="outline" disabled={submitted} onClick={payment.reset}>Replace Payment File</Button>
              </>}
            </CardContent>
          </Card>
        </section>
        <section aria-label="GCash Statement" className="min-w-0">
          <Card><CardHeader><CardTitle>GCash Statements</CardTitle></CardHeader>
            <CardContent className="space-y-4">
               <GcashUploadDropzone onFilesSelected={gcash.select} isProcessing={submitted} />
               {gcash.uploads.length === 1 && <Button variant="outline" disabled={submitted} onClick={gcash.reset}>Replace GCash File</Button>}
              {gcash.uploads.map(({ file, result, pending, error }) => <div key={`${file.name}-${file.lastModified}`} className="space-y-3 rounded-lg border p-3">
                <div className="flex items-start justify-between gap-3"><p className="break-all font-medium">{file.name}</p><Button variant="outline" size="sm" disabled={submitted} onClick={() => gcash.remove(file)}>Remove</Button></div>
                {pending && <p role="status">Reading GCash transactions…</p>}{error && <p role="alert">{error}</p>}
                {result && <><FileSummary result={result} /><p className="text-sm">Duplicate References: {result.summary?.duplicateReferenceCount ?? 0}</p>
                  <details><summary className="cursor-pointer text-sm font-medium">Preview</summary><div className="mt-3"><GcashPreviewTable rows={result.preview ?? []} totalCount={result.validCount ?? 0} /></div></details></>}
              </div>)}
            </CardContent>
          </Card>
        </section>
      </div>
       <div className="space-y-3 border-t pt-6">
         <Button onClick={verify} disabled={!payment.ready || !gcash.ready || saving} aria-describedby="verify-help">
           {saving ? "Verifying payments…" : "Verify Payments"}
         </Button>
         <p id="verify-help" className="max-w-2xl text-sm text-muted-foreground">
            {payment.ready && gcash.ready ? "All files are ready. Save and verify by exact reference. Amounts do not determine verification." : "Upload Payment Records and at least one GCash Statement to continue."}
         </p>
        {error && <p role="alert" className="text-sm text-destructive">{error}</p>}
      </div>
    </div>
  );
}
