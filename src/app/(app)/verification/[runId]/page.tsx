import Link from "next/link";
import { z } from "zod";
import { authorizedRun } from "@/lib/verification/results-server";
import { gcashResultSchema, paymentPageSchema } from "@/lib/verification/results-data";
import { displayDate, displayMoney, matchesResult, reasonMessages, resultsQuerySchema } from "@/lib/verification/review";
import { Pagination, ProofLink, ResultStatus, ResultSummary } from "@/components/verification/result-shared";
import { DownloadExcelButton } from "@/components/verification/download-excel-button";

export default async function VerificationResultsPage({ params, searchParams }: {
  params: Promise<{ runId: string }>;
  searchParams: Promise<Record<string, string | string[] | undefined>>;
}) {
  const { runId } = await params;
  const { workspace, supabase, run } = await authorizedRun(runId);
  const sources = await supabase.from("verification_source_workbooks").select("kind,source_index").eq("verification_run_id", runId).eq("business_id", workspace.id);
  if (sources.error) throw new Error("Could not load source availability.");
  const query = resultsQuerySchema.parse(await searchParams);
  const base = `/verification/${runId}`;
  function href(changes: Record<string, string | number>) {
    return `${base}?${new URLSearchParams({ status: query.status, q: query.q, sort: query.sort, view: query.view, page: String(query.page), ...Object.fromEntries(Object.entries(changes).map(([key, value]) => [key, String(value)])) })}`;
  }
  let payments: z.infer<typeof paymentPageSchema> = { rows: [], total: 0 };
  let transactions: z.infer<typeof gcashResultSchema>[] = [];
  let gcashTotal = 0;
  if (query.view === "payments") {
    const { data, error } = await supabase.rpc("list_payment_results", {
      p_business_id: workspace.id, p_run_id: runId, p_status: query.status, p_search: query.q, p_page: query.page, p_sort: query.sort,
    });
    if (error) throw new Error("Could not load payment results.");
    payments = paymentPageSchema.parse(data);
    if (!payments.rows.every((row) => matchesResult(row, query))) throw new Error("Unexpected result query response.");
  } else {
    const { data, count, error } = await supabase.from("gcash_results").select("*", { count: "exact" })
      .eq("business_id", workspace.id).eq("verification_run_id", runId)
      .order("source_order").order("row_index").order("id").range((query.page - 1) * 25, query.page * 25 - 1);
    if (error) throw new Error("Could not load GCash transactions.");
    transactions = z.array(gcashResultSchema).parse(data ?? []);
    gcashTotal = count ?? 0;
  }
  const filters = [
    { status: "ALL", label: "All", count: run.total }, { status: "VERIFIED", label: "Verified", count: run.verified },
    { status: "NEEDS_REVIEW", label: "Needs Review", count: run.needs_review }, { status: "CASH", label: "Cash", count: run.cash }, { status: "BANK", label: "Bank", count: run.bank },
  ];
  return <div className="min-w-0 space-y-6">
    <header className="space-y-2">
      <Link href="/history" className="text-sm underline">Back to History</Link>
      <h1 className="text-2xl font-semibold tracking-tight">Verification Results</h1>
      <p>Workspace: <strong>{workspace.name}</strong></p>
       <DownloadExcelButton runId={runId} sourcePreserved={!!sources.data?.some((source) => source.kind === "payments") && !!sources.data?.some((source) => source.kind === "gcash")} gcashSourceCount={Math.max(1, sources.data?.filter((source) => source.kind === "gcash").length ?? 0)} />
      <dl className="grid gap-2 text-sm sm:grid-cols-3">
        <div className="min-w-0"><dt className="text-muted-foreground">Payment File</dt><dd className="break-all">{run.payment_filename ?? "Unavailable"}</dd></div>
        <div className="min-w-0"><dt className="text-muted-foreground">GCash File</dt><dd className="break-all">{run.gcash_filename ?? "Unavailable"}</dd></div>
        <div><dt className="text-muted-foreground">Verification Date</dt><dd>{displayDate(run.completed_at)}</dd></div>
      </dl>
    </header>
    <ResultSummary run={run} />
    <nav aria-label="Result views" className="flex gap-4 border-b pb-3">
      <Link className={query.view === "payments" ? "font-semibold underline" : "text-muted-foreground"} href={href({ view: "payments", page: 1 })}>Payment Results</Link>
      <Link className={query.view === "gcash" ? "font-semibold underline" : "text-muted-foreground"} href={href({ view: "gcash", page: 1 })}>GCash Transactions</Link>
    </nav>
    {query.view === "payments" ? <>
      <nav aria-label="Status filters" className="flex flex-wrap gap-2">
        {filters.map((filter) => <Link key={filter.status} aria-current={query.status === filter.status ? "page" : undefined}
          className={`rounded-lg border px-3 py-2 text-sm ${query.status === filter.status ? "bg-primary text-primary-foreground" : "hover:bg-muted"}`}
          href={href({ status: filter.status, page: 1 })}>{filter.label} ({filter.count})</Link>)}
      </nav>
      <form action={base} className="flex flex-wrap items-end gap-3">
        <input type="hidden" name="status" value={query.status} />
        <label className="w-full min-w-0 text-sm sm:w-auto sm:flex-1">Search customer or reference
          <input className="mt-1 block w-full rounded-md border p-2" name="q" defaultValue={query.q} maxLength={100} placeholder="Customer or exact reference text" />
        </label>
        <label className="text-sm">Sort by
          <select name="sort" defaultValue={query.sort} className="mt-1 block rounded-md border p-2">
            <option value="source">Source order</option><option value="customer">Customer</option><option value="payment_date">Payment date</option><option value="status">Final status</option>
          </select>
        </label>
        <button className="rounded-md bg-primary px-4 py-2 text-sm text-primary-foreground">Search</button>
      </form>
      {payments.rows.length === 0 ? <p className="rounded-xl border p-6 text-muted-foreground">No payments match these filters.</p> : <>
        <div className="hidden overflow-x-auto rounded-xl border lg:block">
          <table className="w-full text-left text-sm"><thead className="bg-muted"><tr>
            {["Customer", "Method", "Payment Amount", "Reference", "Payment Date", "Final Status", "Automated Reason", "Proof", "Action"].map((title) => <th key={title} className="p-3 font-medium">{title}</th>)}
          </tr></thead><tbody>
            {payments.rows.map((row) => <tr key={row.id} className="border-t align-top">
              <td className="p-3 font-medium">{row.customer}</td><td className="p-3">{row.method}</td><td className="whitespace-nowrap p-3 tabular-nums">{displayMoney(row.amount_decimal)}</td>
              <td className="whitespace-nowrap p-3">{row.reference_number || "—"}</td><td className="p-3">{displayDate(row.payment_date)}</td>
              <td className="p-3"><ResultStatus status={row.effective_status} manual={!!row.manual_status} /></td>
              <td className="min-w-48 p-3">{reasonMessages[row.automated_reason] ?? row.automated_reason}</td>
              <td className="p-3"><ProofLink url={row.photo_url} /></td><td className="p-3"><Link className="underline" href={`${base}/payments/${row.id}`}>Details</Link></td>
            </tr>)}
          </tbody></table>
        </div>
        <div className="space-y-3 lg:hidden">{payments.rows.map((row) => <article key={row.id} className="space-y-3 rounded-xl border p-4">
          <div className="flex flex-wrap justify-between gap-2"><h2 className="font-semibold">{row.customer}</h2><ResultStatus status={row.effective_status} manual={!!row.manual_status} /></div>
          <p className="text-sm">{row.method} · {displayMoney(row.amount_decimal)}</p><p className="break-all text-sm">Reference: {row.reference_number || "—"}</p>
          <p className="text-sm text-muted-foreground">{reasonMessages[row.automated_reason] ?? row.automated_reason}</p>
          <div className="flex flex-wrap gap-4"><ProofLink url={row.photo_url} /><Link className="text-sm underline" href={`${base}/payments/${row.id}`}>Details</Link></div>
        </article>)}</div>
      </>}
      <Pagination page={query.page} total={payments.total} href={(page) => href({ page })} />
    </> : <>
      <p className="text-sm text-muted-foreground">Original GCash source order. Customer names come from stored transaction links; manual status-only decisions do not attach a customer.</p>
      <div className="hidden overflow-x-auto rounded-xl border lg:block"><table className="w-full text-left text-sm">
        <thead className="bg-muted"><tr>{["Date", "Description", "Reference", "Amount", "Match Status", "Matched Customer"].map((label) => <th key={label} className="p-3">{label}</th>)}</tr></thead>
        <tbody>{transactions.map((row) => <tr key={row.id} className="border-t align-top"><td className="p-3">{displayDate(row.transaction_date)}</td><td className="p-3">{row.description}</td><td className="break-all p-3">{row.reference_number}</td><td className="whitespace-nowrap p-3">{displayMoney(row.amount_decimal)}</td><td className="p-3">{row.matched_customer ? "Automated reference match" : "No linked payment"}</td><td className="p-3 font-medium">{row.matched_customer ?? "—"}</td></tr>)}</tbody>
      </table></div>
      <div className="space-y-3 lg:hidden">{transactions.map((row) => <article key={row.id} className="space-y-2 rounded-xl border p-4 text-sm">
        <p>{displayDate(row.transaction_date)} · {displayMoney(row.amount_decimal)}</p><p>{row.description}</p><p className="break-all">Reference: {row.reference_number}</p><p>Matched Customer: <strong>{row.matched_customer ?? "—"}</strong></p><p className="text-muted-foreground">{row.matched_customer ? "Automated reference match" : "No linked payment"}</p>
      </article>)}</div>
      <Pagination page={query.page} total={gcashTotal} href={(page) => href({ page })} />
    </>}
  </div>;
}
