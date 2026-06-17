"use client";

// Results panel — a React port of index.html's showResults(): KPI cards
// (GGC NOI, cap rate, $/unit, occupancy, optional API cost), diligence flags,
// comps summary, and the authenticated .xlsx download.

import { useState } from "react";
import { downloadExcel, type JobResult } from "@/lib/engine";

function formatMoney(n: number): string {
  return "$" + Math.round(n).toLocaleString("en-US");
}

function KpiCard({
  label,
  value,
  subtitle,
}: {
  label: string;
  value: string;
  subtitle?: string;
}) {
  return (
    <div className="bg-slate-900/50 rounded-lg p-3 border border-slate-700/50">
      <div className="text-xs text-slate-500 uppercase mb-1">{label}</div>
      <div className="text-xl font-bold text-slate-100">{value}</div>
      {subtitle ? <div className="text-[10px] text-slate-500 mt-1">{subtitle}</div> : null}
    </div>
  );
}

// Literal class names (not template-built) so Tailwind generates them.
function severityClass(severity?: string): string {
  const sev = (severity || "").toLowerCase();
  if (sev === "high") return "text-red-400";
  if (sev === "medium") return "text-amber-400";
  return "text-green-400";
}

export default function ResultsPanel({ result }: { result: JobResult }) {
  const [downloading, setDownloading] = useState(false);
  const [downloadError, setDownloadError] = useState<string | null>(null);

  const f = result.financials ?? {};
  const m = result.market ?? {};
  const pi = f.propertyInfo ?? {};
  const rr = f.rentRoll ?? {};
  const rentComps = m.rentComps ?? [];
  const saleComps = m.saleComps ?? [];

  // KPI math. Exclude the Omitt buckets from EGI / OpEx so the displayed
  // NOI matches what the workbook's I47 SUMIFS will compute — Omitt is the
  // explicit non-operating exclusion. The Omitt totals are surfaced as
  // their own tile so the reviewer can sanity-check what was excluded.
  const isOmittIncome = (line: { ggcCategory?: string }) =>
    (line.ggcCategory ?? "").trim() === "Omitt Income";
  const isOmittExpense = (line: { ggcCategory?: string }) =>
    (line.ggcCategory ?? "").trim() === "Omitt Expense";
  const totalInc = (f.income ?? [])
    .filter((l) => !isOmittIncome(l))
    .reduce((sum, line) => sum + (line.ggcUnderwritten ?? 0), 0);
  const totalExp = (f.expenses ?? [])
    .filter((l) => !isOmittExpense(l))
    .reduce((sum, line) => sum + (line.ggcUnderwritten ?? 0), 0);
  const omittIncome = (f.income ?? [])
    .filter(isOmittIncome)
    .reduce((sum, line) => sum + (line.ggcUnderwritten ?? 0), 0);
  const omittExpense = (f.expenses ?? [])
    .filter(isOmittExpense)
    .reduce((sum, line) => sum + (line.ggcUnderwritten ?? 0), 0);
  const omittTotal = omittIncome + omittExpense;
  const ggcNoi = totalInc - totalExp;
  const ask = pi.askingPrice ?? 0;
  const capRate = ask ? ggcNoi / ask : 0;
  const pricePerUnit = pi.totalUnits ? ask / pi.totalUnits : 0;
  const occupancy = ((rr.occupancyRate ?? 0) * 100).toFixed(1) + "%";

  // Per-deal API spend (absent on results that predate usage tracking).
  const cost =
    typeof result.usage?.totals?.cost_usd === "number" ? result.usage.totals.cost_usd : null;

  const flags = (f.flags ?? []).slice(0, 5);
  const avgCompLotRent = rentComps.length
    ? Math.round(rentComps.reduce((sum, c) => sum + (c.lotRent ?? 0), 0) / rentComps.length)
    : 0;

  async function handleDownload() {
    if (!result.download_url) {
      setDownloadError("No download URL on this result. Re-run the analysis.");
      return;
    }
    setDownloading(true);
    setDownloadError(null);
    try {
      await downloadExcel(result.download_url);
    } catch (e) {
      setDownloadError(e instanceof Error ? e.message : "Download failed.");
    } finally {
      setDownloading(false);
    }
  }

  return (
    <div className="card rounded-xl p-5 mb-6">
      <div className="flex items-center gap-3 mb-4">
        <div className="w-8 h-8 rounded-full bg-green-500/20 flex items-center justify-center">
          <svg className="w-4 h-4 text-green-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={3} d="M5 13l4 4L19 7" />
          </svg>
        </div>
        <h3 className="text-lg font-semibold">Analysis Complete</h3>
      </div>

      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-5">
        <KpiCard label="GGC NOI" value={formatMoney(ggcNoi)} />
        <KpiCard label="Cap Rate" value={(capRate * 100).toFixed(2) + "%"} />
        <KpiCard label="$ / Unit" value={formatMoney(pricePerUnit)} />
        <KpiCard label="Occupancy" value={occupancy} />
        {omittTotal !== 0 && (
          <KpiCard
            label="Omitt (excluded)"
            value={formatMoney(omittTotal)}
            subtitle={`+${formatMoney(omittIncome)} inc, −${formatMoney(omittExpense)} exp`}
          />
        )}
        {cost !== null && (
          <KpiCard
            label="API Cost"
            value={"$" + cost.toFixed(2)}
            subtitle={`${result.usage?.calls ?? 0} Claude calls`}
          />
        )}
      </div>

      <div className="grid md:grid-cols-2 gap-4 mb-5">
        <div className="bg-slate-900/50 rounded-lg p-4">
          <h4 className="text-xs font-semibold text-slate-400 uppercase mb-2">Diligence Flags</h4>
          <div className="space-y-2 text-sm">
            {flags.length === 0 ? (
              <span className="text-slate-500 text-xs">None flagged.</span>
            ) : (
              flags.map((flag, i) => (
                <div key={i} className="flex gap-2">
                  <span
                    className={`${severityClass(flag.severity)} text-xs font-bold uppercase mt-0.5`}
                  >
                    {flag.severity}
                  </span>
                  <span className="text-slate-300 flex-1">
                    {flag.item}: {flag.issue}
                  </span>
                </div>
              ))
            )}
          </div>
        </div>

        <div className="bg-slate-900/50 rounded-lg p-4">
          <h4 className="text-xs font-semibold text-slate-400 uppercase mb-2">Comps Found</h4>
          <div className="text-sm space-y-2">
            <div className="flex justify-between">
              <span className="text-slate-400">Rent comps:</span>
              <span className="font-semibold">{rentComps.length}</span>
            </div>
            <div className="flex justify-between">
              <span className="text-slate-400">Sale comps:</span>
              <span className="font-semibold">{saleComps.length}</span>
            </div>
            <div className="flex justify-between">
              <span className="text-slate-400">Avg comp lot rent:</span>
              <span className="font-semibold">${avgCompLotRent}</span>
            </div>
            <div className="flex justify-between">
              <span className="text-slate-400">Demand signal:</span>
              <span className="font-semibold">{m.demandSignal || "-"}</span>
            </div>
          </div>
        </div>
      </div>

      {(result.verification?.hardFails ?? 0) > 0 && (
        <div className="mb-3 rounded-lg border border-amber-500/40 bg-amber-500/10 p-3 text-xs">
          <div className="font-semibold text-amber-300 mb-1">
            Review before trusting — {result.verification?.hardFails} verification check
            {(result.verification?.hardFails ?? 0) === 1 ? "" : "s"} did not tie out.
          </div>
          <div className="text-slate-300">
            Workbook produced. Open the <strong>Extraction Check</strong> tab (first sheet)
            for details. Failed:{" "}
            <span className="text-amber-200">
              {(result.verification?.failedCheckNames ?? []).slice(0, 5).join(", ")}
            </span>
            {(result.verification?.failedCheckNames?.length ?? 0) > 5 && " (+more)"}
          </div>
        </div>
      )}

      <button
        onClick={() => void handleDownload()}
        disabled={downloading}
        className="btn-amber w-full py-3 rounded-lg font-semibold flex items-center justify-center gap-2"
      >
        <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path
            strokeLinecap="round"
            strokeLinejoin="round"
            strokeWidth={2}
            d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"
          />
        </svg>
        {downloading ? "Preparing download…" : "Download GGC UW Sizer (.xlsx)"}
      </button>
      {downloadError && <p className="text-sm text-red-400 mt-3">{downloadError}</p>}
    </div>
  );
}
