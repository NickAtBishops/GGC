"use client";

// Live progress panel — same four pipeline steps, elapsed timer, and raw
// progress string as index.html's status panel.

import { useEffect, useState } from "react";

export type StepState = "pending" | "active" | "done";

export interface ProgressSteps {
  upload: StepState;
  parallel: StepState;
  fill: StepState;
  complete: StepState;
}

const STEP_LABELS: Array<{ key: keyof ProgressSteps; label: string }> = [
  { key: "upload", label: "Uploading files" },
  { key: "parallel", label: "Parallel AI analysis (financials + market research)" },
  { key: "fill", label: "Filling GGC template" },
  { key: "complete", label: "Ready to download" },
];

interface JobProgressProps {
  steps: ProgressSteps;
  /** Raw progress string from GET /api/status. */
  progress: string;
  /** Epoch ms when the run started (persisted, so it survives refresh). */
  startedAt: number | null;
  /** True while the job is uploading or queued/running. */
  running: boolean;
  /** True while a status check just failed but polling is retrying — the
   * job is presumed still alive server-side (see MAX_CONSECUTIVE_POLL_FAILURES
   * in app/page.tsx). Surfaced so a transient blip doesn't look like a
   * silent hang. */
  reconnecting?: boolean;
  /** Optional cancel handler — when present, a Cancel button renders. */
  onCancel?: () => void;
}

function StepIcon({ state }: { state: StepState }) {
  if (state === "done") {
    return (
      <div className="w-5 h-5 rounded-full bg-green-400 flex items-center justify-center shrink-0">
        <svg className="w-3 h-3 text-slate-900" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={4} d="M5 13l4 4L19 7" />
        </svg>
      </div>
    );
  }
  if (state === "active") {
    return <div className="w-5 h-5 rounded-full border-2 border-blue-400 pulse-dot shrink-0" />;
  }
  return <div className="w-5 h-5 rounded-full border-2 border-slate-600 shrink-0" />;
}

export default function JobProgress({
  steps,
  progress,
  startedAt,
  running,
  reconnecting,
  onCancel,
}: JobProgressProps) {
  const [now, setNow] = useState<number>(() => Date.now());

  useEffect(() => {
    if (!running) return;
    setNow(Date.now());
    const timer = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(timer);
  }, [running]);

  let elapsed = "";
  if (startedAt !== null) {
    const secs = Math.max(0, Math.floor((now - startedAt) / 1000));
    const m = Math.floor(secs / 60);
    const s = String(secs % 60).padStart(2, "0");
    elapsed = `${m}:${s}`;
  }

  const headerDot = running
    ? "pulse-dot bg-blue-400"
    : steps.complete === "done"
      ? "bg-green-400"
      : "bg-red-400";

  return (
    <div className="card rounded-xl p-5 mb-6">
      <div className="flex items-center gap-3 mb-4">
        <div className={`w-3 h-3 rounded-full ${headerDot}`} />
        <h3 className="font-semibold">Processing</h3>
        <span className="text-xs text-slate-500 ml-auto">{elapsed}</span>
      </div>
      <div className="space-y-2">
        {STEP_LABELS.map(({ key, label }) => {
          const state = steps[key];
          const textClass =
            state === "done"
              ? "text-green-400"
              : state === "active"
                ? "text-blue-400"
                : "text-slate-400";
          return (
            <div key={key} className={`flex items-center gap-3 text-sm ${textClass}`}>
              <StepIcon state={state} />
              {label}
            </div>
          );
        })}
      </div>
      {progress ? <div className="mt-4 text-xs text-slate-500 font-mono">{progress}</div> : null}
      {reconnecting ? (
        <div className="mt-2 text-xs text-amber-400">
          Lost contact with the engine — retrying…
        </div>
      ) : null}
      {running && onCancel ? (
        <button
          type="button"
          onClick={onCancel}
          className="mt-4 w-full py-2 rounded-lg text-sm font-semibold text-red-300 border border-red-500/40 bg-red-500/10 hover:bg-red-500/20 transition-colors"
        >
          Cancel analysis
        </button>
      ) : null}
    </div>
  );
}
