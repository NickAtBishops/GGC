"use client";

// New-deal flow: form → submit → live progress (poll /api/status every 4s) →
// results + download. The active job id is persisted in localStorage so a
// page refresh resumes polling.

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import DealForm from "@/components/DealForm";
import JobProgress, { type ProgressSteps } from "@/components/JobProgress";
import ResultsPanel from "@/components/ResultsPanel";
import {
  getJobStatus,
  startAnalysis,
  type DealFormFields,
  type JobResult,
} from "@/lib/engine";

const ACTIVE_JOB_KEY = "ggc_active_job";
const POLL_INTERVAL_MS = 4000;

interface StoredJob {
  jobId: string;
  startedAt: number;
}

function loadStoredJob(): StoredJob | null {
  try {
    const raw = localStorage.getItem(ACTIVE_JOB_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as Partial<StoredJob>;
    if (typeof parsed.jobId === "string" && typeof parsed.startedAt === "number") {
      return { jobId: parsed.jobId, startedAt: parsed.startedAt };
    }
    return null;
  } catch {
    return null;
  }
}

function saveStoredJob(job: StoredJob): void {
  try {
    localStorage.setItem(ACTIVE_JOB_KEY, JSON.stringify(job));
  } catch {
    // storage unavailable — polling still works for this tab session
  }
}

function clearStoredJob(): void {
  try {
    localStorage.removeItem(ACTIVE_JOB_KEY);
  } catch {
    // ignore
  }
}

type Phase = "idle" | "submitting" | "polling" | "complete" | "error";

export default function HomePage() {
  const [phase, setPhase] = useState<Phase>("idle");
  const [jobId, setJobId] = useState<string | null>(null);
  const [startedAt, setStartedAt] = useState<number | null>(null);
  const [progress, setProgress] = useState("");
  const [fillActive, setFillActive] = useState(false);
  const [result, setResult] = useState<JobResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const pollRef = useRef<ReturnType<typeof setInterval> | null>(null);

  const stopPolling = useCallback(() => {
    if (pollRef.current !== null) {
      clearInterval(pollRef.current);
      pollRef.current = null;
    }
  }, []);

  const beginPolling = useCallback(
    (id: string) => {
      stopPolling();

      const tick = async () => {
        try {
          const job = await getJobStatus(id);
          setProgress(job.progress ?? "");

          if (job.status === "complete") {
            stopPolling();
            clearStoredJob();
            setFillActive(true);
            if (job.result) {
              setResult(job.result);
              setPhase("complete");
            } else {
              setError("The job completed but the engine returned no result payload.");
              setPhase("error");
            }
          } else if (job.status === "error") {
            stopPolling();
            clearStoredJob();
            setError(job.error || "Analysis failed.");
            setPhase("error");
          } else if (job.progress && job.progress.toLowerCase().includes("filling")) {
            // Same heuristic as index.html: a "Filling…" progress string means
            // the parallel-analysis stage is done and template fill has begun.
            setFillActive(true);
          }
        } catch (e) {
          stopPolling();
          clearStoredJob();
          setError(e instanceof Error ? e.message : "Lost contact with the engine.");
          setPhase("error");
        }
      };

      void tick(); // immediate first check, then every 4 seconds
      pollRef.current = setInterval(() => void tick(), POLL_INTERVAL_MS);
    },
    [stopPolling],
  );

  // Resume an in-flight job after a page refresh.
  useEffect(() => {
    const stored = loadStoredJob();
    if (stored) {
      setJobId(stored.jobId);
      setStartedAt(stored.startedAt);
      setPhase("polling");
      beginPolling(stored.jobId);
    }
    return stopPolling;
  }, [beginPolling, stopPolling]);

  const handleSubmit = useCallback(
    (fields: DealFormFields, files: File[]) => {
      const startedAtMs = Date.now();
      setPhase("submitting");
      setJobId(null);
      setStartedAt(startedAtMs);
      setProgress("");
      setFillActive(false);
      setResult(null);
      setError(null);

      void (async () => {
        try {
          const id = await startAnalysis(fields, files);
          setJobId(id);
          saveStoredJob({ jobId: id, startedAt: startedAtMs });
          setPhase("polling");
          beginPolling(id);
        } catch (e) {
          setError(e instanceof Error ? e.message : "Upload failed.");
          setPhase("error");
        }
      })();
    },
    [beginPolling],
  );

  const steps: ProgressSteps = useMemo(() => {
    const s: ProgressSteps = {
      upload: "pending",
      parallel: "pending",
      fill: "pending",
      complete: "pending",
    };
    if (phase === "idle") return s;
    if (phase === "submitting" || (phase === "error" && !jobId)) {
      s.upload = "active";
      return s;
    }
    s.upload = "done";
    if (phase === "complete") {
      s.parallel = "done";
      s.fill = "done";
      s.complete = "done";
      return s;
    }
    s.parallel = fillActive ? "done" : "active";
    s.fill = fillActive ? "active" : "pending";
    return s;
  }, [phase, jobId, fillActive]);

  const busy = phase === "submitting" || phase === "polling";

  return (
    <>
      <DealForm busy={busy} onSubmit={handleSubmit} />

      {phase !== "idle" && (
        <JobProgress steps={steps} progress={progress} startedAt={startedAt} running={busy} />
      )}

      {phase === "complete" && result && <ResultsPanel result={result} />}

      {phase === "error" && error && (
        <div className="card rounded-xl p-5 mb-6 border-red-500/30">
          <div className="flex items-start gap-3">
            <svg
              className="w-5 h-5 text-red-400 shrink-0 mt-0.5"
              fill="none"
              stroke="currentColor"
              viewBox="0 0 24 24"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={2}
                d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"
              />
            </svg>
            <div className="flex-1">
              <h3 className="font-semibold text-red-400 mb-1">Analysis Failed</h3>
              <p className="text-sm text-slate-300">{error}</p>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
