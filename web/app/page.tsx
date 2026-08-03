"use client";

// New-deal flow: form → submit → live progress (poll /api/status every 4s) →
// results + download. The active job id is persisted in localStorage so a
// page refresh resumes polling.

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import ApiKeyCard from "@/components/ApiKeyCard";
import DealForm from "@/components/DealForm";
import DocAiKeyCard from "@/components/DocAiKeyCard";
import JobProgress, { type ProgressSteps } from "@/components/JobProgress";
import ResultsPanel from "@/components/ResultsPanel";
import {
  cancelJob,
  EMPTY_GCP_CONFIG,
  getConfig,
  getJobStatus,
  startAnalysis,
  type DealFormFields,
  type GcpDocAiConfig,
  type JobResult,
} from "@/lib/engine";

const ACTIVE_JOB_KEY = "ggc_active_job";
const API_KEY_STORAGE_KEY = "anthropic_api_key";
const GCP_CONFIG_STORAGE_KEY = "gcp_doc_ai_config";
const POLL_INTERVAL_MS = 4000;
// A single failed status check (e.g. a Cloud Run revision cutover, or a
// momentary network blip) used to kill an otherwise-healthy job outright.
// Tolerate a run of transient failures before giving up — the job itself
// lives server-side and is almost always still fine seconds later.
const MAX_CONSECUTIVE_POLL_FAILURES = 5;

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
  const [reconnecting, setReconnecting] = useState(false);
  const [apiKey, setApiKey] = useState("");
  const [defaultKeyPresent, setDefaultKeyPresent] = useState(false);
  const [gcpConfig, setGcpConfig] = useState<GcpDocAiConfig>(EMPTY_GCP_CONFIG);
  const [defaultDocAiPresent, setDefaultDocAiPresent] = useState(false);
  const pollRef = useRef<ReturnType<typeof setInterval> | null>(null);
  const consecutiveFailuresRef = useRef(0);

  // Restore the visitor's own key/GCP config (if saved) and check what the
  // server has configured by default, so the form knows whether a key is
  // required and the status badges reflect reality on first paint.
  useEffect(() => {
    try {
      const savedKey = localStorage.getItem(API_KEY_STORAGE_KEY);
      if (savedKey) setApiKey(savedKey);
      const savedGcp = localStorage.getItem(GCP_CONFIG_STORAGE_KEY);
      if (savedGcp) setGcpConfig({ ...EMPTY_GCP_CONFIG, ...JSON.parse(savedGcp) });
    } catch {
      // localStorage unavailable — keys still work for this session
    }
    getConfig()
      .then((cfg) => {
        setDefaultKeyPresent(Boolean(cfg.default_api_key_present));
        setDefaultDocAiPresent(Boolean(cfg.default_doc_ai_present));
      })
      .catch(() => {});
  }, []);

  function handleApiKeyChange(key: string) {
    setApiKey(key);
    try {
      if (key) localStorage.setItem(API_KEY_STORAGE_KEY, key);
      else localStorage.removeItem(API_KEY_STORAGE_KEY);
    } catch {
      // localStorage unavailable — key still works for this session
    }
  }

  function handleGcpConfigChange(config: GcpDocAiConfig) {
    setGcpConfig(config);
    try {
      const isEmpty = Object.values(config).every((v) => !v);
      if (isEmpty) localStorage.removeItem(GCP_CONFIG_STORAGE_KEY);
      else localStorage.setItem(GCP_CONFIG_STORAGE_KEY, JSON.stringify(config));
    } catch {
      // localStorage unavailable — config still works for this session
    }
  }

  const stopPolling = useCallback(() => {
    if (pollRef.current !== null) {
      clearInterval(pollRef.current);
      pollRef.current = null;
    }
  }, []);

  const beginPolling = useCallback(
    (id: string) => {
      stopPolling();
      consecutiveFailuresRef.current = 0;
      setReconnecting(false);

      const tick = async () => {
        try {
          const job = await getJobStatus(id);
          consecutiveFailuresRef.current = 0;
          setReconnecting(false);
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
          } else if (job.status === "needs_review") {
            // Verification gate fired — engine refused to produce a workbook
            // because hard checks did not tie out. Show the failing checks in
            // the error panel so the user knows what to fix and retry.
            stopPolling();
            clearStoredJob();
            const v = job.result?.verification;
            const names = v?.failedCheckNames ?? [];
            const msg =
              job.result?.message ||
              `Verification failed: ${v?.hardFails ?? 0} hard checks did not tie out.` +
                (names.length ? ` Failed: ${names.slice(0, 5).join(", ")}.` : "") +
                " No workbook was produced — re-upload corrected docs or adjust inputs and retry.";
            setError(msg);
            setPhase("error");
          } else if (job.status === "error") {
            stopPolling();
            clearStoredJob();
            setError(job.error || "Analysis failed.");
            setPhase("error");
          } else if (job.status === "cancelled") {
            stopPolling();
            clearStoredJob();
            setError(job.error || "Cancelled by user.");
            setPhase("error");
          } else if (job.progress && job.progress.toLowerCase().includes("filling")) {
            // Same heuristic as index.html: a "Filling…" progress string means
            // the parallel-analysis stage is done and template fill has begun.
            setFillActive(true);
          }
        } catch (e) {
          consecutiveFailuresRef.current += 1;
          if (consecutiveFailuresRef.current < MAX_CONSECUTIVE_POLL_FAILURES) {
            // Likely transient (a Cloud Run revision cutover, a momentary
            // network blip) — the job is still running server-side. Keep
            // polling instead of throwing away an otherwise-healthy job.
            setReconnecting(true);
            return;
          }
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
          const id = await startAnalysis(fields, files, apiKey, gcpConfig);
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
    [apiKey, gcpConfig, beginPolling],
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
      <ApiKeyCard apiKey={apiKey} defaultKeyPresent={defaultKeyPresent} onApiKeyChange={handleApiKeyChange} />
      <DocAiKeyCard
        gcpConfig={gcpConfig}
        defaultDocAiPresent={defaultDocAiPresent}
        onChange={handleGcpConfigChange}
      />
      <DealForm
        busy={busy}
        apiKeyReady={Boolean(apiKey) || defaultKeyPresent}
        onSubmit={handleSubmit}
      />

      {phase !== "idle" && (
        <JobProgress
          steps={steps}
          progress={progress}
          startedAt={startedAt}
          running={busy}
          reconnecting={reconnecting}
          onCancel={
            jobId
              ? () => {
                  void cancelJob(jobId).catch((e) => {
                    setError(e instanceof Error ? e.message : "Cancel failed.");
                  });
                }
              : undefined
          }
        />
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
