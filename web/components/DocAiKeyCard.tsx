"use client";

// Lets a visitor supply their own Google Document AI project instead of the
// server's default — same pattern as ApiKeyCard, stored only in this
// browser's localStorage and sent with each /api/analyze request. Unlike
// the Anthropic key, this is optional even when you have no server default:
// the engine falls back to sending PDFs to Claude directly (lower table
// fidelity on dense financial statements, but it still works), so this
// section stays collapsed by default.

import type { GcpDocAiConfig } from "@/lib/engine";

interface DocAiKeyCardProps {
  gcpConfig: GcpDocAiConfig;
  defaultDocAiPresent: boolean;
  onChange: (config: GcpDocAiConfig) => void;
}

export default function DocAiKeyCard({ gcpConfig, defaultDocAiPresent, onChange }: DocAiKeyCardProps) {
  const hasOwnConfig = Boolean(
    gcpConfig.project_id || gcpConfig.processor_id || gcpConfig.credentials_json,
  );
  const statusLabel = hasOwnConfig
    ? "Using your Document AI"
    : defaultDocAiPresent
      ? "Using shared Document AI"
      : "No Document AI (PDFs sent to Claude directly)";

  function setField<K extends keyof GcpDocAiConfig>(key: K, value: string) {
    onChange({ ...gcpConfig, [key]: value });
  }

  return (
    <details className="card rounded-xl p-5 mb-6 group">
      <summary className="flex items-center justify-between cursor-pointer list-none">
        <span className="flex items-center gap-2 text-sm font-semibold">
          <svg className="w-4 h-4 text-blue-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              strokeWidth={2}
              d="M9 13h6m-3-3v6m-9 1V7a2 2 0 012-2h3.28a2 2 0 001.68-.9l.9-1.35a2 2 0 011.68-.9h2.88a2 2 0 011.68.9l.9 1.35a2 2 0 001.68.9H19a2 2 0 012 2v10a2 2 0 01-2 2H5a2 2 0 01-2-2z"
            />
          </svg>
          Google Document AI (optional)
          <span className="text-slate-500 font-normal">— advanced, PDF table parsing</span>
        </span>
        <div className="flex items-center gap-2 px-3 py-1.5 rounded-full border bg-slate-800/40 border-slate-700/50 shrink-0">
          <div className={`w-2 h-2 rounded-full ${hasOwnConfig ? "bg-green-400" : "bg-slate-500"}`} />
          <span className="text-xs text-slate-400">{statusLabel}</span>
        </div>
      </summary>

      <div className="mt-4 space-y-3">
        <p className="text-xs text-slate-500">
          Only needed if you want PDF parsing billed to your own Google Cloud project instead of
          the shared one. Leave everything blank to keep using
          {defaultDocAiPresent ? " the shared Document AI project." : " Claude's direct PDF reading (no Document AI configured)."}
        </p>
        <div className="grid md:grid-cols-2 gap-2">
          <input
            value={gcpConfig.project_id}
            onChange={(e) => setField("project_id", e.target.value)}
            placeholder="GCP Project ID"
            className="input-field w-full rounded-lg px-4 py-2 text-sm"
          />
          <input
            value={gcpConfig.location}
            onChange={(e) => setField("location", e.target.value)}
            placeholder="Location (e.g. us)"
            className="input-field w-full rounded-lg px-4 py-2 text-sm"
          />
          <input
            value={gcpConfig.processor_id}
            onChange={(e) => setField("processor_id", e.target.value)}
            placeholder="Layout Processor ID"
            className="input-field w-full rounded-lg px-4 py-2 text-sm md:col-span-2"
          />
        </div>
        <textarea
          value={gcpConfig.credentials_json}
          onChange={(e) => setField("credentials_json", e.target.value)}
          placeholder='Service account JSON key — paste the full contents of the .json file, e.g. {"type": "service_account", ...}'
          rows={3}
          spellCheck={false}
          className="input-field w-full rounded-lg px-4 py-2 text-sm font-mono text-xs"
        />
        <p className="text-xs text-slate-500">
          Stored only in this browser and sent with each analysis request — never saved on the
          server. Needs a GCP project with the Document AI API enabled, a Layout Parser processor,
          and a service account key with the Document AI API User role.
        </p>
      </div>
    </details>
  );
}
