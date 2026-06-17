"use client";

// New-deal form: property info + document drop zone. Mirrors index.html —
// same field names/placeholders, same required-field validation, same file
// constraints, and the same localStorage persistence of property fields.

import {
  useEffect,
  useRef,
  useState,
  type ChangeEvent,
  type DragEvent,
  type FormEvent,
} from "react";
import type { DealFormFields, FloodZone } from "@/lib/engine";

// Fields persisted across reloads. The REQUIRED subset gates the submit
// button; the rest persist for convenience but can be left blank — the
// engine derives them from the uploaded documents (units from the rent
// roll, etc.).
const PERSISTED_FIELDS = [
  "property_name",
  "address",
  "city",
  "state",
  "county",
  "units",
  "asking_price",
] as const;

const REQUIRED_FIELDS: readonly (typeof PERSISTED_FIELDS)[number][] = [
  "property_name",
  "address",
  "city",
  "state",
  "county",
  "asking_price",
];

const ALLOWED_EXTS = ["pdf", "xlsx", "xls", "csv", "png", "jpg", "jpeg", "txt", "md"];
const MAX_FILE_BYTES = 50 * 1024 * 1024; // 50 MB per file
const MAX_TOTAL_BYTES = 150 * 1024 * 1024; // 150 MB across all uploads

const EMPTY_FIELDS: DealFormFields = {
  property_name: "",
  address: "",
  city: "",
  state: "",
  county: "",
  county_tax_rate: "",
  units: "",
  poh_count: "",
  asking_price: "",
  flood_zone: "unknown",
  deep_search: false,
  cost_mode: "max",
  n_runs: "3",
  skip_market: false,
};

// Per-run cost estimate (USD), used for the "≈ $X.XX" hint on the form so
// the user sees what they're about to spend before clicking Analyze. These
// are rough order-of-magnitude estimates based on typical token counts on
// MHC deals; actual cost is reported on the result panel after the run.
function estimateCostUsd(mode: string, nRuns: number, skipMarket: boolean): number {
  // Per-call cost ≈ $0.20 on Haiku, $2.50 on Opus (input + output + thinking).
  const perCall: Record<string, { ext: number; meth: number; mkt: number }> = {
    economy:  { ext: 0.20, meth: 0.20, mkt: 0.30 },
    balanced: { ext: 0.20, meth: 2.50, mkt: 2.50 },
    max:      { ext: 2.50, meth: 2.50, mkt: 2.50 },
  };
  const p = perCall[mode] ?? perCall.max;
  const market = skipMarket ? 0 : p.mkt;
  return p.ext * nRuns + p.meth * nRuns + market;
}

interface DealFormProps {
  /** True while a job is uploading or running — disables submission. */
  busy: boolean;
  onSubmit: (fields: DealFormFields, files: File[]) => void;
}

export default function DealForm({ busy, onSubmit }: DealFormProps) {
  const [fields, setFields] = useState<DealFormFields>(EMPTY_FIELDS);
  const [files, setFiles] = useState<File[]>([]);
  const [fileErrors, setFileErrors] = useState<string[]>([]);
  const [dragging, setDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Restore persisted property fields after mount (client-only, avoids
  // SSR/hydration mismatch).
  useEffect(() => {
    setFields((prev) => {
      const next = { ...prev };
      for (const key of PERSISTED_FIELDS) {
        try {
          const saved = localStorage.getItem(`property_${key}`);
          if (saved !== null) next[key] = saved;
        } catch {
          // localStorage unavailable — keep defaults
        }
      }
      return next;
    });
  }, []);

  function setField<K extends keyof DealFormFields>(key: K, value: DealFormFields[K]) {
    setFields((prev) => ({ ...prev, [key]: value }));
    if ((PERSISTED_FIELDS as readonly string[]).includes(key) && typeof value === "string") {
      try {
        localStorage.setItem(`property_${key}`, value);
      } catch {
        // ignore storage failures
      }
    }
  }

  function addFiles(incoming: ArrayLike<File>) {
    const errors: string[] = [];
    const accepted: File[] = [];
    let runningTotal = files.reduce((sum, f) => sum + f.size, 0);

    for (const f of Array.from(incoming)) {
      const ext = (f.name.split(".").pop() || "").toLowerCase();
      if (!ALLOWED_EXTS.includes(ext)) {
        errors.push(`${f.name}: unsupported file type .${ext}`);
        continue;
      }
      if (f.size > MAX_FILE_BYTES) {
        errors.push(`${f.name}: ${(f.size / 1024 / 1024).toFixed(1)} MB exceeds 50 MB limit`);
        continue;
      }
      if (runningTotal + f.size > MAX_TOTAL_BYTES) {
        errors.push(`${f.name}: would exceed 150 MB total upload limit`);
        continue;
      }
      runningTotal += f.size;
      accepted.push(f);
    }

    if (accepted.length > 0) setFiles((prev) => [...prev, ...accepted]);
    setFileErrors(errors);
  }

  function removeFile(index: number) {
    setFiles((prev) => prev.filter((_, i) => i !== index));
  }

  function onFileInputChange(e: ChangeEvent<HTMLInputElement>) {
    if (e.target.files && e.target.files.length > 0) addFiles(e.target.files);
    e.target.value = ""; // allow re-selecting the same file
  }

  function onDragOver(e: DragEvent<HTMLDivElement>) {
    e.preventDefault();
    setDragging(true);
  }

  function onDragLeave() {
    setDragging(false);
  }

  function onDrop(e: DragEvent<HTMLDivElement>) {
    e.preventDefault();
    setDragging(false);
    addFiles(e.dataTransfer.files);
  }

  const allRequiredFilled = REQUIRED_FIELDS.every((key) => fields[key].trim().length > 0);
  const canSubmit = allRequiredFilled && files.length > 0 && !busy;

  function handleSubmit(e: FormEvent<HTMLFormElement>) {
    e.preventDefault();
    if (!canSubmit) return;
    onSubmit(fields, files);
  }

  return (
    <form onSubmit={handleSubmit}>
      <div className="grid md:grid-cols-2 gap-6 mb-6">
        {/* Property Info */}
        <div className="card rounded-xl p-5">
          <h2 className="text-sm font-semibold mb-4 flex items-center gap-2">
            <svg className="w-4 h-4 text-blue-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={2}
                d="M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z"
              />
            </svg>
            Property Information
          </h2>

          <div className="space-y-3">
            <input
              value={fields.property_name}
              onChange={(e) => setField("property_name", e.target.value)}
              placeholder="Property Name"
              className="input-field w-full rounded-lg px-4 py-2 text-sm"
            />
            <input
              value={fields.address}
              onChange={(e) => setField("address", e.target.value)}
              placeholder="Street Address"
              className="input-field w-full rounded-lg px-4 py-2 text-sm"
            />
            <div className="grid grid-cols-3 gap-2">
              <input
                value={fields.city}
                onChange={(e) => setField("city", e.target.value)}
                placeholder="City"
                className="input-field col-span-2 rounded-lg px-4 py-2 text-sm"
              />
              <input
                value={fields.state}
                onChange={(e) => setField("state", e.target.value)}
                placeholder="ST"
                maxLength={2}
                className="input-field rounded-lg px-4 py-2 text-sm uppercase"
              />
            </div>
            <div className="grid grid-cols-2 gap-2">
              <input
                value={fields.county}
                onChange={(e) => setField("county", e.target.value)}
                placeholder="County"
                title="County name — used for tax reassessment lookup and demographics."
                className="input-field rounded-lg px-4 py-2 text-sm"
              />
              <input
                value={fields.county_tax_rate}
                onChange={(e) => setField("county_tax_rate", e.target.value)}
                type="number"
                step="0.0001"
                placeholder="County Tax Rate (e.g. 0.0125)"
                title="Effective millage / tax rate as a decimal (1.25% = 0.0125). Leave blank if unknown; backend will fall back to T12 × 1.15."
                className="input-field rounded-lg px-4 py-2 text-sm"
              />
            </div>
            <div className="grid grid-cols-3 gap-2">
              <input
                value={fields.units}
                onChange={(e) => setField("units", e.target.value)}
                type="number"
                placeholder="Total Units (optional)"
                title="Leave blank to auto-derive from the rent roll. Enter only if you want the engine to cross-check the rent-roll row count against your number."
                className="input-field rounded-lg px-4 py-2 text-sm"
              />
              <input
                value={fields.poh_count}
                onChange={(e) => setField("poh_count", e.target.value)}
                type="number"
                placeholder="POH Count"
                title="Number of Park-Owned Homes. Enter 0 if all tenant-owned (TOH)."
                className="input-field rounded-lg px-4 py-2 text-sm"
              />
              <input
                value={fields.asking_price}
                onChange={(e) => setField("asking_price", e.target.value)}
                type="number"
                placeholder="Asking Price ($)"
                className="input-field rounded-lg px-4 py-2 text-sm"
              />
            </div>

            <div className="pt-1">
              <label
                htmlFor="flood_zone"
                className="flex items-center justify-between text-xs font-medium text-slate-400 mb-1.5"
              >
                <span className="flex items-center gap-1.5">
                  <svg
                    className="w-3.5 h-3.5 text-blue-400"
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      strokeWidth={2}
                      d="M3 15l4-4 4 4 5-5 5 5M3 15v4a2 2 0 002 2h14a2 2 0 002-2v-4M3 15V7a2 2 0 012-2h14a2 2 0 012 2v8"
                    />
                  </svg>
                  FEMA Flood Zone
                </span>
                <a
                  href="https://msc.fema.gov/portal/home"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="text-blue-400 hover:text-blue-300 transition-colors flex items-center gap-1"
                >
                  Check map
                  <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      strokeWidth={2}
                      d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14"
                    />
                  </svg>
                </a>
              </label>
              <select
                id="flood_zone"
                name="flood_zone"
                value={fields.flood_zone}
                onChange={(e) => setField("flood_zone", e.target.value as FloodZone)}
                className="input-field w-full rounded-lg px-4 py-2 text-sm appearance-none cursor-pointer"
              >
                <option value="unknown">Unknown — flag as diligence item</option>
                <option value="no">No special flood hazard (Zone X)</option>
                <option value="yes">
                  Special hazard zone (A, AE, V, etc.) — apply 15% insurance trend
                </option>
              </select>
            </div>

            <div className="flex items-center justify-between p-3 rounded-lg bg-slate-900/40 border border-slate-700/50">
              <div className="flex items-center gap-2">
                <svg
                  className="w-4 h-4 text-blue-400 shrink-0"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d="M9 12l2 2 4-4M7.835 4.697a3.42 3.42 0 001.946-.806 3.42 3.42 0 014.438 0 3.42 3.42 0 001.946.806 3.42 3.42 0 013.138 3.138 3.42 3.42 0 00.806 1.946 3.42 3.42 0 010 4.438 3.42 3.42 0 00-.806 1.946 3.42 3.42 0 01-3.138 3.138 3.42 3.42 0 00-1.946.806 3.42 3.42 0 01-4.438 0 3.42 3.42 0 00-1.946-.806 3.42 3.42 0 01-3.138-3.138 3.42 3.42 0 00-.806-1.946 3.42 3.42 0 010-4.438 3.42 3.42 0 00.806-1.946 3.42 3.42 0 013.138-3.138z"
                  />
                </svg>
                <div>
                  <div className="text-xs font-medium text-slate-300">Run 3× and reconcile</div>
                  <div className="text-[10px] text-slate-500">
                    Three parallel extractions, majority-vote per field. Catches rent-roll dropouts
                    &amp; category drift. ~3× API cost (~$6/deal).
                  </div>
                </div>
              </div>
              <label className="relative inline-flex items-center cursor-pointer">
                <input
                  type="checkbox"
                  checked={fields.deep_search}
                  onChange={(e) => setField("deep_search", e.target.checked)}
                  className="sr-only peer"
                />
                <div
                  className="w-9 h-5 bg-slate-700 rounded-full
                             peer-checked:after:translate-x-full peer-checked:after:border-white
                             after:content-[''] after:absolute after:top-[2px] after:left-[2px]
                             after:bg-white after:border after:border-slate-300 after:rounded-full
                             after:h-4 after:w-4 after:transition-all peer-checked:bg-blue-500"
                />
              </label>
            </div>
          </div>
        </div>

        {/* File Upload */}
        <div className="card rounded-xl p-5">
          <h2 className="text-sm font-semibold mb-4 flex items-center gap-2">
            <svg className="w-4 h-4 text-blue-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={2}
                d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"
              />
            </svg>
            Deal Documents
          </h2>

          <div
            onClick={() => fileInputRef.current?.click()}
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
            onDrop={onDrop}
            className={`drop-zone border-2 border-dashed border-slate-600 rounded-lg p-6 text-center cursor-pointer hover:border-blue-500 ${
              dragging ? "dragging" : ""
            }`}
          >
            <input
              ref={fileInputRef}
              type="file"
              multiple
              accept=".pdf,.xlsx,.xls,.csv,.png,.jpg,.jpeg"
              onChange={onFileInputChange}
              className="hidden"
            />
            <svg
              className="w-10 h-10 mx-auto mb-2 text-slate-500"
              fill="none"
              stroke="currentColor"
              viewBox="0 0 24 24"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={1.5}
                d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
              />
            </svg>
            <p className="text-sm text-slate-400">Drop T12, rent roll, or P&amp;L files here</p>
            <p className="text-xs text-slate-500 mt-1">PDF, Excel, CSV, or images</p>
          </div>

          <div className="mt-3 space-y-1">
            {files.map((f, i) => (
              <div
                key={`${f.name}-${f.size}-${i}`}
                className="flex items-center justify-between bg-slate-900/40 rounded px-3 py-2 text-sm"
              >
                <span className="truncate flex-1">{f.name}</span>
                <span className="text-xs text-slate-500 mx-3 shrink-0">
                  {(f.size / 1024).toFixed(0)} KB
                </span>
                <button
                  type="button"
                  onClick={() => removeFile(i)}
                  aria-label={`Remove ${f.name}`}
                  className="text-slate-500 hover:text-red-400"
                >
                  ✕
                </button>
              </div>
            ))}
          </div>

          {fileErrors.length > 0 && (
            <div className="mt-3 rounded-lg border border-red-500/30 bg-red-500/10 p-3 text-xs text-red-300 space-y-1">
              <div className="font-semibold text-red-400">Some files were rejected:</div>
              {fileErrors.map((msg, i) => (
                <div key={i}>{msg}</div>
              ))}
            </div>
          )}
        </div>
      </div>

      {/* Cost controls — lets the user trade accuracy for spend per-run.
          Cost estimate below the controls so the spend is visible before
          clicking Analyze. */}
      <div className="mb-4 card rounded-xl p-4 border-slate-700/50">
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-sm font-semibold text-slate-300">Run mode</h3>
          <span className="text-xs text-slate-400">
            est. ≈ ${estimateCostUsd(
              fields.cost_mode,
              parseInt(fields.n_runs, 10) || 3,
              fields.skip_market,
            ).toFixed(2)} / run
          </span>
        </div>
        <div className="grid grid-cols-3 gap-2 text-xs">
          <div>
            <label className="block text-slate-400 mb-1">Model quality</label>
            <select
              value={fields.cost_mode}
              onChange={(e) => setField("cost_mode", e.target.value as "economy" | "balanced" | "max")}
              className="input-field rounded-lg px-3 py-2 w-full"
              title="Economy = Haiku 4.5 everywhere (cheap, decent). Balanced = Haiku for transcription + Opus for judgment. Max = Opus everywhere (most accurate)."
            >
              <option value="economy">Economy (Haiku)</option>
              <option value="balanced">Balanced</option>
              <option value="max">Max accuracy (Opus)</option>
            </select>
          </div>
          <div>
            <label className="block text-slate-400 mb-1">Self-consistency runs</label>
            <select
              value={fields.n_runs}
              onChange={(e) => setField("n_runs", e.target.value as "1" | "3" | "5")}
              className="input-field rounded-lg px-3 py-2 w-full"
              title="N parallel runs per stage, voted to consensus. 1 disables voting (cheapest, less deterministic). 3 is the recommended default. 5 is max accuracy."
            >
              <option value="1">1 (no voting, cheap)</option>
              <option value="3">3 (recommended)</option>
              <option value="5">5 (max determinism)</option>
            </select>
          </div>
          <div>
            <label className="block text-slate-400 mb-1">Market research</label>
            <select
              value={fields.skip_market ? "skip" : "run"}
              onChange={(e) => setField("skip_market", e.target.value === "skip")}
              className="input-field rounded-lg px-3 py-2 w-full"
              title="Skips the web_search comps lookup. Saves the priciest single call but leaves the Comps Analysis tab blank."
            >
              <option value="run">Run market research</option>
              <option value="skip">Skip (save ~$2.50)</option>
            </select>
          </div>
        </div>
      </div>

      {/* Action Button */}
      <div className="mb-6">
        <button
          type="submit"
          disabled={!canSubmit}
          className="btn-primary w-full py-3.5 rounded-xl text-base font-semibold text-white flex items-center justify-center gap-2"
        >
          {busy ? (
            <>
              <span className="pulse-dot w-3 h-3 rounded-full bg-white" />
              Analyzing…
            </>
          ) : (
            <>
              <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth={2}
                  d="M13 10V3L4 14h7v7l9-11h-7z"
                />
              </svg>
              Analyze Deal
            </>
          )}
        </button>
      </div>
    </form>
  );
}
