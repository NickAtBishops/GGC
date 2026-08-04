// Typed fetch helpers for the deal-engine endpoints:
//   GET  /api/config         → { default_api_key_present, google_maps_enabled }
//   POST /api/analyze        → { job_id }
//   GET  /api/status/{id}    → { status, progress, result, error? }
//   GET  /api/download/{id}  → .xlsx attachment

const ENGINE_URL = (
  process.env.NEXT_PUBLIC_ENGINE_URL || "http://localhost:5001"
).replace(/\/+$/, "");

// ─────────────────────────── Form fields (analyze) ──────────────────────────

export type FloodZone = "unknown" | "no" | "yes";

/** Exact multipart field names the engine expects (mirrors index.html). */
export interface DealFormFields {
  property_name: string;
  address: string;
  city: string;
  state: string;
  county: string;
  county_tax_rate: string;
  /** Per-site RE tax assumption ($/unit/year). Optional. When present,
   *  the engine writes it to `GGC Underwriting!J22` so the template's
   *  `I22 = J22 × N7` formula yields the underwritten RE Taxes. Typical
   *  range is $100–$2,000 depending on county (Whaleshead was $400). */
  tax_per_site: string;
  units: string;
  poh_count: string;
  asking_price: string;
  flood_zone: FloodZone;
  deep_search: boolean;
  /** "economy" = Haiku only, "balanced" = Haiku extract + Opus methodology,
   *  "max" = Opus all stages (current default). */
  cost_mode: "economy" | "balanced" | "max";
  /** Self-consistency runs per stage. "1" disables voting (cheap), "3"
   *  is the recommended default, "5" is max accuracy. */
  n_runs: "1" | "3" | "5";
  /** Skip the market-research Claude call (cheapest single saving). */
  skip_market: boolean;
}

// ───────────────────── Engine response types (contract) ─────────────────────

export interface FinancialLine {
  sellerLabel?: string;
  ggcCategory?: string;
  ggcUnderwritten?: number;
}

export interface DiligenceFlag {
  severity?: string;
  item?: string;
  issue?: string;
}

export interface PropertyInfo {
  askingPrice?: number;
  totalUnits?: number;
}

export interface RentRollSummary {
  occupancyRate?: number;
}

export interface Financials {
  propertyInfo?: PropertyInfo;
  rentRoll?: RentRollSummary;
  income?: FinancialLine[];
  expenses?: FinancialLine[];
  flags?: DiligenceFlag[];
}

export interface RentComp {
  lotRent?: number;
}

export interface MarketData {
  rentComps?: RentComp[];
  saleComps?: unknown[];
  demandSignal?: string;
}

export interface UsageTotals {
  cost_usd?: number;
  input_tokens?: number;
  output_tokens?: number;
  cache_read_input_tokens?: number;
  cache_creation_input_tokens?: number;
}

export interface UsagePerModel {
  calls: number;
  cost_usd: number;
}

export interface Usage {
  totals?: UsageTotals;
  calls?: number;
  /** Cost + call count broken out by model id (e.g. "claude-haiku-4-5" vs
   * "claude-opus-4-8") — lets the UI show which stage actually drove spend
   * instead of just a single opaque total. */
  per_model?: Record<string, UsagePerModel>;
}

export interface VerificationSummary {
  hardFails?: number;
  warnings?: number;
  failedCheckNames?: string[];
}

export interface JobResult {
  financials?: Financials;
  market?: MarketData;
  /**
   * Engine-relative path, e.g. "/api/download/<job_id>". Absent when the
   * job ended in `needs_review` — verification blocked the write-back.
   */
  download_url?: string;
  usage?: Usage;
  verification?: VerificationSummary;
  /** Human-readable note set by the engine on `needs_review`. */
  message?: string;
}

export type JobState =
  | "queued"
  | "running"
  | "complete"
  | "error"
  | "needs_review"
  | "cancelled";

export async function cancelJob(jobId: string): Promise<void> {
  const res = await engineFetch(`/api/cancel/${jobId}`, { method: "POST" });
  if (!res.ok) {
    const data = await readJson(res);
    throw new Error(serverError(data, `Cancel failed (HTTP ${res.status}).`));
  }
}

export interface JobStatus {
  status: JobState;
  progress?: string;
  result?: JobResult | null;
  error?: string;
}

// ─────────────────────────────── Internals ──────────────────────────────────

function resolveEngineUrl(pathOrUrl: string): string {
  if (/^https?:\/\//i.test(pathOrUrl)) return pathOrUrl;
  return `${ENGINE_URL}${pathOrUrl.startsWith("/") ? "" : "/"}${pathOrUrl}`;
}

async function engineFetch(pathOrUrl: string, init: RequestInit): Promise<Response> {
  try {
    return await fetch(resolveEngineUrl(pathOrUrl), init);
  } catch {
    throw new Error(
      `Could not reach the deal engine at ${ENGINE_URL}. ` +
        `Check that it is running and that NEXT_PUBLIC_ENGINE_URL is correct.`,
    );
  }
}

async function readJson(res: Response): Promise<Record<string, unknown> | null> {
  try {
    return (await res.json()) as Record<string, unknown>;
  } catch {
    return null;
  }
}

function serverError(data: Record<string, unknown> | null, fallback: string): string {
  return data && typeof data.error === "string" && data.error ? data.error : fallback;
}

// ─────────────────────────────── Endpoints ──────────────────────────────────

export interface EngineConfig {
  default_api_key_present?: boolean;
  google_maps_enabled?: boolean;
  default_doc_ai_present?: boolean;
}

/**
 * GET /api/config — whether the server has its own Anthropic key configured.
 * The key itself is never returned (see backend.py's `/api/config` docstring
 * for why); callers use this only to decide whether a visitor's own key is
 * required or optional.
 */
export async function getConfig(): Promise<EngineConfig> {
  const res = await engineFetch("/api/config", { method: "GET" });
  const data = await readJson(res);
  if (!res.ok || !data) throw new Error("Could not reach the deal engine.");
  return data as EngineConfig;
}

/**
 * POST /api/analyze — multipart upload. Resolves to the new job id.
 * `apiKey` is the caller's own Anthropic key (from browser localStorage);
 * pass an empty string to fall back to the server's default key, if any.
 */
export async function startAnalysis(
  fields: DealFormFields,
  files: File[],
  apiKey: string,
): Promise<string> {
  const fd = new FormData();
  fd.append("api_key", apiKey);
  fd.append("property_name", fields.property_name);
  fd.append("address", fields.address);
  fd.append("city", fields.city);
  fd.append("state", fields.state);
  fd.append("county", fields.county);
  fd.append("county_tax_rate", fields.county_tax_rate);
  fd.append("tax_per_site", fields.tax_per_site);
  fd.append("units", fields.units);
  fd.append("poh_count", fields.poh_count || "0");
  fd.append("asking_price", fields.asking_price);
  fd.append("flood_zone", fields.flood_zone);
  fd.append("deep_search", fields.deep_search ? "on" : "off");
  fd.append("cost_mode", fields.cost_mode);
  fd.append("n_runs", fields.n_runs);
  fd.append("skip_market", fields.skip_market ? "1" : "0");
  for (const file of files) fd.append("files", file);

  const res = await engineFetch("/api/analyze", { method: "POST", body: fd });
  const data = await readJson(res);

  if (!res.ok) {
    throw new Error(serverError(data, `Upload failed (HTTP ${res.status}).`));
  }
  if (!data || typeof data.job_id !== "string" || !data.job_id) {
    throw new Error("The engine did not return a job id.");
  }
  return data.job_id;
}

/** GET /api/status/{job_id} — poll while queued/running. */
export async function getJobStatus(jobId: string): Promise<JobStatus> {
  const res = await engineFetch(`/api/status/${encodeURIComponent(jobId)}`, {
    method: "GET",
    cache: "no-store",
  });
  const data = await readJson(res);

  if (!res.ok) {
    throw new Error(serverError(data, `Status check failed (HTTP ${res.status}).`));
  }
  if (!data || typeof data.status !== "string") {
    throw new Error("The engine returned an unexpected status payload.");
  }
  return data as unknown as JobStatus;
}

function filenameFromDisposition(header: string | null): string | null {
  if (!header) return null;
  // RFC 5987 form: filename*=UTF-8''GGC%20UW.xlsx
  const star = /filename\*\s*=\s*(?:UTF-8|utf-8)''([^;]+)/.exec(header);
  if (star?.[1]) {
    try {
      return decodeURIComponent(star[1].trim().replace(/^"|"$/g, ""));
    } catch {
      // fall through to the plain form
    }
  }
  // Plain form: filename="GGC UW.xlsx" or filename=GGC_UW.xlsx
  const plain = /filename\s*=\s*("?)([^";]+)\1/i.exec(header);
  if (plain?.[2]) return plain[2].trim();
  return null;
}

/**
 * GET /api/download/{job_id} — fetch as blob → object URL → programmatic <a download> click.
 * `downloadUrl` is the engine-relative path from the job result.
 */
export async function downloadExcel(downloadUrl: string): Promise<void> {
  const res = await engineFetch(downloadUrl, { method: "GET" });

  if (!res.ok) {
    const data = await readJson(res);
    throw new Error(serverError(data, `Download failed (HTTP ${res.status}).`));
  }

  const blob = await res.blob();
  const filename =
    filenameFromDisposition(res.headers.get("Content-Disposition")) ?? "GGC_Underwriting.xlsx";

  const objectUrl = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = objectUrl;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 10_000);
}
