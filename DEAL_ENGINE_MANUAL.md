# The GGC Deal Engine — Technical Manual

*An AI-driven underwriting platform for manufactured-housing communities and RV parks*

Author: Nicholas Revenco
Built for: Gary Group Capital (GGC)
Document version: June 2026

---

## Table of contents

1. Executive summary — what this thing is and why it exists
2. The product at a glance — concrete numbers
3. System architecture — the full stack
4. The user journey — what happens between upload and download
5. Inputs — exactly what the engine ingests and from where
6. Stage 0: Document parsing — turning PDFs into structured text
7. Stage 1: Extraction — faithful transcription, no judgment
8. Stage 2: Verification — pure-Python tie-out checks
9. Stage 3: Methodology — GGC's underwriting logic, applied
10. Stage 3.5: Deterministic Python overrides — the rules that cannot drift
11. Stage 4: Market research — comps, demographics, demand signal
12. Stage 5: Write-back — populating the 16-tab Excel model
13. Cell-by-cell map: exactly where every value lands
14. What is fully automated vs. what a human still touches
15. Self-consistency, schema enforcement, and the accuracy engine
16. Caching and re-run determinism
17. The frontend — Next.js, Firebase Auth, live job dashboard
18. Hosting — Cloud Run, Firebase, Vercel
19. Security model — auth, rules, secrets
20. Testing, observability, and operational hygiene
21. What is unique about this build
22. Appendix A — canonical category strings
23. Appendix B — the file inventory

---

## 1. Executive summary — what this thing is and why it exists

The GGC Deal Engine is an end-to-end AI underwriting system for a private equity firm that acquires manufactured-housing communities (MHCs) and RV parks. It takes a seller's raw financial package — typically a trailing-twelve-month operating statement, a rent roll, a P&L, and a broker offering memorandum, in arbitrary formats — and produces GGC's official 16-tab Excel underwriting model, fully populated, with every formula preserved, in roughly five to fifteen minutes.

The starting condition is a real-world mess. Seller financials arrive as scanned PDFs from one of a dozen property-management systems (Yardi, RealPage, AppFolio, Rent Manager, MRI, QuickBooks exports, and many more), in layouts that change deal-to-deal, with dirty account labels, inconsistent column headers, multiple reporting periods stacked next to each other, and rent rolls that frequently omit vacant lots. The ending condition is a 1.3-MB Excel workbook with thirteen sheets, ~9,500 pre-wired formulas, and a fully populated underwriting view that ties out to the source documents.

The engine does not replace the analyst. It produces a *first-draft screening model* that a human reviews, ties out, and finalizes. The bar is "useful, consistent, auditable" — every dollar in the output is traceable back to the source document it came from, and any value the engine could not verify is flagged on a dedicated Extraction Check tab before the workbook is shipped.

This document explains, in detail, how every part of that pipeline works.

---

## 2. The product at a glance — concrete numbers

| Component | Detail |
|---|---|
| Backend engine | Python (Flask + gunicorn), one file: `backend.py`, **6,794 lines** |
| Extraction prompt | **96 lines, 7,054 characters, ~1,800 tokens** |
| Methodology prompt | **526 lines, 39,024 characters, ~9,800 tokens** (`FINANCIAL_PARSE_PROMPT`) |
| Market-research prompt | **112 lines, 5,265 characters, ~1,300 tokens** |
| Combined prompt corpus | ~52 KB / ~13,000 tokens of static instructions per run |
| Output JSON schemas | Extraction = 4,426 chars; Methodology = 7,368 chars (JSON Schema enforced by Anthropic Structured Outputs) |
| GGC chart-of-accounts mapping | **13 canonical income categories, 23 canonical expense categories**, enforced as JSON-schema enums on every LLM call |
| Deterministic verification functions | `verify_extraction` (296 lines), `verify_methodology` (287 lines), `apply_ggc_overrides` (368 lines), `_check_template_wiring` (also pure-Python) |
| Tie-out checks emitted per run | 36 distinct `checks.append(...)` sites, expanding to ~30-90 checks per workbook depending on document complexity |
| Self-consistency runs per stage | Default N=3 (extraction) + N=3 (methodology); deep_search mode is N=5 + N=5 |
| 16-tab template | **13 worksheets, 9,559 pre-wired formulas, 2,000-row rent-roll capacity** |
| Largest sheet | Rent Roll Input: 2,002 rows × 131 cols, 4,175 formulas |
| Cell-write paths | `fill_template` writes 478 lines of cell logic across 6 worksheets; the Extraction Check, Comps Analysis, and Miscellaneous tabs are generated procedurally |
| Frontend (legacy) | `index.html` — single-page Tailwind, no framework, 742 lines |
| Frontend (hosted) | Next.js 15 + React 19 + TypeScript + Firebase Auth, **1,654 lines** across 10 files in `web/` |
| Backend lines (Python + helpers) | `backend.py` 6,794 + `fix_template.py` 1,639 + `build_template.py` 352 + tests 481 = **~9,200 lines** |
| Hosting | Cloud Run (engine) + Vercel (web) + Firebase Auth / Firestore / Storage |
| Average run cost | ~$0.50 – $2.50 per deal in Claude tokens, $0.05 in Google APIs |
| Tests | 11 pytest tests in `tests/test_pipeline.py` + `tests/test_template_contract.py` covering total-units exactness, NOI tie-out, loan scenario, sources & uses, waterfall, and category-string enum reachability |
| External APIs | Anthropic Claude (Opus 4.8 across all three stages, soon Fable 5) · Anthropic web_search_20260209 (market) · Google Document AI Layout Parser · Google Static Maps · Google Street View · Google Geocoding · Firebase Admin SDK |

---

## 3. System architecture — the full stack

The product is three independently-deployable surfaces tied together by a stateless HTTPS contract:

```
                ┌──────────────────────────────────────────┐
                │  Browser (any modern Chromium / Safari)  │
                └──────────────────┬───────────────────────┘
                                   │  Firebase ID token (JWT)
                                   ▼
        ┌───────────────────────────────────────────────────────┐
        │  Vercel — Next.js 15 + React 19 + Tailwind            │
        │  - AuthGate (Firebase Google sign-in)                 │
        │  - DealForm (multipart upload, localStorage)          │
        │  - JobProgress (4-second polling of engine status)    │
        │  - ResultsPanel (KPI cards + Excel download)          │
        └──────────────────┬────────────────────────────────────┘
                           │  POST /api/analyze (multipart)
                           │  GET  /api/status/{job_id}
                           │  GET  /api/download/{job_id}
                           ▼
        ┌───────────────────────────────────────────────────────┐
        │  Google Cloud Run — single-instance gunicorn          │
        │  ┌─────────────────────────────────────────────────┐  │
        │  │ backend.py (6,794 lines)                        │  │
        │  │  • Anthropic SDK (Claude Opus 4.8)              │  │
        │  │  • Google Document AI Layout Parser             │  │
        │  │  • openpyxl + Pillow (Excel writer + imagery)   │  │
        │  │  • Pydantic v2 + JSON Schema validators         │  │
        │  │  • Threaded job runner with cancellation        │  │
        │  └─────────────────────────────────────────────────┘  │
        └──┬───────────────────────┬───────────────────────┬────┘
           │                       │                       │
           ▼                       ▼                       ▼
   ┌──────────────┐       ┌──────────────────┐    ┌────────────────┐
   │ Anthropic    │       │ Google Document  │    │ Firebase       │
   │ Claude API   │       │ AI Layout Parser │    │ Admin SDK      │
   │ + web_search │       │ (PDF → markdown) │    │ (Firestore +   │
   └──────────────┘       └──────────────────┘    │  Storage)      │
                                                  └────────────────┘
```

The engine and the web app talk over plain HTTP with three endpoints. There is no message bus, no queue, no Redis. Job state lives in the engine process's memory; durable copies of finished workbooks land in Firebase Storage so a Cloud Run restart loses the in-flight job log but not the user's completed deals.

The reasoning behind this choice is deliberate. A single-process design makes it impossible to ship a partially-correct workbook to one user while another user races into the same code path. It also makes debugging trivial — every job's full trace lives in one log stream — and the engine has no horizontal-scaling requirement for the firm's current volume (sub-thousand deals/year).

Files in the engine directory that you should know:

```
backend.py                       6,794 lines — the engine
fix_template.py                  1,639 lines — one-time template surgery
build_template.py                  352 lines — regenerates 2,000-row rent roll
GGC_Blank_Underwriting_Sizer_Extended.xlsx  — the 13-sheet template
index.html                         742 lines — legacy single-page UI (debugging)
Dockerfile                                  — Cloud Run image (single worker)
firebase.json + firestore.rules + storage.rules
DEPLOYMENT.md                              — full runbook
CLAUDE.md                                  — architectural north star
web/                            1,654 lines — Next.js hosted UI
tests/                            481 lines — pytest suite
```

---

## 4. The user journey — what happens between upload and download

A complete run, from the user's point of view:

1. **Sign in** — the Vercel app shows a Google sign-in screen. Firebase Auth verifies the Google identity, and on success the page reveals the new-deal form. Access is restricted to an explicit `ALLOWED_EMAILS` allowlist enforced by the Cloud Run backend, so users outside the firm cannot run deals even if they sign in.

2. **Fill the form** — name, address, city, state, county, asking price, unit count (optional — the engine can derive it from the rent roll), POH count, flood zone yes/no, optional per-site tax assumption (typical range $100-$2,000/site/year), and optional cost-mode toggles. Form values persist in `localStorage` so a reload doesn't lose them.

3. **Drop the documents** — the drag-and-drop zone accepts PDFs, Excel, CSV, images, and plain text. Up to 50 MB per file, 150 MB across all uploads. The engine whitelists file extensions on the server side before reading bytes, rejecting anything unexpected with a 400.

4. **Submit** — the browser fires a `multipart/form-data` POST to `/api/analyze`. The engine validates everything, opens a new in-memory job under a 22-character cryptographically-random ID, spawns a daemon thread to run the analysis, and returns the job ID immediately.

5. **Live progress** — the page polls `/api/status/{job_id}` every 4 seconds and renders the engine's `progress` field ("Extracting financials…", "Verifying tie-outs…", "Applying methodology…", "Filling workbook…"). A cancel button is wired to `/api/cancel/{job_id}`, which sets a flag the analysis thread checks between Claude calls and raises `CancelledError` on the next stage boundary.

6. **Complete** — when the status flips to `complete`, the panel renders KPI cards (purchase price, units, occupancy, NOI, ingoing cap rate, stabilized yield on cost) and a Download button that streams the Excel from `/api/download/{job_id}` as an attachment.

7. **Needs review** — if any hard verification check failed and the workbook would mislead the reviewer, the status flips to `needs_review` instead of `complete`. The panel shows the failed check names and a "fix-and-retry" message. No workbook is shipped.

Total wall-clock: typically 4-12 minutes. The market-research call (Claude with web_search) is the long tail.

---

## 5. Inputs — exactly what the engine ingests and from where

The engine consumes two sources of information per deal: the **uploaded documents** and the **form fields**. Everything else is derived.

### 5.1 Form fields (multipart POST, exact names)

| Field | Required | Purpose |
|---|---|---|
| `property_name` | Yes | Used in the workbook filename and on Underwriting!N4 |
| `address` | Yes | Geocoded to a full address; written to Underwriting!N5; used for Street View/satellite imagery and market research |
| `city`, `state`, `county` | Yes | Required by the market-research prompt and the Underwriting!N10 county cell |
| `units` | Optional | If left blank, derived from the count of rent-roll rows the extraction step finds |
| `poh_count` | Optional | Park-owned-home count; reconciled against rent-roll home-rent rows |
| `asking_price` | Yes | Written to Underwriting!P9 → drives P4 (Purchase Price) → drives cap rate, debt sizing, returns |
| `flood_zone` | unknown / yes / no | When "yes", `apply_ggc_overrides` multiplies insurance by an additional 1.15× |
| `county_tax_rate` | Optional | Carried into the methodology prompt for tax-reassessment context |
| `tax_per_site` | Optional | $/unit/year; written to Underwriting!J22; template formula `I22 = J22 × N7` then produces underwritten RE Taxes |
| `cost_mode` | economy / balanced / max | Down-shifts which model handles which stage |
| `n_runs` | 1 / 3 / 5 | Self-consistency runs per LLM stage |
| `skip_market` | Boolean | Drops the market-research call (fastest cost saving) |
| `deep_search` | on / off | Bumps n_runs to 5 across all stages and enables broader market research |
| `files[]` | Yes (1-N) | The seller documents |

### 5.2 Allowed file types

`pdf`, `xlsx`, `xls`, `csv`, `png`, `jpg`, `jpeg`, `txt`, `md` (`ALLOWED_UPLOAD_EXTS` in `backend.py`).

### 5.3 Where inputs are read

- **Browser:** `web/components/DealForm.tsx` (538 lines) collects form fields and files, persists the persistable subset to `localStorage`, and submits via `web/lib/engine.ts`'s typed `startAnalysis(fields, files)` helper.
- **Engine:** `@app.route("/api/analyze")` in `backend.py:6619` validates the multipart payload, builds a `property_info` dict, encodes each file via `encode_file_for_claude(...)`, generates a job ID, and hands off to a background `Thread`.

There is one hard validation cliff: `city` and `state` must be present, because the market-research prompt cannot disambiguate "Las Brisas" from the dozens of properties with similar names without a geographic anchor. Everything else fails gracefully.

---

## 6. Stage 0: Document parsing — turning PDFs into structured text

Before any LLM sees a document, the engine routes it through a parser stack. The goal is to deliver dense-table fidelity high enough that the extraction model does not have to guess at column boundaries on a scanned T-12.

### 6.1 Routing

`encode_file_for_claude` in `backend.py:1603` inspects the file extension and dispatches:

- **Excel (`.xlsx`, `.xls`) and CSV** — parsed directly with `openpyxl` and a custom merged-range flattening pass (`backend.py:1668`). Highest fidelity. Multi-sheet workbooks become multiple text blocks. Merged headers get their value mapped to every cell in the merge range, eliminating the "top-row label only appears once" trap that crashes naive parsers.
- **PDFs** — handed to a parser backend selected by the `PARSER_BACKEND` env var: `docai` (Google Document AI Layout Parser, the default), `azure` (Azure Document Intelligence Layout), `tensorlake`, or `reducto`. Each returns markdown-shaped text that preserves table structure. The fallback when the parser is disabled or errors is to send the raw PDF base64 directly to Claude, which is lower-fidelity but never blocks the run.
- **Images** — base64 encoded with the appropriate media type, sent as image content blocks.
- **Text/markdown** — passed through verbatim.

### 6.2 The Document AI path (the default)

`_parse_via_docai(...)` (`backend.py:831`) sends the PDF bytes to Google's Layout Parser processor, which returns a structured response with blocks, table cells, and reading order. The engine re-flows that into a markdown document where each table is a properly-aligned grid. This is the layer that lets Claude read a scanned T-12 from a 2007 Yardi printout and still pick the right column.

### 6.3 Content-hash caching

Every parse result is cached on disk by SHA-256 of the file bytes plus the parser version (`PARSER_VERSION = "v1"`). A re-upload of the same PDF — or a re-run on the same deal — hits the cache and returns the parsed text in milliseconds with zero parser cost.

### 6.4 Multi-document handling

The engine accepts an arbitrary number of files in a single submission. Each file is parsed independently and arrives at the extraction stage as its own labeled content block. The extraction prompt instructs the model to classify each document (T-12, rent roll, P&L, OM) and only pull numbers from the appropriate one.

---

## 7. Stage 1: Extraction — faithful transcription, no judgment

This is the first LLM call. Its only job is to read the parsed documents and emit a clean, structured JSON record of what is actually there. No categorization, no methodology, no interpretation.

### 7.1 The model

- **Model:** `claude-opus-4-8` (configurable via `MODEL_EXTRACTION` env var; the codebase is staged for Fable 5 once org-level access lands).
- **Thinking:** off. Extraction is a transcription task — adaptive thinking would burn budget without improving fidelity.
- **Temperature:** 0 when the model accepts it (Opus does; Fable 5 does not — `_accepts_sampling(model_id)` in `backend.py:1246` reattaches it conditionally).
- **Output enforcement:** Anthropic Structured Outputs (GA, header `structured-outputs-2025-11-13`) with a `EXTRACTION_OUTPUT_SCHEMA` (4,426 characters of JSON Schema). The grammar mask makes it impossible for the model to emit malformed output. If the schema fails to compile and the API falls back to prompt-only enforcement, `_SCHEMA_FALLBACK_COUNT` ticks and the workbook surfaces `schema-degraded` on the Extraction Check tab.

### 7.2 The prompt (the "extraction engine" prompt)

96 lines and 7,054 characters. The opening line establishes the constraint:

> *"Your ONLY job is to faithfully transcribe the numbers from the attached seller documents into clean structured JSON. You do NOT categorize, analyze, or apply any methodology. You transcribe exactly what is in the documents."*

The prompt is organized into three steps:

1. **Identify the correct reporting period.** Seller statements frequently include multiple period columns side by side ("T-12 Ended 5/23", "T-12 Ended 9/22", "Oct 2022 - May 2023"). The prompt instructs the model to enumerate every period it sees, pick the most recent complete trailing-twelve-month column, and reject any column covering fewer than 12 months as a partial period.
2. **Extract the income statement.** For every line item, transcribe seller label (with the GL account number preserved verbatim), annual total, the 12 monthly values, and any pro-forma column and notes adjacent to it. Subtotals are flagged with `isSubtotal: true` rather than dropped.
3. **Extract the rent roll, twice.** Once as aggregated unit-type summaries (counts, average rents) and once as per-tenant rows. The per-row form is mandatory — the downstream Unit Mix Summary tab counts each row, and missing per-row data zeros out every per-unit metric in the workbook.

### 7.3 The output schema

`EXTRACTION_OUTPUT_SCHEMA` mirrors a Pydantic class (`ExtractedFinancials`) with these top-level fields:

```python
class ExtractedFinancials(BaseModel):
    reportingPeriod: ExtractedReportingPeriod
    income:          list[ExtractedLineItem]
    expenses:        list[ExtractedLineItem]
    rentRoll:        ExtractedRentRoll
    documentsSeen:   list[str]
    extractionNotes: str
```

Every line item carries `sellerLabel`, `annualTotal`, `monthly[12]`, `proFormaTotal`, and `sellerNotes`. The notes field is critical because it drives Omitt-routing downstream (the seller's own "Non-recurring" tag forces the line out of EGI/OpEx).

### 7.4 Self-consistency

For accuracy-critical runs (`deep_search=on` or `n_runs=3+`), `call_extract_financials_merged` (`backend.py:3070`) runs the same prompt N times in parallel and reconciles the runs in `_merge_extraction` (`backend.py:3162`). Numeric fields use a median; categorical fields use a confidence-weighted mode. A non-unanimous vote on a field is itself a signal — it gets recorded on the Extraction Check tab so the reviewer can see that the model disagreed with itself.

This pattern exists because `temperature=0` does **not** make hosted LLMs deterministic. GPU floating-point non-associativity and MoE expert-routing drift produce real run-to-run variance even on identical inputs. Self-consistency voting is the answer, not lower temperature.

---

## 8. Stage 2: Verification — pure-Python tie-out checks

After extraction, before the methodology stage, the engine runs a battery of deterministic, AI-free checks. This is `verify_extraction` (296 lines, `backend.py:2394`) and `verify_methodology` (287 lines, `backend.py:2690`). Together they emit between 30 and 90 checks per run, each one a dict `{item, check, status, detail}` with status `ok`, `warn`, or `fail`.

### 8.1 What gets checked

- **Reporting period sanity** — `monthsCovered == 12`? Were there multiple candidate periods? Was the right one picked?
- **Monthly-to-annual reconciliation per line** — for every income and expense line that has both a monthly array and an annual total, does the sum of the 12 monthlies tie to the annual within 1% (ok), 5% (warn), or above 5% (fail)?
- **Section subtotals** — do the leaf line items inside each section sum to the section subtotal?
- **EGI − OpEx = NOI** — does the implied NOI from the extracted lines match the methodology's emitted NOI within rounding?
- **Period covers 12 consecutive months** — no gaps, no overlaps, no partial period treated as annual.
- **Rent roll rows vs. stated unit count** — `rows == units` is the happy path; `rows < units` is the common "seller omitted vacant lots" case (flagged warn, downstream methodology imputes market rent for the gap); `rows > units` is the "form had wrong unit count" case (flagged warn).
- **Rent roll per-row data emitted** — if the extraction returned aggregated unit types but no per-tenant rows, the Rent Roll Input tab will be synthesized from group averages with uniform rents and placeholder names. Flagged warn so the reviewer knows the displayed rent roll is synthesized rather than transcribed.
- **POH count vs. rent-roll home-rent entries** — user said 30 POH but only 10 rent-roll rows show home rent? Possible employee allowance (manager comp) or hidden home-rent income.
- **Unit IDs unique** — duplicates surface as warn.
- **Occupied × avg lot rent ≈ scheduled rent** — within 5%.
- **Template SUMIFS wiring** — `_check_template_wiring` (`backend.py:2977`) verifies that every emitted `ggcCategory` exists in the canonical enum and that the implied I47 (NOI) SUMIFS sum would match the methodology's NOI.

### 8.2 What happens on failure

Each check lands on the **Extraction Check tab** (sheet index 0 — the first tab the reviewer sees), color-coded OK / WARN / FAIL with the detail string explaining what was off. A summary banner at the top of the tab says, for example, "12 OK / 3 warn / 0 fail — review warnings before trusting."

For hard fails, the engine takes one of two routes:

1. **Self-correction retry** — `MAX_PARSE_RETRIES = 2` (env-tunable). The validator output is fed back into the methodology prompt and the model re-runs. Pure-Python validation catches LLM transcription errors at the boundary where they can still be fixed cheaply.
2. **Block the write-back** — if hard fails persist after retries, the job ends in status `needs_review` and the workbook is **not** produced. The web app shows the failed check names and a fix-and-retry message.

The Extraction Check tab is the single most important interface decision in the product. The alternative — quietly emit a workbook and let the reviewer find the mismatch by eye — was the design before this tab existed and was responsible for nearly every accuracy incident.

---

## 9. Stage 3: Methodology — GGC's underwriting logic, applied

This is the second LLM call. It receives the clean, verified extraction output and applies GGC's specific underwriting methodology to it. It does not re-read the source documents.

### 9.1 The model

- **Model:** `claude-opus-4-8` (same env-override + Fable-5-readiness as extraction).
- **Thinking:** **on**, adaptive, effort `high` (configurable via `THINKING_EFFORT`). This stage involves judgment — categorizing 50 seller line items into GGC's 36-bucket chart of accounts, deciding which historical value to use for bad-debt, identifying one-time spikes — and the thinking budget materially improves the output.
- **Max tokens:** `MAX_TOKENS_METHODOLOGY = 96,000` (~80K thinking headroom + ~16K for the JSON output).
- **Output enforcement:** Anthropic Structured Outputs with `METHODOLOGY_OUTPUT_SCHEMA` (7,368 characters of JSON Schema). The schema declares `ggcCategory` as a hard `enum` of the 13 canonical income strings (for income lines) or 23 canonical expense strings (for expense lines). The grammar mask makes it physically impossible for the model to emit a category that isn't in GGC's chart of accounts.

### 9.2 The prompt (the "underwriting analyst" prompt)

`FINANCIAL_PARSE_PROMPT` is **526 lines, 39,024 characters, ~9,800 tokens** — and it is the heart of the system. It opens:

> *"You are a real estate underwriting analyst at Gary Group Capital (GGC), a private equity firm focused on mobile home parks… Your job is NOT to re-read raw documents — it is to apply GGC's categorization and underwriting methodology to the clean data you are given."*

The prompt is structured as a series of explicit, enumerated rules:

- **GGC chart of accounts** — the 13 income categories and 23 expense categories printed verbatim from `GGC_INCOME_CATEGORIES` and `GGC_EXPENSE_CATEGORIES`, with a hard rule: "schema validation rejects deviations."
- **GL-account-prefix mapping table** — a markdown table mapping every common seller GL prefix to the GGC bucket. `5701-5710 → Payroll` with explicit instruction to emit one row per leaf GL and never the rolled-up "5700 Total Personnel" subtotal. `5051 Car Insurance → Omitt Expense` regardless of GL prefix because vehicles are not property opex. `4108 Storage → Parking Income`. The table reflects ten-plus deals' worth of pattern-matching and is the most expensive piece of institutional knowledge in the file.
- **Decision rules** for ambiguous cases — Utility Reimbursement vs Other Income, negative-amount handling, "Less:" prefix prohibition (the template's SUMIFS keys on bare "Bad Debt" and would silently zero out a "Less: Bad Debt" line), one-time spike detection.
- **The four-step collections build** — GPR → physical vacancy → concessions → bad-debt-as-plug. With explicit guidance on which annualization to use (T3 if trending up, T12 if flat, T3 + flag if trending down, T6 if a one-time distorts T3).
- **Income spike detection** — any month ≥1.5× the average of the other 11 months in the same line is flagged and excluded from T3/T6 annualization while preserved in T12.
- **Broker pro-forma extraction** — every line must carry the broker's pro-forma column value when present, with an explicit rule that null means "no pro-forma column existed" and never "lazy extraction."
- **GPR multi-column disambiguation** — explicit guidance on the "rent roll total is monthly, multiply by 12" trap and a sanity check (Expected GPR ≈ Units × Avg Lot Rent × 12) with a hard floor of "if your GPR is <50% of this, you grabbed the wrong column."
- **Taxes** — primary method is `(Purchase Price × 0.65) × Local Tax Rate`, fallback is `Historical T12 × 1.15`, with a sanity check that the primary method can never produce a number lower than historical.
- **Insurance** — `T12 × 1.05` by default, `× 1.15` for flood-zone properties.
- **Management fee** — override the seller's number entirely. `5% of EGI` under 200 sites, `4% of EGI` at 200+. Explicit instruction not to reassign payroll into the management-fee bucket; both lines coexist.
- **POH bifurcation** — Lot Rent NOI and Home Rent NOI capitalized at different cap rates because GGC is in the land business, not the home-rental business.
- **Stabilized column and yield on cost** — `Stabilized Yield on Cost = Stabilized NOI / (Purchase Price + CapEx)`; the spread vs. ingoing cap rate must be ≥200 bps for the deal to clear.
- **Per-unit benchmarks** for every expense line so the model can flag any line that looks materially off.
- **Output schema** — every line item carries `ggcCategory`, `sellerName`, `fyPrior`, `fyCurrent`, `brokerProforma`, `t12Total`, `monthly[12]`, `ggcUnderwritten`, `confidence` (high/medium/low), and `notes`.

### 9.3 The Pydantic mirror

The methodology output is parsed through `MethodologyIncomeItem` / `MethodologyExpenseItem` Pydantic v2 models with three layers of validation:

1. `ggcCategory` enum check against the canonical lists.
2. `_none_to_zero` pre-validator on every numeric field — the model frequently emits `null` for `fyPrior` when only 2024 + T12 are in the source, and treating null as 0 keeps downstream math clean.
3. `monthly_ties_to_total` post-validator — if a line's `monthly[12]` sums to more than 5% off the `t12Total`, the parse fails and the error is fed back to the model on retry.

### 9.4 Self-consistency on methodology

Like extraction, methodology runs N times in parallel (default N=3, deep_search N=5). `_merge_methodology` (`backend.py:4044`) reconciles by category: for each `(ggcCategory, sellerName)` pair across runs, the numeric fields are field-level merged via confidence-weighted median, and categorical disagreements are escalated to a check.

---

## 10. Stage 3.5: Deterministic Python overrides — the rules that cannot drift

This is the most under-appreciated and most accuracy-critical piece of the engine. After the methodology stage produces its merged JSON, `apply_ggc_overrides` (368 lines, `backend.py:4261`) walks the data and **forcibly replaces** the LLM's output on every GGC rule that is a rule, not a judgment call. Run-to-run drift on these lines is not allowed.

What it forces:

- **Category normalization** — any LLM variant ("General and Administrative" → "G&A", "Less: Bad Debt" → "Bad Debt", trailing whitespace, capitalization drift) is normalized to the canonical enum value the Underwriting tab SUMIFS expect.
- **Omitt routing from seller notes** — any line whose `sellerNotes` field contains "Non-recurring", "Discontinued", "Seller Specific", "One-time", or "Non-operating" is forced to `Omitt Income` or `Omitt Expense` regardless of what the LLM picked. The Omitt bucket is preserved in Data Consolidation but explicitly excluded from EGI / OpEx / NOI by the template's SUMIFS.
- **Vehicle-related expenses → Omitt Expense** — Car Insurance, Vehicle Fuel, Vehicle Maintenance, regardless of GL prefix.
- **Bad-debt sign** — always negative on `t12Total`, `ggcUnderwritten`, and the `monthly[12]` array. The Underwriting tab's NRI formula adds bad debt (so a negative value reduces NRI); a stray positive sign would inflate revenue.
- **Management fee** — recomputed from scratch: `5% × EGI` under 200 sites, `4% × EGI` at 200+. Any existing seller mgmt-fee line is overwritten; if missing, a synthetic line is inserted.
- **Insurance** — `T12 × 1.05`, or `T12 × 1.05 × 1.15` when the user marked the property flood zone on the form.
- **Taxes** — never below the historical `T12 × 1.15` floor; the user's per-site assumption (if provided) writes to Underwriting!J22 and the template formula `I22 = J22 × N7` computes the final number.
- **CapEx reserve** — $75/unit/year, written to Underwriting!J43, formula `I43 = J43 × N7`.
- **Subtotal collapse cleanup** — drops placeholder "5700 Total Personnel" rows that survived the methodology prompt's prohibition.

Every override is recorded on `financials["_ggcOverrides"]` and surfaced on the Extraction Check tab so the reviewer can see exactly what changed and why.

The reason this layer exists is that LLMs are non-deterministic, and on a rule like "bad debt must be negative" or "mgmt fee is 5% of EGI," any deviation between runs is unacceptable. A pure-Python override layer is the only way to guarantee zero drift on the rules.

---

## 11. Stage 4: Market research — comps, demographics, demand signal

In parallel with the financial pipeline (extraction → verification → methodology → overrides), a third LLM call runs market research. This is independent of the seller financials and can be skipped via the `skip_market` form toggle.

### 11.1 The model and tools

- **Model:** `claude-opus-4-8`.
- **Thinking:** on, effort `high`.
- **Tool:** `web_search_20260209` (Anthropic's built-in web-search tool) with 8 searches available. The model is told explicitly to "use web_search aggressively — better data means a better deal decision."

### 11.2 The prompt

`MARKET_RESEARCH_PROMPT` is 112 lines, 5,265 characters. It tells the model to return:

1. **Rent comps** — 12-20 neighboring MHPs within ~25 miles, each with name, address, distance, units, lot rent, occupancy, year built, POH %, amenities, quality rating, source URL.
2. **Sale comps** — 8-15 recent transactions in the broader region with sale date, units, sale price, $/unit, cap rate, NOI, buyer/seller, source URL.
3. **Demographics (rich)** — county and MSA level: population, 1/5/10-year growth, median HH income, unemployment, poverty rate, % HH under $50K, top 10 named employers with counts, major industries, planned developments.
4. **Alternative housing** — avg single-family home price + 1-yr appreciation, avg 1BR/2BR/3BR apartment rents + growth, apartment vacancy, construction permits.
5. **MHP affordability calculation** — MHP all-in monthly cost vs. 2BR apartment rent, expressed as savings %.
6. **Landmarks** — distance to 3 named nearest major employers, Walmart, grocery, hospital, schools, college, highway, downtown, airport.
7. **Market conclusions** — `marketRentConclusion` (3-4 sentences positioning subject vs. comp range), `marketCapRateConclusion`, `demandSignal` (STRONG/MODERATE/WEAK), `demandRationale` (4-6 sentences citing specific numbers).

### 11.3 Imagery

Independently of the LLM, the engine fetches:

- **Satellite view** via Google Static Maps API (`fetch_google_static_map`).
- **Street view entrance** via Google Street View API with `fetch_google_streetview` and `_fetch_streetview_by_address` — including a heading-correction pass that uses the geocoded location to point the camera at the property rather than down the street.

These images are embedded into the Miscellaneous tab of the workbook via openpyxl + Pillow, with the image bytes content-hash cached so a re-run on the same address is free.

---

## 12. Stage 5: Write-back — populating the 16-tab Excel model

This is `fill_template` (478 lines, `backend.py:5168`) plus three procedurally-generated tabs (`add_extraction_check_tab`, `add_comps_analysis_tab`, `add_miscellaneous_tab`). It is where the engine produces the deliverable.

### 12.1 The template

`GGC_Blank_Underwriting_Sizer_Extended.xlsx` is GGC's official 13-sheet underwriting model with **9,559 pre-wired formulas**, extended from the firm's standard template to support 2,000 rent-roll rows. The engine never regenerates this file from scratch — it loads it with openpyxl, writes values into specific cells, and saves a copy. All formatting, column widths, merged headers, hyperlinks, embedded styling, and every formula are preserved exactly.

Sheet inventory and formula counts:

| Sheet | Dimensions | Formulas |
|---|---|---|
| Comps | 11 × 8 | 0 (replaced by procedural Comps Analysis tab) |
| Data Consolidation | 1,157 × 66 | 347 |
| Rent Roll Input | 2,002 × 131 | **4,175** |
| Collections | 40 × 112 | 163 |
| Unit Mix Rent Growth | 49 × 109 | 109 |
| Unit Mix Summary | 39 × 85 | 42 |
| GGC Underwriting | 278 × 110 | 347 |
| GGC Pro Forma | 180 × 100 | 682 |
| Investor Return | 22 × 14 | 6 |
| Waterfall (10-yr) | 277 × 19 | **1,992** |
| Waterfall (5-yr) | 277 × 14 | 1,166 |
| Sources and Uses | 27 × 10 | 24 |
| Loan Scenario (acquisition) | 1,003 × 241 | 506 |
| **Total** | | **9,559** |

After the engine runs, three additional sheets are inserted at the front of the workbook:

- **Extraction Check** (sheet index 0) — verification results, color-coded OK/WARN/FAIL, with an overall summary banner.
- **Comps Analysis** (sheet index 1) — rent comps table, sale comps table, demographics block, with statistics rows (range, average) and a market conclusion.
- **Miscellaneous** (sheet index 2) — property overview with embedded satellite/street-view imagery, landmarks, demographics, top employers.

### 12.2 Formula protection

Every write to a worksheet that contains pre-wired formulas goes through `_protect_formulas(ws)` (`backend.py:5053`), which wraps the worksheet's `cell()` method so any write to a cell that already holds a formula is skipped and counted. The blocked-write tally is surfaced on the Extraction Check tab as an informational warning. This guard exists because a single accidental literal write to a formula cell (for example, overwriting `P4`'s `=IFERROR(IF(ISNUMBER(P9),P9,0),0)` with the asking price literal) silently breaks the entire downstream Sources & Uses → Loan Scenario → Pro Forma chain.

### 12.3 Atomic save

`fill_template` writes to a sibling `.tmp` file and `os.replace`'s it onto the destination only after the openpyxl save succeeds. A mid-save crash never leaves a half-written workbook in the jobs directory.

### 12.4 Force-recalc on open

`wb.calculation.fullCalcOnLoad = True` ensures Excel recalculates every formula when the user opens the file. Without this, openpyxl-written formula cells render as blank until the user presses F9.

---

## 13. Cell-by-cell map: exactly where every value lands

This is the table for "where does each piece of data go." Sheet names and cell coordinates are exact.

### 13.1 Data Consolidation (the source data that feeds every SUMIFS in the model)

| Range | Column | Meaning | How populated |
|---|---|---|---|
| A3:A36 | GGC Income Category | One of the 13 canonical income strings — drives every SUMIFS in the Underwriting tab | Written from `financials.income[].ggcCategory` after `_normalize_ggc_category` |
| A43:A102 | GGC Expense Category | One of the 23 canonical expense strings | Written from `financials.expenses[].ggcCategory` |
| B (rows 3-36, 43-102) | Seller's original label | Audit trail | Written from `sellerName` |
| D | FY Prior | Prior-year actual | Written from `fyPrior` |
| E | FY Current | Current-year actual | Written from `fyCurrent` |
| F | Broker Proforma | Broker's pro-forma column | Written from `brokerProforma` |
| G | T12 Total | Trailing-twelve-month total | Written from `t12Total` |
| H | Annualization | Pre-wired formula — not touched | — |
| J:U | Monthly (12 columns) | The 12 monthly values | Written from `monthly[]`, or T12/12 evenly distributed when monthly is missing |
| V | Row total | Pre-wired formula — not touched | — |

The engine writes only into non-structural rows. `_structural_rows(ws, 3, 36)` and `_structural_rows(ws, 43, 102)` scan once at template load and identify rows that contain pre-wired SUM/IF/NOI formulas (income rows 22-27 and expense rows 60/62/64). Writing a category label into one of those rows would make the SUMIFS pick up the row's own formula output as if it were a line item — exactly how the early "Advertising T-12 = $1.17M" bug happened.

Overflow handling is strict: if methodology emits more income lines than there are non-structural slots, the overflow is recorded as a hard FAIL on the Extraction Check tab with the dropped line names. Silent truncation past the SUMIFS range is exactly the kind of invisible accuracy degradation forbidden by the project's accuracy contract.

### 13.2 Rent Roll Input (rows 3-1002, capacity 2,000)

| Column | Meaning | Source |
|---|---|---|
| A | Count (formula) | Pre-wired — not touched |
| B | Unit ID | `rentRollRows[].unitId` |
| C | Unit Type (canonical) | `rentRollRows[].unitType`, run through `_canonicalize_unit_type` to map to one of: `TOH MH Site`, `POH-Infilled units`, `Long term RV Site`, `Retail/Commercial` |
| D | Status (Occupied/Vacant) | `rentRollRows[].status` |
| F | Tenant Name | `rentRollRows[].tenantName` |
| I | Lot Rent | `rentRollRows[].lotRent` |
| J | Home Rent | `rentRollRows[].homeRent` |
| K | Combined (formula) | Pre-wired — not touched |

The canonical unit-type mapping matters because the Unit Mix Summary tab's COUNTIFS/SUMIFS key on these exact four strings. A drift to a non-canonical string would zero out the count, which cascades to `# of Units (Underwriting!N7)` and every per-unit metric downstream.

If extraction only returned aggregated unit-type summaries (no per-tenant rows), `fill_template` synthesizes individual rows from the group counts and averages, with placeholder tenant names. This path is flagged on the Extraction Check tab so the reviewer knows the displayed rent roll is synthesized, not transcribed.

### 13.3 GGC Underwriting tab — the subject-property block

| Cell | Value | Source |
|---|---|---|
| N4 | Property name | `propertyInfo.name` (form input) |
| N5 | Full address | `propertyInfo.address`, auto-completed via Google Geocoding API to a full city/state/zip when the form input is partial |
| N6 | Property type | `propertyInfo.propertyType` (methodology output: MHC / RV / Hybrid) |
| N9 | Acreage | `propertyInfo.acreage` (methodology output from OM) |
| N10 | County | `propertyInfo.county` (form input) |
| P9 | Asking price | `propertyInfo.askingPrice` (form input). Note: P4 (Purchase Price) is wired to `=IFERROR(IF(ISNUMBER(P9),P9,0),0)`, so writing P9 drives the entire downstream Sources & Uses / Loan Scenario / Pro Forma chain. The engine never writes P4 directly — `_protect_formulas` blocks it. |
| J22 | Per-site tax assumption ($/unit/year) | `propertyInfo.taxPerSite` (form input). Template formula `I22 = J22 × N7` produces underwritten RE Taxes. A sanity-clamp warns when outside $100-$2,000/site. |
| J43 | CapEx reserve per site ($/unit/year) | $75 (forced by `apply_ggc_overrides`). Template formula `I43 = J43 × N7`. |
| N2 | Underwritten-on date stamp | `datetime.now()` at run time, formatted as `dddd, mmmm dd, yyyy` |
| R3 | Property website URL | `propertyInfo.websiteUrl` (methodology, from OM if present) |
| R4 | Year built | `propertyInfo.yearBuilt` (methodology) |
| R5 | Flood zone | `propertyInfo.floodZone` — form input takes precedence over methodology extraction |
| R6 | Utility structure | `propertyInfo.utilityStructure` (methodology) |
| R7 | Electricity notes | `propertyInfo.electricityNotes` (methodology) |
| R8 | Trash notes | `propertyInfo.trashNotes` (methodology) |
| N19 | County tax-assessor URL | `propertyInfo.taxAssessorUrl` (methodology) |
| M26:R32 | Parcel-level tax table (up to 7 parcels) | `propertyInfo.taxParcels[]` (methodology). Summary formulas at N20/N21/N22 are wired by `fix_template.py` to roll up from this block. |

### 13.4 Computed cells the engine never writes directly

Because the template carries 9,559 pre-wired formulas, the vast majority of the workbook's outputs are computed, not written. Examples:

- **EGI** at Underwriting!I19 — SUMIFS across Data Consolidation column G filtered to canonical income categories ex-Omitt.
- **OpEx** at Underwriting!I44 — SUMIFS across the canonical expense categories ex-Omitt.
- **NOI** at Underwriting!I47 — `=I19 − I44`.
- **Cap rate** at Underwriting!P6 — `=I47/P4`.
- **# of Units** at Underwriting!N7 — `COUNTA` of Rent Roll Input!C3:C2002 filtered to occupied.
- **Occupancy** at Underwriting!N8 — `COUNTIF(Rent Roll Input!D3:D2002, "Occupied") / COUNTA(...)`.
- **Loan amount** at Sources and Uses — `=Purchase Price × LTV`.
- **Pro Forma Y1-Y10** — full pro-forma chain in GGC Pro Forma tab.
- **Waterfall** — IRR / equity multiple / cash-on-cash to GP and LP at the (10-yr) and (5-yr) waterfall tabs.

Engine populates the inputs; the template's formula chain produces the outputs.

---

## 14. What is fully automated vs. what a human still touches

### 14.1 Fully automated (zero human edit needed for the engine to produce the workbook)

- Document type classification.
- Reporting-period selection from a multi-column statement.
- Line-item extraction with full monthly detail.
- Rent-roll extraction with per-tenant rows and unit-type canonicalization.
- Reconciliation tie-out checks (monthlies sum to annual, rent-roll-vs-units, period sanity).
- GGC chart-of-accounts categorization (constrained to the canonical enum at the API level).
- Deterministic GGC rule application (mgmt fee, taxes, insurance, CapEx, bad-debt sign, Omitt routing, vehicle re-bucket).
- 4-step collections build (GPR → vacancy → concessions → bad-debt-as-plug).
- Income spike detection.
- POH bifurcation flagging.
- Stabilized column computation.
- Market research (comps, demographics, demand signal) via web-search-tool LLM.
- Property imagery (satellite + street view) fetched and embedded.
- Address auto-completion via Google Geocoding.
- 13-tab workbook write-back with formula preservation, atomic save, and force-recalc.
- Three additional procedurally generated tabs (Extraction Check, Comps Analysis, Miscellaneous).
- Run history, cost tracking, and download URL.

### 14.2 What the human still owns

The engine is explicit that this is a *first-draft screening model*, not a final underwrite. The reviewer's remaining work:

1. **Open the Extraction Check tab first.** Confirm the warnings and any soft items (rent-roll-short, period ambiguity, partial-period flags). If anything looks wrong, re-upload corrected docs.
2. **Sanity-check the bad-debt line.** The engine's plug is based on the trend the LLM picked (T3 / T6 / T12); the reviewer confirms that trend is correct given the deal narrative.
3. **Confirm one-time-item flags.** Spikes are flagged but not auto-stripped; the reviewer decides whether each is genuinely one-time and adjusts the underwriting basis accordingly.
4. **Sanity-check market cap rate.** The market-research stage produces a `marketCapRateConclusion`; the reviewer picks the specific cap rate to apply against stabilized NOI for the recommended purchase price.
5. **Negotiation overrides.** P9 (Asking Price) is what the engine writes; if the deal is moving at a different number, the reviewer overwrites P9 in-cell and the entire downstream chain updates. The formula chain is preserved exactly so the model behaves as a working underwriting tool, not a static PDF.
6. **Final sign-off.** GGC's underwriting committee reviews the workbook with the reviewer's adjustments before any IC submission.

Everything the engine does is reproducible — re-uploading the same documents with the same form values produces a byte-identical workbook (see Caching, §16).

---

## 15. Self-consistency, schema enforcement, and the accuracy engine

The project has an explicit "zero acknowledged accuracy gap" contract. Five mechanisms layer together to enforce it:

1. **Hard-enum schemas at every LLM boundary.** `METHODOLOGY_OUTPUT_SCHEMA` declares `ggcCategory` as a `enum` of the canonical strings. Anthropic Structured Outputs enforces this at the grammar mask, making it impossible for the model to emit a category outside the chart of accounts. The Pydantic validator does it again belt-and-suspenders. The `apply_ggc_overrides` normalizer does it a third time.

2. **Self-consistency voting on every LLM stage.** Default N=3 (extraction) + N=3 (methodology); deep_search N=5 + N=5. Numeric fields reconcile via confidence-weighted median; categorical fields via confidence-weighted mode. Non-unanimous votes become checks on the Extraction Check tab.

3. **Deterministic Python overrides for GGC rules.** Mgmt fee, taxes, insurance, CapEx, bad-debt sign, Omitt routing — all forced in Python after methodology returns. Zero run-to-run drift on the rules.

4. **Deterministic reconciliation surfaced prominently.** 36 distinct check sites in `verify_extraction` / `verify_methodology` / `_check_template_wiring` produce 30-90 checks per workbook on the Extraction Check tab. Hard fails block the write-back; warnings surface but ship.

5. **Versioned cache.** The cache key hashes (PDF bytes + parser version + model + prompt + schema + n_runs), so a re-run with identical inputs returns byte-identical output (see §16).

The deliberate design principle running through all five: **flag and prompt, never silently assume**. A scraped/blank/zero value never silently overwrites a real one. The reviewer either sees the number or sees the warning explaining why.

---

## 16. Caching and re-run determinism

### 16.1 The extraction cache

`extraction_cache_key(...)` (`backend.py:1122`) computes a SHA-256 over:

- The parsed bytes of every uploaded file.
- The parser backend identifier (`PARSER_BACKEND`) and version (`PARSER_VERSION`).
- The form's `property_info` dict (sorted JSON).
- The number of self-consistency runs (`n_extraction_runs`, `n_methodology_runs`).
- The full text of `EXTRACTION_PROMPT` and `FINANCIAL_PARSE_PROMPT`.
- The methodology output schema.

Any change to *any* of those — a new PDF, a different parser, a prompt edit, a schema tweak — invalidates the cache. The cache itself lives on disk (Cloud Run uses `/tmp/ggc/extraction_cache`, local dev uses the project's `extraction_cache/` folder).

A cache hit short-circuits the entire LLM pipeline and returns the merged-financials JSON in milliseconds. The cache write happens only when verification produced zero hard fails — a degraded extraction is never memoized.

### 16.2 The parser cache

`parse_pdf(pdf_bytes, filename)` (`backend.py:784`) caches by `_stable_pdf_hash(pdf_bytes)` keyed against `PARSER_VERSION`, so re-uploading the same document or re-running on the same deal hits the parser cache instantly with no Document AI cost.

### 16.3 The imagery cache

Google Static Maps and Street View responses are cached by URL hash in `img_cache/`, so a re-run on the same property fetches no imagery and embeds the cached images.

The combined effect is that a clean re-run on the same deal with the same code is byte-identical and runs in seconds. This is what "no let me try again and hope for a better answer" looks like operationally.

---

## 17. The frontend — Next.js, Firebase Auth, live job dashboard

The hosted UI lives in `web/` and is a Next.js 15 application using React 19, TypeScript, and Tailwind CSS 4. It deploys to Vercel.

### 17.1 File layout

```
web/
├── app/
│   ├── layout.tsx       — root layout, fonts, AuthGate wrapper
│   ├── page.tsx         — main new-deal flow (254 lines)
│   ├── history/         — run history page
│   ├── icon.svg
│   └── globals.css
├── components/
│   ├── AuthGate.tsx     — Firebase Auth Google sign-in gate (141 lines)
│   ├── DealForm.tsx     — property form + drag-drop upload (538 lines)
│   ├── JobProgress.tsx  — live progress stepper (112 lines)
│   ├── ResultsPanel.tsx — KPI cards + download (216 lines)
│   ├── Header.tsx       — top-bar with user avatar (54 lines)
│   └── RunHistory.tsx
├── lib/
│   ├── engine.ts        — typed fetch wrappers for /api/* (257 lines)
│   └── firebase.ts      — Firebase web SDK init + token plumbing (49 lines)
├── package.json
├── next.config.ts
└── tsconfig.json
```

### 17.2 Auth flow

`AuthGate.tsx` wraps the entire app. On mount it subscribes to `onAuthStateChanged`. When the user is signed in, it renders the children inside an `AuthContext.Provider` carrying the `User` object. When the user isn't signed in, it renders a sign-in screen with a "Sign in with Google" button that calls `signInWithPopup(firebaseAuth(), googleProvider())`.

The signed-in user's Firebase ID token (a JWT) is attached as `Authorization: Bearer <token>` on every engine call by `web/lib/engine.ts`. On the backend, `@require_auth` (`backend.py:461`) verifies the token via the Firebase Admin SDK, checks the email against `ALLOWED_EMAILS`, and attaches the `uid`/`email` to Flask's `g`.

### 17.3 New-deal flow

`page.tsx` is a state machine with five phases: `idle | submitting | polling | complete | error`.

- `idle` — `DealForm` is showing.
- `submitting` — multipart POST to `/api/analyze` is in flight.
- `polling` — `setInterval(tick, 4000)` calls `/api/status/{job_id}`. The `progress` field renders as a stepper:
    - `Upload` — done as soon as the job ID is back.
    - `Parallel analysis` — active while extraction + verify + methodology + market research run.
    - `Fill template` — active when `progress.toLowerCase().includes("filling")`.
    - `Complete` — terminal.
- `complete` — `ResultsPanel` renders KPI cards (asking price, total units, occupancy %, EGI, OpEx, NOI, ingoing cap rate, stabilized yield on cost, rent-comp count, sale-comp count, demand signal, total token cost, number of Claude calls) and a Download button.
- `error` / `needs_review` — the failure detail with the list of failed verification check names so the user knows what to fix.

The active job ID persists in `localStorage` under `ggc_active_job` so a page refresh resumes polling without losing the in-flight job. The cancel button POSTs to `/api/cancel/{job_id}`, which sets a global flag the analysis thread checks between Claude calls — cancellation lands at the next stage boundary (up to ~60s overhang on the in-flight call), and the bill stops.

### 17.4 Run history

A `history/` page lists previous runs by reading the user's documents in the Firestore `deal_runs` collection (mirrored by the engine on every status change). Each row links to the durable copy of the finished workbook in Firebase Storage at `runs/{uid}/{job_id}.xlsx`. This survives Cloud Run restarts that would otherwise lose the in-memory job state.

### 17.5 Legacy UI

The original `index.html` (742 lines, hand-rolled single-page Tailwind + vanilla JS, no framework) still ships in the Docker image and is reachable at the engine's `/`. It runs in local dev with no auth. In production, with `REQUIRE_AUTH=1`, it loads but its `/api/*` calls 401 — the production UI is the Vercel app. This is intentional: it keeps a no-dependencies debugging surface available even when the Next.js app is misconfigured.

---

## 18. Hosting — Cloud Run, Firebase, Vercel

### 18.1 The engine: Google Cloud Run

The engine is containerized with a single `Dockerfile`:

```dockerfile
FROM python:3.12-slim
ENV PYTHONUNBUFFERED=1 \
    JOBS_DIR=/tmp/ggc/jobs \
    IMG_CACHE_DIR=/tmp/ggc/img_cache \
    EXTRACTION_CACHE_DIR=/tmp/ggc/extraction_cache
WORKDIR /app
COPY requirements.txt ./
RUN pip install -r requirements.txt
COPY backend.py index.html GGC_Blank_Underwriting_Sizer_Extended.xlsx ./
CMD exec gunicorn --bind :${PORT:-8080} --workers 1 --threads 16 --timeout 0 backend:app
```

Deploy command (`gcloud run deploy ggc-deal-engine --source . --max-instances 1 --no-cpu-throttling --cpu 2 --memory 2Gi --timeout 3600 --concurrency 40 ...`).

The flags are not optional:

- `--workers 1` and `--max-instances 1` are required because job state lives in process memory — a second worker or instance would see none of the first's jobs.
- `--no-cpu-throttling` keeps background analysis threads alive between status polls.
- `--timeout 0` (gunicorn) lets analyses run for minutes without being killed.
- `--cpu 2 --memory 2Gi` is enough for one concurrent run plus 15-20 polled jobs.

### 18.2 The web app: Vercel

`web/` deploys to Vercel as a standard Next.js 15 project (`vercel --prod`). Production env vars:

- `NEXT_PUBLIC_FIREBASE_API_KEY` / `_AUTH_DOMAIN` / `_PROJECT_ID` / `_STORAGE_BUCKET` / `_MESSAGING_SENDER_ID` / `_APP_ID` — Firebase web config.
- `NEXT_PUBLIC_ENGINE_URL` — the Cloud Run service URL.

### 18.3 Firebase

One Google Cloud project hosts everything: Document AI processor, Cloud Run service, Firebase Auth, Firestore, and Storage. Sharing the project means Cloud Run's default compute service account has Application Default Credentials to call Document AI, Firestore, and Storage with no key file in the container.

- **Firebase Auth** — Google sign-in, `ALLOWED_EMAILS` enforced server-side.
- **Firestore** — `deal_runs/{runId}` documents (one per job) mirror lightweight status fields so the web app's history page survives engine restarts. Deny-by-default rules; clients can only read their own.
- **Firebase Storage** — finished workbooks land at `runs/{uid}/{jobId}.xlsx`. Deny-by-default rules; clients can only download their own.

### 18.4 Cost profile

Per deal at default settings:

- Anthropic Claude tokens: ~$0.50-$2.50 (varies with document size and self-consistency runs)
- Google Document AI: ~$0.02 per PDF page
- Google Maps APIs: <$0.05
- Cloud Run compute: negligible (sub-second of CPU per second of wall-clock, scales to zero between deals)
- Firebase: free tier comfortably covers the firm's volume

Total: well under $5/deal even on a deep-search run with five concurrent self-consistency votes per stage.

---

## 19. Security model — auth, rules, secrets

### 19.1 Authentication

Every `/api/*` endpoint is decorated with `@require_auth`. When `REQUIRE_AUTH=1` is set on Cloud Run, the decorator:

1. Reads the `Authorization: Bearer <token>` header.
2. Verifies the JWT via `firebase_admin.auth.verify_id_token(token)`.
3. Checks the verified email against `ALLOWED_EMAILS` (empty → any signed-in user, populated → strict allowlist).
4. Attaches `uid` and `email` to Flask's `g` for downstream ownership checks.

### 19.2 Per-job ownership

Every job dict carries the creating user's `uid`. `/api/status/{job_id}`, `/api/cancel/{job_id}`, and `/api/download/{job_id}` all check ownership and return 404 (not 403) for jobs owned by other users — so other users' job IDs are never confirmed to exist.

### 19.3 Job-ID safety

Job IDs are generated by `secrets.token_urlsafe(16)` and validated against `JOB_ID_RE = ^[A-Za-z0-9_-]{16,64}$` before any dict or filesystem lookup. The download endpoint additionally resolves the file path and verifies it stays inside `JOBS_DIR` before serving, so a crafted `../../etc/passwd` job ID cannot probe the filesystem.

### 19.4 Firestore and Storage rules

Deny-by-default. Clients can only read `deal_runs` documents whose `uid` matches their sign-in, and only download `runs/{their-uid}/...` files. All writes go through the engine's Admin SDK (which bypasses rules), so clients have zero write access.

### 19.5 Secrets handling

- API keys (Anthropic, Google Maps) live in Cloud Run env vars set via `--set-env-vars` or Secret Manager.
- The `.env` file is gitignored and dockerignored; only used in local dev.
- `gcp-credentials.json` is gitignored and is **not** copied into the container — Cloud Run's default service account handles all Google APIs via Application Default Credentials.
- The `/api/config` endpoint never returns the server's Anthropic key over the wire — it returns only booleans indicating which integrations are configured.

### 19.6 Upload validation

- Whitelisted extensions only (`ALLOWED_UPLOAD_EXTS`).
- 50 MB per file cap, 150 MB total per request (`MAX_UPLOAD_BYTES`).
- City + state are required form fields (cannot be empty).

### 19.7 CORS

`ALLOWED_ORIGINS` env var sets the CORS allowlist on the engine. Production lists the Vercel URL plus localhost for dev. Anything else gets a CORS rejection.

---

## 20. Testing, observability, and operational hygiene

### 20.1 Tests

11 pytest tests across two files (`tests/test_pipeline.py` and `tests/test_template_contract.py`) covering:

- `test_total_units_exact` — the workbook's N7 matches the rent-roll row count.
- `test_total_noi_within_tolerance` — Underwriting!I47 matches `EGI − OpEx` from extracted inputs within rounding.
- `test_egi_within_tolerance` — Underwriting!I19 matches the methodology output.
- `test_loan_scenario_rates_nonzero` — Loan Scenario tab produces nonzero numbers (catches broken-formula regressions).
- `test_sources_and_uses_no_ref_errors` — no `#REF!` in Sources and Uses.
- `test_waterfall_no_name_errors` — no `#NAME?` in either waterfall tab.
- `test_run_analysis_job_with_mocked_claude` — end-to-end pipeline with the Anthropic API mocked.
- `test_normalize_ggc_category` — every known LLM variant maps to the canonical enum.
- `test_structural_rows_income_band` / `test_structural_rows_expense_band` — the structural-row detector finds exactly the pre-wired SUM/IF/NOI rows that must not be overwritten.
- `test_enough_slots_after_skip` — after skipping structural rows, there's enough capacity for a realistic deal.
- `test_every_enum_string_is_reachable` — every category in `GGC_INCOME_CATEGORIES` and `GGC_EXPENSE_CATEGORIES` is reachable by the methodology prompt's GL-mapping table.
- `test_end_to_end_write_preserves_structural_rows` — a full fill_template run preserves the formula rows.
- `test_correct_output_still_present` — the gold-standard `CorrectOutput.xlsx` reference is intact.

### 20.2 Observability

- Per-request logging from every Claude call (`[Claude] Stage 1/2 — EXTRACTION (claude-opus-4-8, no-thinking)...`).
- Per-run token accounting: `record_usage(model_id, usage)` accumulates input/output/cache-read/cache-write tokens to a thread-local, summed at the end of the job into `result.usage.totals.cost_usd` and `result.usage.calls`.
- Job state is fully introspectable via `/api/status/{job_id}` while running.
- Extraction-cache hits/misses log the 8-char key prefix and the deep_search/n_runs configuration.
- Verification check count logs after each verify pass.

### 20.3 Operational hygiene

- Jobs older than 50 are evicted from the in-memory `JOBS` `OrderedDict` to bound memory.
- Cancelled jobs are tracked in `CANCELLED_JOBS` and checked between Claude calls via `_check_cancelled()`.
- Atomic save on workbook write prevents shipping a half-written file.
- Formula-protection guard prevents accidental writes to formula cells.
- `_evict_old_jobs()` and `_set_job_thread()` are protected by `JOBS_LOCK` for safe concurrent access.

---

## 21. What is unique about this build

A short list of the things in this project that are unusual enough to call out in a recruiting conversation:

1. **Grammar-mask category enforcement.** GGC has 36 canonical chart-of-account strings the downstream Excel SUMIFS key on. The methodology output schema declares them as a JSON Schema `enum`, and Anthropic Structured Outputs (GA) compiles that enum into a grammar mask that makes it physically impossible for the model to emit an off-list category. A third belt-and-suspenders layer normalizes any drift in pure Python before write-back. Three layers of category protection, one rule of strictness: zero deviation reaches the workbook.

2. **A 526-line methodology prompt that encodes ten deals of institutional knowledge.** The mapping table from seller GL prefix to GGC bucket reflects pattern-matching across every property GGC has reviewed. It is the single most expensive piece of intellectual property in the file and the reason the engine exists.

3. **Self-consistency voting with confidence-weighted median.** Default N=3 runs per stage; deep_search N=5. Numeric fields reconcile by median; categorical by confidence-weighted mode. Non-unanimous votes surface as warnings. The pattern exists because `temperature=0` does not make hosted LLMs deterministic — GPU non-associativity and MoE routing drift produce real run-to-run variance, and voting is the only honest answer.

4. **Pure-Python deterministic overrides on top of the LLM output.** For rules that are rules (mgmt fee = 5% of EGI under 200 sites; bad debt is always negative; vehicles are not opex), `apply_ggc_overrides` rewrites the LLM's output in Python and records every override. Zero run-to-run drift on the rules.

5. **An "Extraction Check" tab placed at sheet index 0.** 30-90 deterministic tie-out checks per workbook, color-coded OK / WARN / FAIL, with the operative reporting period and every override change recorded. The first thing the reviewer sees when they open the file. The product principle is "flag and prompt, never silently assume" — and this tab is where that principle becomes a UI.

6. **A versioned cache keyed on the prompt + schema + parser version.** A re-run with identical inputs and unchanged code returns byte-identical output. The cache key includes the full text of the prompts and the schema, so any edit invalidates the cache safely. No "let me re-run and hope."

7. **A 9,559-formula Excel template preserved with byte-level fidelity.** openpyxl, no regeneration, formula-protection guard on every write, atomic save, force-recalc-on-open. The engine writes inputs; the template's formula chain produces outputs. The reviewer can edit any input in-cell and the entire downstream chain updates — the workbook stays a working tool, not a static report.

8. **Three procedurally-generated tabs.** Extraction Check, Comps Analysis, and Miscellaneous tabs are built fresh on every run with embedded styling, color coding, and (for Miscellaneous) embedded satellite + street view imagery pulled via Google Maps APIs and inserted as Pillow-resized PNG cells.

9. **The legacy + hosted UI duality.** `index.html` ships in the container as a zero-dependency debugging surface that works without Firebase Auth in local dev; the Next.js app in `web/` is the production UI with full auth. Both share the same engine endpoints. The legacy UI is a deliberate operational lever, not a relic.

10. **Single-instance production deployment with intentional architectural constraints.** The engine's `--workers 1 --max-instances 1` design makes it impossible to ship a partially-correct workbook to one user while another races into the same code path. Scaling out is a clearly-documented future migration to Firestore-backed job state — not a current concern because the firm does sub-thousand deals per year.

11. **Hard model pinning, no `-latest`.** The codebase pins explicit model snapshots (`claude-opus-4-8`) and is staged for Fable 5 via a single env-var flip. The `_accepts_sampling(model_id)` helper conditionally re-attaches `temperature=0` for models that accept it (Opus does; Fable 5 does not), so the switch is one line.

12. **End-to-end testing against a gold-standard reference workbook.** `CorrectOutput.xlsx` is the manually-prepared canonical output for the Whaleshead Beach Resort deal. The test suite diffs the engine's output against it for total units, NOI, EGI, loan scenario, sources and uses, and waterfall — so the regression net catches both code bugs and prompt regressions.

---

## 22. Appendix A — canonical category strings

These are the exact strings the Excel SUMIFS key on. Drift breaks the wiring.

**Income (13):** `Gross Potential Rent` · `RV Site Rental Income` · `Parking Income` · `Retail` · `Utility Reimbursement` · `Other Income` · `Bad Debt` · `Omitt Income` · `Home Rent Income` · `Employee Allowance` · `Model Units` · `Vacancy` · `Concessions`

**Expense (23):** `RE Taxes` · `Insurance` · `Water and Sewer` · `Electricity` · `Gas/Fuel` · `Trash Removal` · `Ground Maintenance` · `Repair and Maintenance` · `Recreational Amenities` · `Management Fee` · `Payroll` · `G&A` · `Professional Fees` · `Advertising` · `Home Rent Expense (MH)` · `Omitt Expense` · `Other` · `Cap-Ex Reserve` (plus 5 additional buckets for less-common cases)

**Unit types (4, canonical for Unit Mix Summary COUNTIFS):** `TOH MH Site` · `POH-Infilled units` · `Long term RV Site` · `Retail/Commercial`

Note one display-only quirk preserved by `fix_template.py`: the Underwriting tab row 26 label shows `Electrcitiy` (the typo is intentional — it is how GGC has always displayed it), but the SUMIFS criterion underneath is the correctly-spelled `Electricity`. The display typo is purely visible; the SUMIFS still matches the canonical string.

---

## 23. Appendix B — the file inventory

```
ggc-deal-engine/
├── backend.py                                    6,794 lines  — the engine
├── fix_template.py                               1,639 lines  — one-time template surgery
├── build_template.py                               352 lines  — regenerates 2,000-row rent roll
├── extract_categories.py                            38 lines  — category audit helper
├── GGC_Blank_Underwriting_Sizer_Extended.xlsx              — 13 sheets, 9,559 formulas
├── index.html                                      742 lines  — legacy single-page UI
├── Dockerfile                                                — Cloud Run image
├── requirements.txt                                          — Python deps
├── firebase.json                                             — Firebase project config
├── firestore.rules                                           — deny-by-default Firestore rules
├── firestore.indexes.json                                    — Firestore indexes
├── storage.rules                                             — deny-by-default Storage rules
├── DEPLOYMENT.md                                             — full deploy runbook
├── CLAUDE.md                                                 — architectural north star
├── README.md
├── tests/
│   ├── test_pipeline.py                            217 lines  — end-to-end pipeline tests
│   └── test_template_contract.py                   264 lines  — template-contract tests
├── extraction_cache/                                         — versioned LLM cache
├── img_cache/                                                — Google Maps imagery cache
├── jobs/                                                     — per-job .xlsx outputs
├── Outputs/                                                  — reference outputs (CorrectOutput, etc.)
└── web/                                          1,654 lines  — Next.js hosted UI
    ├── app/
    │   ├── layout.tsx                               24 lines
    │   ├── page.tsx                                254 lines
    │   └── history/
    ├── components/
    │   ├── AuthGate.tsx                            141 lines
    │   ├── DealForm.tsx                            538 lines
    │   ├── JobProgress.tsx                         112 lines
    │   ├── ResultsPanel.tsx                        216 lines
    │   ├── Header.tsx                               54 lines
    │   └── RunHistory.tsx
    ├── lib/
    │   ├── engine.ts                               257 lines
    │   └── firebase.ts                              49 lines
    ├── package.json
    └── tsconfig.json
```

Total code I wrote across the project: **roughly 11,000 lines** of Python, TypeScript, JSON Schema, and HTML, plus a 526-line, 39,000-character methodology prompt that captures the firm's underwriting logic, plus ~3,500 lines of template surgery to extend GGC's 1,000-row rent-roll model to 2,000 rows.

---

*End of manual.*
