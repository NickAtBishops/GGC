# CLAUDE.md — GGC Deal Engine

This file gives you (Claude) the context you need to work on this project effectively.
Read it before making changes.

---

## What this project is

The **GGC Deal Engine** is an AI-powered underwriting tool for **Gary Group Capital (GGC)**,
a private equity firm that acquires **manufactured housing communities (MHCs)** and **RV parks**.

The tool ingests a seller's raw financial documents — T-12 operating statements, rent rolls,
profit & loss statements, and offering memorandums (OMs), which arrive as PDF, Excel, or
scanned images in wildly inconsistent formats — and produces a populated version of GGC's
official **16-tab Excel underwriting model**.

The goal: take a deal from "a messy pile of seller PDFs" to "a populated underwriting model"
in a few minutes, so the acquisitions team can screen more deals faster. The output is a
**screening and first-draft tool**, not a final underwrite — a human still reviews and ties
out the numbers before any decision. That framing matters: the tool needs to be *useful and
consistent*, not perfect.

---

## ⚠️ THE CORE PROBLEM: accuracy and consistency

**This is the #1 issue and the thing most work should focus on.**

Current state of the output:
- Important values (NOI, GPR, total income, total expense) land within ~6% of the correct
  hand-underwritten figures.
- All values land within ~10%.
- The numbers are **not exact**, and — worse — they are **not consistent run-to-run**. The
  same document processed twice can produce different category mappings, different monthly
  values, and occasionally a wrong reporting period.

**What "accuracy" failures look like in practice:**
1. **Category drift** — the same dollar amount gets bucketed into a different GGC category on
   different runs (e.g. RV lot rent folded into GPR one run, broken out as "RV Site Rental
   Income" the next). This is the most common failure. The dollars are right; the *placement*
   is inconsistent.
2. **Wrong reporting period** — seller T-12s often contain multiple period columns side by
   side (e.g. "T-12 Ended 5/23" next to "Oct 2022–May 2023"). The model sometimes extracts the
   wrong column, or treats an 8-month partial period as if it were a full year.
3. **Values that don't tie out** — extracted monthly figures don't sum to the stated annual
   total, or line items don't sum to category subtotals, and nothing catches it before it
   reaches Excel.
4. **Untraceable final numbers** — occasionally the GGC Underwriting tab shows an NOI that
   doesn't trace back to the extracted inputs (a template-wiring bug, distinct from extraction).

**Root causes (diagnosed):**
- `temperature=0` does NOT guarantee determinism on hosted LLMs. Run-to-run drift comes from
  batch-invariance failures, GPU floating-point non-associativity, and MoE routing — not from
  sampling temperature. Do not assume temp=0 fixes consistency.
- The PDF parsing layer (currently Google Document AI Layout Parser) is a weak link for dense
  financial tables, and inconsistent parsing cascades into inconsistent extraction.
- There is no hard schema enforcement, so the model can place values in categories that
  shouldn't exist or vary its output structure.
- There is no deterministic reconciliation/tie-out layer between extraction and Excel writing.

**When working on accuracy, prioritize in this order:**
1. **Schema enforcement** — constrain category mappings to a hard enum the model literally
   cannot deviate from (Anthropic Structured Outputs beta, or strict tool use). This is the
   biggest lever for the category-drift problem.
2. **Reconciliation checks** — deterministic Python validators that confirm monthlies sum to
   annual, line items sum to subtotals, and EGI − OpEx = NOI, BEFORE anything is written to
   Excel. Surface failures, don't silently coerce.
3. **Better parsing** — evaluate replacing Google Document AI with a stronger table parser
   (Reducto, Tensorlake, or Azure Document Intelligence) if parsing variance is the bottleneck.
4. **Self-consistency** — for high-stakes fields, run extraction multiple times and majority-
   vote; flag fields where runs disagree for human review.

**Guiding principle for the tool's behavior: flag and prompt, never silently assume.**
When the tool is uncertain (mismatched unit counts, a value that doesn't tie out, an ambiguous
period), it should surface the issue for the user to resolve — transparency over convenience.

---

## Architecture / stack

- **`backend.py`** — Python Flask server, runs on port 5001. ~7,000 lines. The whole pipeline.
- **`index.html`** — single-page frontend (Tailwind CSS, drag-and-drop upload, localStorage).
- **`GGC_Blank_Underwriting_Sizer_Extended.xlsx`** — GGC's official 16-tab template, extended
  to 1,000 rent-roll rows. This is the canonical output format. Its SUMIFS formulas and exact
  category strings must be preserved — do not regenerate it from scratch.
- **APIs/services:** Anthropic Claude API (extraction + methodology), Google Document AI
  (PDF parsing), Google Maps Static + Street View (property images).
- **Deployment:** Vercel. Live demo at ggcunderwritingdemo.com.

### The pipeline (current design)

The financial side runs as a **multi-stage sequence**, deliberately split so each model does
one job well instead of one giant call doing everything (the single-call design was a major
source of the accuracy problems):

1. **Extraction** — Sonnet (deterministic settings) reads the Document-AI-parsed documents and
   pulls clean numbers: every line item with 12 monthly values + annual total, the correct
   reporting period, and the rent roll. **No GGC categorization here** — just faithful
   transcription.
2. **Verification** — pure Python, no AI, fully deterministic. Checks that the numbers tie out
   (monthlies → annual, rows → unit count, period is a full 12 months). Produces OK/WARN/FAIL
   checks surfaced on an **"Extraction Check" tab** so the reviewer can confirm the numbers tie
   before trusting anything downstream.
3. **Methodology** — Opus applies GGC's categorization and underwriting logic to the clean,
   verified data. Works from the extracted JSON, not the raw PDFs.
4. **Market research** — runs in parallel; web-searches for comps, market rents, alternative
   housing data.

PDFs still go through Google Document AI first (in `encode_file_for_claude` /
`parse_pdf_with_document_ai`), with a fallback to sending the raw PDF to Claude if Doc AI is
disabled or fails.

---

## GGC's underwriting methodology (encoded in the prompt)

The methodology prompt encodes GGC's specific rules. Key pieces, so you understand what the
output is supposed to do:

- **Collections (4-step):** GPR (from rent roll) → physical vacancy (tied to rent roll) →
  concessions (tied to T12) → bad debt (goal-seek plug so NRI ties to historical collections:
  T3 annualized if trending up, T12 if flat, T3 + flag if down, T6 for one-time distortions).
- **NRI = GPR − Vacancy − Concessions − Bad Debt. EGI = NRI + Other Income. NOI = EGI − OpEx.**
- **Lot-rent vs. home-rent bifurcation (THE structural centerpiece of MHC underwriting):**
  GGC is in the *land* business (lot rent, ~5% cap, premium) not the *home rental* business
  (home rent, ~15–20% cap, inferior). These must NEVER be blended. Output needs a **Total NOI
  column AND a Lot-Rent-Only NOI column**; home rent NOI = Total − Lot-Rent-Only. (Two-column
  output is a known gap still being built.)
- **POH (Park-Owned Home) handling:** POH count is a form input. The tool reconciles it against
  rent roll home-rent entries. POH % drives expense-ratio benchmarks (~25–35% all-TOH, 35–40%
  mixed, 45–50% high-POH). Home-rent-related expenses ("RM home," "POH maintenance") get
  re-bucketed into Home Rent Expense, not general R&M.
- **Management fee:** override the seller's number. 5% under 200 sites, 4% at 200+. EGI-based.
  Separate from payroll — both lines coexist.
- **Taxes:** primary method, then fall back to historical × 1.15. **Never below historical.**
- **Insurance:** T12 × 1.05; × 1.15 if in a flood zone (form toggle).
- **CapEx reserve:** $50/unit/year, added on top.
- **One-time items:** flag them with the month + amount, do NOT auto-strip.
- **Stabilized column:** replace contracted lot rents with **market** lot rents. Stabilized
  Yield on Cost = Stabilized NOI ÷ (Purchase Price + CapEx). Target ≥200 bps spread vs. the
  **market** cap rate (per a deal-profile grid GGC will provide), not just the ingoing cap.

**Property-type categories that must exist** (RV/mixed properties revealed these were missing
and caused mapping errors): RV Site Rental Income (separate from GPR), Parking Income, Retail,
and Omitt Income / Omitt Expense (explicit exclusion buckets for non-operating items like
vehicle costs).

**Preserve GGC's exact category strings**, including the deliberate misspelling "Electrcitiy"
if it's in their template — the SUMIFS formulas key off the exact strings.

---

## Coding conventions & working style

- **When editing `backend.py`:** it's a large single file. Make surgical, targeted edits.
  Always view the surrounding context before a `str_replace`. After editing, confirm the file
  still compiles (`python3 -m py_compile backend.py`) and that no references to removed
  names remain.
- **Validate changes with real test data** when possible. There are sample deals (Las Brisas,
  Whaleshead Beach Resort) with known-correct outputs to diff against. Test extraction logic in
  isolation with synthetic payloads before running the whole pipeline.
- **Don't break the template wiring.** The Excel template's formulas depend on exact category
  strings and cell positions. Changes to categorization must stay consistent with what the
  template expects.
- **Models:** extraction uses a deterministic Sonnet config (temperature=0, no thinking);
  methodology and market research use Opus (adaptive thinking, no temperature param — Opus 4.7+
  rejects temperature/top_p/top_k). Pin model versions explicitly; don't use `-latest`.
- **Cost is not the constraint; accuracy is.** Don't avoid an extra API call or a more capable
  model to save a few cents per deal if it improves accuracy or consistency.

### Communication / writing style (for any text you draft — emails, docs, comments)
- Natural and direct. No em dashes. No buzzwords. No formulaic or listy phrasing.
- Short and specific. Lead with the most important/concrete thing.
- Present work as ongoing, not finished. Don't overclaim scope.

---

## Key people & context

- **Michael Janabi** — Managing Partner, acquisitions/underwriting. Primary contact for the
  tool and the methodology. Reviews and ties out numbers himself.
- **Sean Cunningham** — Managing Partner, operations.
- **John Curia** — Director of Asset Management; built a separate address-input deal screener
  (complementary, not competing).

This is an unpaid internship-style project. The tool may eventually be sold to other MHC
operators, so treat it as a product, not a throwaway script.

---

## Current priorities (in rough order)

1. **Accuracy & consistency** — the whole point. Schema enforcement → reconciliation layer →
   better parsing → self-consistency. (See the core-problem section above.)
2. **Two-column NOI output** (Total + Lot-Rent-Only) — the biggest missing methodology feature.
3. **Add/verify the RV/mixed-property categories** (RV Site Rental Income, Parking Income,
   Retail, Omitt Income/Expense) and their routing rules.
4. **Fix any template-wiring bugs** where the final NOI doesn't trace to the extracted inputs.
5. **Run-to-run variance reduction** as a distinct workstream after the above.
