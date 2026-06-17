"""
build_template.py — generate the blank GGC underwriting template from
the CorrectOutput.xlsx gold standard.

Why: the prior blank template (GGC_Blank_Underwriting_Sizer_Extended.xlsx)
was hand-patched over time via fix_template.py and drifted from CorrectOutput's
layout in ways that didn't reproduce the gold-standard look. Starting from
CorrectOutput as the literal base — and only stripping the Whaleshead-specific
input values — guarantees the workbook produced by the engine matches the
analyst-built reference structurally and visually.

How: this script
  1. Loads `Claude/CorrectOutput.xlsx` (the analyst's hand-built model)
  2. Drops engine-managed tabs (Comps, questions, miscellaneous — the
     engine creates fresh `Comps Analysis`, `Extraction Check`, `Miscellaneous`
     tabs at fill_template time)
  3. Strips deal-specific input cells while preserving every formula,
     label, header, structural cell, number format, font, fill, border,
     column width, row height, and freeze pane
  4. Restores P4 (Purchase/Offer Price) as a formula reading from P9
     so the engine's single write to P9 cascades correctly
  5. Saves as `GGC_Blank_Underwriting_Sizer_Extended.xlsx` (the deployed
     template that backend.py loads at fill_template time)

Run: `python3 build_template.py`. This replaces `fix_template.py`'s former
role of patching the legacy blank. fix_template.py is kept as a thin
compatibility layer for any backend.py imports that still reference it.
"""
from __future__ import annotations
import re
from pathlib import Path
from openpyxl import load_workbook

SRC  = Path(__file__).parent / "Claude" / "CorrectOutput.xlsx"
DEST = Path(__file__).parent / "GGC_Blank_Underwriting_Sizer_Extended.xlsx"

# Rent Roll Input capacity. CorrectOutput was hand-tailored to Whaleshead's
# 148 units (scan range `Rent Roll Input!$X$3:$X$150`), but the deployed
# template needs to cover any deal. backend.py truncates at this same cap
# and emits a hard-fail extraction check when a real deal exceeds it.
RENT_ROLL_CAPACITY = 2000   # 2000 tenant rows → scan range 3:2002
LAST_RR_ROW = 2 + RENT_ROLL_CAPACITY


# ─────────────────────────────────────────────────────────────────────────────
# Per-sheet strip lists. Each entry is either a single coordinate ("N4") or a
# range ("M26:R32"). Only the VALUE on each cell is cleared — number format,
# font, fill, border, alignment, merged-cell state, etc. are preserved.
# ─────────────────────────────────────────────────────────────────────────────

# GGC Underwriting: the right-side property info block and tax-parcel section
# are deal-specific. Everything in the left-side P&L block (rows 3-50) is
# already formula-driven from Data Consolidation.
STRIP_UNDERWRITING = [
    # Property Information block (M3:R10)
    "N4",   # Property Name        — backend writes
    "N5",   # Property Address     — backend writes
    "N6",   # Property Type        — backend writes (when methodology provides)
    "N9",   # Acreage              — backend writes (when methodology provides)
    "N10",  # County               — backend writes
    "R3",   # Website URL          — backend writes
    "R4",   # Year Built           — backend writes
    "R5",   # Flood Zone           — backend writes
    "R6",   # Utility Structure    — backend writes
    "R7",   # Electricity notes    — backend writes
    "R8",   # Trash notes          — backend writes
    # Pricing block
    "P4",   # Purchase Price — REPLACED below with formula reading from P9
    "P9",   # Asking Price by Seller — backend writes
    "N2",   # Underwritten Date — backend writes
    # Tax Analysis Section (M19:R32)
    "N19",  # Tax assessor URL — backend writes
    # Parcel table M26:R32 — all parcel-specific data
    "M26", "N26", "O26", "P26", "Q26", "R26",
    "M27", "N27", "O27", "P27", "Q27", "R27",
    "M28", "N28", "O28", "P28", "Q28", "R28",
    "M29", "N29", "O29", "P29", "Q29", "R29",
    "M30", "N30", "O30", "P30", "Q30", "R30",
    "M31", "N31", "O31", "P31", "Q31", "R31",
    "M32", "N32", "O32", "P32", "Q32", "R32",
]

# Data Consolidation: all rows 3-36 (income leaves) and 43-102 (expense leaves)
# carry deal-specific values in columns A/B/D-G/J-U. Strip those; preserve
# the section-sum rows (38, 104), the input-source-check rows (39-40, 105),
# and the section markers (row 42 "Choose Expense Category" stays as a
# template separator — CorrectOutput keeps it).
DC_INCOME_LEAF_ROWS  = list(range(3, 37))   # rows 3-36
DC_EXPENSE_LEAF_ROWS = list(range(43, 103)) # rows 43-102
DC_STRIP_COLS = ["A", "B", "D", "E", "F", "G"]    # cat, name, fyPrior, fyCurrent, brokerProforma, t12
DC_STRIP_MONTHLY_COLS = list("JKLMNOPQRSTU")      # 12 monthly cols (J-U)

# Rent Roll Input: data rows 3 to LAST_RR_ROW (RENT_ROLL_CAPACITY rows
# after the 2 header rows). Strip columns B-J (unit id, type, status,
# name, type detail, type code, lot rent, home rent) — keep column A
# (Count formula = previous + 1) and column K (Combined formula = lot
# + home).
RR_DATA_ROWS = list(range(3, LAST_RR_ROW + 1))
RR_STRIP_COLS = list("BCDEFGHIJ")

# Comps: strip rows 3+ (comp data). Keep row 2 header. The engine adds a
# dedicated "Comps Analysis" tab at runtime, but keeping the gold-standard
# `Comps` tab with a clean structure lets analysts reference it as a hand-fill.
COMPS_STRIP_ROWS = list(range(3, 12))


def strip_cell(ws, coord):
    """Clear the value AND any attached hyperlink on a cell. Formatting
    (number format, font, fill, border, alignment, merged-cell state) is
    preserved — openpyxl's .value = None / .hyperlink = None mutations
    don't touch styling. CorrectOutput attaches hyperlinks to R3 (website)
    and N19 (county assessor URL) that survive a .value strip; the
    hyperlink object holds the URL even when the displayed text is blank,
    so it has to be explicitly nulled."""
    cell = ws[coord]
    cell.value = None
    cell.hyperlink = None


def main():
    if not SRC.exists():
        raise FileNotFoundError(f"Source gold standard missing: {SRC}")
    print(f"[build] Loading {SRC.name}")
    wb = load_workbook(SRC)
    print(f"[build] Source tabs: {wb.sheetnames}")

    # ── 1. Remove engine-managed tabs ───────────────────────────────────
    # `questions` and `miscellaneous` are analyst scratch tabs — the engine
    # doesn't use them and they shouldn't ship in the blank template.
    # The hand-built `Comps` tab from CorrectOutput is kept (analysts may
    # populate it manually); the engine's runtime `Comps Analysis` tab is
    # a different sheet and gets added separately at fill_template time.
    for name in ("questions", "miscellaneous"):
        if name in wb.sheetnames:
            del wb[name]
            print(f"[build] Removed '{name}' tab")

    # ── 1b. Strip external workbook links ───────────────────────────────
    # CorrectOutput references two external workbooks (probably the
    # seller's files or earlier GGC templates) that don't exist on the
    # deploy target. Excel pops "We found a problem with some content"
    # on open when these dangle. Clear them.
    if getattr(wb, "_external_links", None):
        n = len(wb._external_links)
        wb._external_links = []
        print(f"[build] Removed {n} external workbook link(s)")

    # ── 1c. Drop defined names with invalid characters ──────────────────
    # CorrectOutput has a defined name "CU?" that probably originally
    # had a non-ASCII char. Excel can't resolve it and warns on open.
    if hasattr(wb.defined_names, "delete"):
        bad = [n for n in list(wb.defined_names) if "?" in n or not n.isascii()]
        for n in bad:
            wb.defined_names.delete(n)
        if bad:
            print(f"[build] Removed {len(bad)} invalid defined name(s): {bad}")

    # ── 2. GGC Underwriting — strip property + parcel inputs ─────────────
    uw = wb["GGC Underwriting"]
    for coord in STRIP_UNDERWRITING:
        strip_cell(uw, coord)
    # P4 needs to remain functional. Restore as formula reading from P9 so
    # backend.py's single write to P9 cascades through to P4 (Purchase
    # Price), I47/P4 (cap rate), and the rest of the pricing chain.
    uw["P4"] = "=IFERROR(IF(ISNUMBER(P9),P9,0),0)"
    # Preserve P4's number format from CorrectOutput (currency).
    uw["P4"].number_format = '"$"#,##0_);[Red]\\("$"#,##0\\)'
    print(f"[build] GGC Underwriting: stripped {len(STRIP_UNDERWRITING)} input cells, "
          f"restored P4 formula")

    # ── 3. Data Consolidation — strip per-line input data ────────────────
    # Each leaf row carries the seller's account info (col B), category
    # tag (col A), FY columns (D/E/F), T12 (G), and 12 monthly values
    # (J-U). The engine writes all of these per-run; strip CorrectOutput's
    # Whaleshead values so the blank template carries no signal.
    dc = wb["Data Consolidation"]
    n_cleared = 0
    for r in DC_INCOME_LEAF_ROWS + DC_EXPENSE_LEAF_ROWS:
        for col in DC_STRIP_COLS + DC_STRIP_MONTHLY_COLS:
            cell = dc[f"{col}{r}"]
            if cell.value is not None and not (
                isinstance(cell.value, str) and cell.value.startswith("=")
            ):
                # Skip cells that hold a formula — preserve structural math.
                cell.value = None
                n_cleared += 1
    print(f"[build] Data Consolidation: cleared {n_cleared} per-line input cells "
          f"(94 leaf rows × ~18 cols)")

    # ── 4. Rent Roll Input — strip tenant rows ───────────────────────────
    # Per-tenant data lives in rows 3-1002 (the template scans 3:1002 in
    # its COUNTIFS/SUMIFS). Strip the input columns; keep column A
    # (=A_prev+1 count formula) and column K (=I+J combined rent).
    rr = wb["Rent Roll Input"]
    n_cleared = 0
    for r in RR_DATA_ROWS:
        for col in RR_STRIP_COLS:
            cell = rr[f"{col}{r}"]
            if cell.value is not None and not (
                isinstance(cell.value, str) and cell.value.startswith("=")
            ):
                cell.value = None
                n_cleared += 1
    print(f"[build] Rent Roll Input: cleared {n_cleared} per-tenant input cells "
          f"({len(RR_DATA_ROWS)} rows × {len(RR_STRIP_COLS)} cols)")

    # ── 4b. Extend Rent Roll scan ranges + per-row formulas ──────────────
    # CorrectOutput's Unit Mix Summary COUNTIFS/SUMIFS scan only rows 3:150
    # of Rent Roll Input (Whaleshead-specific 148-unit fit). Bump every
    # such reference to row LAST_RR_ROW (2002 with the default cap) so the
    # template works on any deal up to RENT_ROLL_CAPACITY tenants.
    print(f"[build] Extending Rent Roll scan ranges to row {LAST_RR_ROW}...")
    _range_pat = re.compile(
        r"(Rent Roll Input'?!\$?[A-Z]+\$?)(\d+):(\$?[A-Z]+\$?)(\d+)"
    )
    n_extended = 0
    for sname in wb.sheetnames:
        ws = wb[sname]
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str) or "Rent Roll Input" not in v:
                    continue
                def _bump(m):
                    return f"{m.group(1)}{m.group(2)}:{m.group(3)}{LAST_RR_ROW}"
                new = _range_pat.sub(_bump, v)
                if new != v:
                    cell.value = new
                    n_extended += 1
    print(f"[build] Rewrote {n_extended} formulas referencing Rent Roll Input ranges")

    # Extend the per-row formulas in column A (=A_prev+1 row count) and
    # column K (=I+J combined rent) so the new rows participate in the
    # count chain and combined-rent display. Without this, rows past 150
    # have no row count and combined-rent column stays blank.
    rr_ws = wb["Rent Roll Input"]
    n_seeded = 0
    for r in range(3, LAST_RR_ROW + 1):
        a_cell = rr_ws.cell(row=r, column=1)
        if a_cell.value is None or (isinstance(a_cell.value, int) and a_cell.value == r - 2):
            a_cell.value = f"=A{r-1}+1" if r > 3 else 1
            n_seeded += 1
        k_cell = rr_ws.cell(row=r, column=11)
        if k_cell.value is None:
            k_cell.value = f"=I{r}+J{r}"
    print(f"[build] Seeded {n_seeded} new Rent Roll Input row-formula pairs")

    # ── 5. Comps — strip comp data rows ──────────────────────────────────
    if "Comps" in wb.sheetnames:
        cp = wb["Comps"]
        n_cleared = 0
        for r in COMPS_STRIP_ROWS:
            for col_idx in range(1, cp.max_column + 1):
                cell = cp.cell(row=r, column=col_idx)
                if cell.value is not None and not (
                    isinstance(cell.value, str) and cell.value.startswith("=")
                ):
                    cell.value = None
                    n_cleared += 1
        print(f"[build] Comps: cleared {n_cleared} cells in data rows")

    # ── 6. Force Excel to fully recalculate on open ──────────────────────
    # Without this, openpyxl-written formulas show blank in Excel until
    # the user presses F9. Setting both calcMode=auto AND fullCalcOnLoad
    # covers the cases where Excel ignores one or the other.
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.calcCompleted = False
    wb.calculation.calcOnSave = True

    # ── 7. Save ──────────────────────────────────────────────────────────
    wb.save(DEST)
    print(f"[build] Wrote {DEST.name}")
    print(f"[build] Final tabs: {load_workbook(DEST).sheetnames}")


if __name__ == "__main__":
    main()
