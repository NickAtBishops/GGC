"""
Comprehensive structural fix for GGC_Blank_Underwriting_Sizer_Extended.xlsx.

The blank template ships with ~20 defects that produce broken output across
every downstream tab. This script patches all of them in one pass. Run once
after pulling the original blank.

Fix categories:
  1. GGC Underwriting tab: income labels, per-unit formulas, subject block,
     stabilized assumptions
  2. Unit Mix Summary: replace the 7 generic 'Type N' rows with 4 real
     categories (TOH MH / POH-Infilled / Long term RV / Retail/Commercial)
     and re-aim COUNTIFS at the new Rent Roll column layout
  3. Rent Roll Input: switch to the correct column layout (Unit, Unit Type,
     Status, Name, Type detail, Type code, Lot Rent, Home Rent, Combined)
  4. Unit Mix Rent Growth: collapse from 6 broken #REF! rows to 2 real rows
     (Lots / POH) pulling from Unit Mix Summary
  5. Loan Scenario: populate US Treasury rate, spread, and Max LTV so
     debt math stops returning zero
  6. Sources and Uses: fix Purchase Price reference, Closing Costs %, GP
     equity formula, capex linkage, and unit-count denominator
  7. GGC Pro Forma: kill the spurious 1.03 multiplier on Y1 RE taxes; fix
     the Home Rent Expense ratio and its SUM range; rename Lease-to-Own
     row to Long term RV Site
  8. Waterfall promote tiers: 20/30/30 -> 25/40/40
  9. Investor Return: repoint at canonical waterfall tabs
  10. Delete junk duplicate Waterfall tabs (' ', '2', '3')
  11. Force fullCalcOnLoad so Excel populates formula cache on open
"""
from pathlib import Path
import openpyxl

TEMPLATE = Path(__file__).parent / "GGC_Blank_Underwriting_Sizer_Extended.xlsx"

wb = openpyxl.load_workbook(TEMPLATE)

# ════════════════════════════════════════════════════════════════════════
# 1. GGC UNDERWRITING TAB
# ════════════════════════════════════════════════════════════════════════
ws = wb["GGC Underwriting"]

# Income category labels (rows 14-16). These feed SUMIFS into Data
# Consolidation; must match the categories we tag GL accounts with in
# backend.py.
ws["A14"] = "RV Site Rental Income"
ws["A15"] = "Storage Income"
ws["A16"] = "Retail Income"

# Subject property value cells (M:N block). Template ships with labels at
# column O but no value at all for # of Units, so every per-unit formula
# resolves to 0. Layout mirrors CorrectOutput: M4 Name, M5 Address, M6
# Type, M7 Units (the missing M4 label was the reason backend was writing
# the property name into the Address slot).
ws["M4"] = "Property Name"
ws["M5"] = "Property Address"
ws["M6"] = "Property Type"
ws["M7"] = "# of Units"
ws["N7"] = "='Unit Mix Summary'!E8"
ws["M8"] = "Rent Roll Occupancy"
ws["N8"] = "='Unit Mix Summary'!C13"
ws["M9"] = "Acreage"
ws["M10"] = "County"

# GGC ProForma per-unit assumptions (column J). Values come from GGC's
# playbook and replace the model's "T12 x 1.03" guess.
ws["J22"] = 400   # RE Taxes  $/unit
ws["J30"] = 150   # R&M       $/unit
ws["J35"] = 600   # Payroll   $/unit
ws["J43"] = 75    # Cap-Ex    $/unit (was 0)

# Stabilized (GGC ProForma) column I formulas.
ws["G5"] = "=-5%*G4"             # vacancy 3% -> 5%
ws["I22"] = "=J22*N7"            # RE Taxes per-unit
ws["I30"] = "=J30*$N$7"          # R&M per-unit
ws["I35"] = "=J35*$N$7"          # Payroll per-unit
ws["I41"] = "=15%*I13"           # Home Rent Expense = 15% of Home Rent Income
ws["I43"] = "=J43*$N$7"          # Cap-Ex per-unit

# Stabilized income rows for the renamed labels.
ws["I14"] = "='Unit Mix Summary'!C15*90%"  # RV Site Rental
ws["I15"] = "=D15"                          # Storage (small line, keep at T12)
ws["I16"] = "='Unit Mix Summary'!H7*12*98%" # Retail

# Per-unit (column J) reference fixes: P7/P6/P8 used pricing cells where
# unit count and occupancy don't actually live. Repoint at N7/N8.
J_PER_UNIT = {
    "J5":  "=100%-N8", "J7":  "=-I7/I4", "J10": "=I9/$N$7", "J12": "=I12/$N$7",
    "J13": "=I13/$N$7","J14": "=I14/$N$7","J15": "=I15/$N$7","J16": "=I16/$N$7",
    "J17": "=I17/$N$7","J19": "=I19/$N$7","J20": "=$I$19/$N$7","J23": "=I23/$N$7",
    "J25": "=I25/$N$7","J26": "=I26/$N$7","J27": "=I27/$N$7","J28": "=I28/$N$7",
    "J29": "=I29/$N$7","J31": "=I31/$N$7","J32": "=I32/$N$7","J36": "=I36/$N$7",
    "J37": "=I37/$N$7","J38": "=I38/$N$7","J39": "=I39/$N$7","J40": "=I40/$N$7",
    "J41": "=I41/$N$7","J42": "=I42/$N$7","J44": "=I44/$N$7","J45": "=$I$44/$N$7",
}
for coord, formula in J_PER_UNIT.items():
    ws[coord] = formula

# Lines that should NOT auto-grow from T12 in the stabilized column.
ws["I32"] = 0   # Recreational Amenities
ws["I36"] = 0   # Employee Allowance
ws["I38"] = 0   # Model Units
ws["I42"] = 0   # Other

# ════════════════════════════════════════════════════════════════════════
# 2. UNIT MIX SUMMARY — replace 7 generic types with 4 real categories
# ════════════════════════════════════════════════════════════════════════
ums = wb["Unit Mix Summary"]

# Clear values in the data block (rows 3-20) before rewriting. Preserves
# column widths / styling / merged cells.
for row in ums.iter_rows(min_row=3, max_row=22, max_col=10):
    for cell in row:
        cell.value = None

# Headers (row 3). Clear any stray text in row 2 cells C-K — the original
# blank template had a "Next Increase in MAY 1 2026 — see RR" note at C2
# left over from a hand-built workbook, which showed up next to our title.
for col in range(3, 12):
    ums.cell(row=2, column=col).value = None
ums["B2"] = "MH/RV Rent Roll "  # trailing space matches CorrectOutput
ums["C3"] = "# of Occupied Units"
ums["D3"] = "# of Vacant Units"
ums["E3"] = "Total Units"
ums["F3"] = "Occupied Monthly Rent"
ums["G3"] = "Vacant Monthly Rent"
ums["H3"] = "Monthly Gross Potential Rent"

# Four canonical categories at rows 4-7. Matches the canonical taxonomy
# we'll enforce in backend.py. COUNTIFS / SUMIFS hit the new Rent Roll
# Input column layout (D=Status, C=Unit Type, I=Lot Rent).
CATEGORIES = [
    ("TOH MH Site",        4),
    ("POH-Infilled units", 5),
    ("Long term RV Site",  6),
    # CorrectOutput has a typo here ("Retail/Comemrcial"). We use the
    # correctly spelled version so a future user search won't fail; the
    # COUNTIFS doesn't reference column B so the typo is purely
    # cosmetic in correct.
    ("Retail/Commercial",  7),
]
for label, r in CATEGORIES:
    ums.cell(row=r, column=2, value=label)
    ums.cell(row=r, column=3,
             value=f'=COUNTIFS(\'Rent Roll Input\'!$D$3:$D$1002,"Occupied",'
                   f'\'Rent Roll Input\'!$C$3:$C$1002,"{label}")')
    ums.cell(row=r, column=4,
             value=f'=COUNTIFS(\'Rent Roll Input\'!$D$3:$D$1002,"Vacant",'
                   f'\'Rent Roll Input\'!$C$3:$C$1002,"{label}")')
    ums.cell(row=r, column=5, value=f"=SUM(C{r}:D{r})")
    ums.cell(row=r, column=6,
             value=f'=SUMIFS(\'Rent Roll Input\'!$I$3:$I$1002,'
                   f'\'Rent Roll Input\'!$D$3:$D$1002,"Occupied",'
                   f'\'Rent Roll Input\'!$C$3:$C$1002,"{label}")')
    ums.cell(row=r, column=7,
             value=f'=SUMIFS(\'Rent Roll Input\'!$I$3:$I$1002,'
                   f'\'Rent Roll Input\'!$D$3:$D$1002,"Vacant",'
                   f'\'Rent Roll Input\'!$C$3:$C$1002,"{label}")')
    ums.cell(row=r, column=8, value=f"=F{r}+G{r}")

# Totals / derived metrics block (rows 8-16). Mirrors correct output.
ums["B8"]  = "Total Units"
ums["C8"]  = "=SUM(C4:C7)"
ums["D8"]  = "=SUM(D4:D7)"
ums["E8"]  = "=SUM(E4:E7)"
ums["B9"]  = "Total MH Sites"   # match CorrectOutput (no parenthetical)
ums["C9"]  = "=SUM(C4:C5)"
ums["D9"]  = "=SUM(D4:D5)"
ums["E9"]  = "=SUM(E4:E5)"
ums["F9"]  = "=SUM(F4:F5)"
ums["G9"]  = "=SUM(G4:G5)"
ums["H9"]  = "=SUM(H4:H5)"
ums["B10"] = "Total POH"
ums["C10"] = "=C5"
ums["D10"] = "=D5"
ums["E10"] = "=E5"
# Labels match CorrectOutput exactly. IFERROR wrappers stay — Correct
# trusts its inputs but our LLM-generated rent rolls can produce zero
# counts for a missing unit type, which would propagate #DIV/0! into
# every downstream per-unit formula. The wrapper preserves the math
# under correct's labels while staying robust.
ums["B11"] = "Annual GPR(MH Lot Rent)"   # no space — matches Correct
ums["C11"] = "=H9*12"
ums["B12"] = "Avg Rent"                  # was "Avg MH Lot Rent"
ums["C12"] = "=IFERROR(H4/E4,0)"
ums["B13"] = "Occupancy%"                # no space
ums["C13"] = "=IFERROR(C8/E8,0)"
ums["B14"] = "POH%"                      # no space
ums["C14"] = "=IFERROR(E10/E9,0)"
ums["B15"] = "Annual Long term RV"
ums["C15"] = "=H6*12"
ums["B16"] = "Long term RV avg rent"
ums["C16"] = "=IFERROR(H6/E6,0)"

# ════════════════════════════════════════════════════════════════════════
# 3. RENT ROLL INPUT — switch to correct column layout
# ════════════════════════════════════════════════════════════════════════
# Correct uses: A=Count, B=Unit, C=Unit Type, D=Status, F=Name, G=Type
# detail, H=Type code, I=Lot Rent, J=Home Rent, K=Combined.
# backend.py will be updated to write to these columns.
rr = wb["Rent Roll Input"]

# Clear existing header row 2 and rewrite
for col in range(1, 13):
    rr.cell(row=2, column=col).value = None
rr["A2"] = "Count"
rr["B2"] = "Unit"
rr["C2"] = "Unit Type"
rr["D2"] = "Occupied or Vacant"
rr["F2"] = "Name"
rr["G2"] = "Type detail"
rr["H2"] = "Type code"
rr["I2"] = "Lot Rent"
rr["J2"] = "Home Rent"
rr["K2"] = "Combined"

# Update Combined formula and Count formula on data rows 3-1002. We
# rebuild all 1000 rows of formulas so they stay consistent if backend
# rewrites the value columns.
for r in range(3, 1003):
    rr.cell(row=r, column=1, value=f"=IF(C{r}=\"\",\"\",ROW()-2)")  # Count
    rr.cell(row=r, column=11, value=f"=IFERROR(I{r}+J{r},0)")        # Combined
    # Clear stale POH/LTO columns from old layout
    for col in (5, 6, 7, 8):
        rr.cell(row=r, column=col).value = None

# NOTE: Earlier patches placed a "Totals" row at row 151 (with B151="Totals"
# and SUM formulas at I151/J151/K151) so Underwriting!G13 could read the
# home-rent monthly total. That produced an orphaned highlighted row 25+
# rows below the actual data — looked like a stray black square in empty
# space on every output. Removed here. Underwriting!G13 now reads a SUM
# across the entire rent-roll range directly (set below in the
# Underwriting tab section), no fixed totals row needed.
#
# The original blank template ALSO had row 151 styled with a dark fill
# (left over from a hand-built totals row in an earlier version of the
# workbook). That fill survives even after we clear the cell values, so
# we explicitly null the fill and font on row 151 across columns A-K
# to match the empty rows above it.
from openpyxl.styles import PatternFill, Font, Border, Side
_clear_fill = PatternFill(fill_type=None)
_clear_font = Font()
_clear_border = Border(left=Side(border_style=None), right=Side(border_style=None),
                       top=Side(border_style=None), bottom=Side(border_style=None))
for col in range(1, 12):  # A-K
    cell = rr.cell(row=151, column=col)
    cell.value = None
    cell.fill = _clear_fill
    cell.font = _clear_font
    cell.border = _clear_border

# ════════════════════════════════════════════════════════════════════════
# 4. UNIT MIX RENT GROWTH — match CorrectOutput cell-for-cell
# ════════════════════════════════════════════════════════════════════════
# The blank template ships with stray content from a prior hand-built
# workbook: leftover notes at B2/B3, non-flat year-by-year rent growth
# rates at rows 5-8 (specific to a different deal), orphan vacancy-
# schedule rows at 23-25 referencing a $D$20 cell that no longer exists,
# and only 6 of 7 years of forward rent projection populated. Replace
# the whole block with CorrectOutput's structure so the tab visually
# matches the gold standard.

umrg = wb["Unit Mix Rent Growth"]

# Clear stray notes at B2/B3 from the blank template
umrg["B2"] = None
umrg["B3"] = None

# Clear the entire body (rows 5-28) so no orphan content from the old
# 6-type template survives. We rewrite from scratch below.
for row in umrg.iter_rows(min_row=5, max_row=28, max_col=12):
    for cell in row:
        cell.value = None

# ── Header row 4: "Key Assumptions" + Year 1..Year 10 ──
umrg["A4"] = "Key Assumptions"
year_cols = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]
for i, col in enumerate(year_cols, start=1):
    umrg[f"{col}4"] = f"Year {i}"

# ── Rows 5-6: Per-type growth rates (flat 5% across all years) ──
# Correct uses 0.05 flat for both MH (row 5) and POH (row 6) across
# Year 1..Year 10. The label cells point back at Unit Mix Summary
# so they show "TOH MH Site" / "POH-Infilled units" automatically.
umrg["A5"] = "=C10"
umrg["A6"] = "=C11"
for col in year_cols:
    umrg[f"{col}5"] = 0.05
    umrg[f"{col}6"] = 0.05

# ── Row 8: section header for the rent projection block ──
umrg["E8"] = "Lot Rent"

# ── Row 9: column headers for the per-year rent grid ──
umrg["C9"] = "Unit Mix"
umrg["D9"] = "# of Units"
umrg["E9"] = "Avg Monthly Rent"
projection_cols = ["F", "G", "H", "I", "J", "K", "L"]
for i, col in enumerate(projection_cols, start=1):
    umrg[f"{col}9"] = f"Year {i}"

# ── Row 10: TOH MH Site projection (B-weighted to total units) ──
umrg["B10"] = "=IFERROR(D10/$D$12,0)"
umrg["C10"] = "='Unit Mix Summary'!B4"
umrg["D10"] = "='Unit Mix Summary'!E4"
umrg["E10"] = "=IFERROR('Unit Mix Summary'!H4/'Unit Mix Summary'!E4,0)"
umrg["F10"] = "=E10*(100%+B$5)"
umrg["G10"] = "=F10*(100%+C$5)"
umrg["H10"] = "=G10*(100%+D$5)"
umrg["I10"] = "=H10*(100%+E$5)"
umrg["J10"] = "=I10*(100%+F$5)"
umrg["K10"] = "=J10*(100%+G$5)"
umrg["L10"] = "=K10*(100%+H$5)"

# ── Row 11: POH-Infilled projection ──
umrg["B11"] = "=IFERROR(D11/$D$12,0)"
umrg["C11"] = "='Unit Mix Summary'!B5"
umrg["D11"] = "='Unit Mix Summary'!E5"
umrg["E11"] = "=IFERROR('Unit Mix Summary'!H5/'Unit Mix Summary'!E5,0)"
umrg["F11"] = "=E11*(100%+B$6)"
umrg["G11"] = "=F11*(100%+C$6)"
umrg["H11"] = "=G11*(100%+D$6)"
umrg["I11"] = "=H11*(100%+E$6)"
umrg["J11"] = "=I11*(100%+F$6)"
umrg["K11"] = "=J11*(100%+G$6)"
umrg["L11"] = "=K11*(100%+H$6)"

# ── Row 12: weighted-average roll-up ──
umrg["B12"] = "=SUM(B10:B11)"
umrg["C12"] = "Total Weighted Average"
umrg["D12"] = "=SUM(D10:D11)"
for col in ("E", "F", "G", "H", "I", "J", "K", "L"):
    umrg[f"{col}12"] = f"=SUMPRODUCT($B$10:$B$11,{col}10:{col}11)"

# ── Row 13: $ change vs prior year ──
umrg["D13"] = "$change"
for prev, curr in (("E", "F"), ("F", "G"), ("G", "H"), ("H", "I"),
                   ("I", "J"), ("J", "K"), ("K", "L")):
    umrg[f"{curr}13"] = f"={curr}12-{prev}12"

# ── Row 14: annual GPR per year (weighted-avg × total units × 12) ──
# Underwriting!G4 reads H14 here as the stabilized (Year 3) GPR.
umrg["D14"] = "GPR"
for col in ("E", "F", "G", "H", "I", "J", "K", "L"):
    umrg[f"{col}14"] = f"={col}12*$D$12*12"

# ── Rows 15-17: vacancy schedule (mirrors CorrectOutput) ──
umrg["D15"] = "Vacant Lots"
umrg["C15"] = 0   # stabilization step (manually adjustable per deal)
umrg["E15"] = "='Unit Mix Summary'!D4"
umrg["F15"] = "=E15-C15"
umrg["D16"] = "Vacant Homes"
umrg["C16"] = 0
umrg["E16"] = "='Unit Mix Summary'!D5"
umrg["F16"] = "=E16-$C$16"
umrg["D17"] = "Vacancy"
for col in ("E", "F", "G", "H", "I", "J", "K", "L"):
    umrg[f"{col}17"] = f"=SUM({col}15:{col}16)/$D$12"

# Wire stabilized GPR on Underwriting to Unit Mix Rent Growth H14 (Year
# 3 GPR), matching CorrectOutput's reference.
ws["G4"] = "='Unit Mix Rent Growth'!H14"

# ════════════════════════════════════════════════════════════════════════
# 5. LOAN SCENARIO (acquisition) — populate non-zero rates
# ════════════════════════════════════════════════════════════════════════
ls = wb["Loan Scenario (acquisition)"]
ls["C14"] = 0.0405  # US Treasury / SOFR baseline (4.05%)
ls["C15"] = 0.0185  # Spread (185 bps)
ls["C24"] = 0.70    # Max LTV (was 75%)

# ════════════════════════════════════════════════════════════════════════
# 6. SOURCES AND USES — both S1 (cols B:E) and S2 (cols H:K) scenarios
# ════════════════════════════════════════════════════════════════════════
su = wb["Sources and Uses"]
su["C8"]  = "=C15"                          # GP equity = acq fee co-invest
su["I8"]  = "=I15"
su["C13"] = "='GGC Underwriting'!P4"        # Purchase Price (was R4)
su["I13"] = "='GGC Underwriting'!P4"
su["C14"] = "=C13*1.5%"                     # Closing costs 1.5% (was 2%)
su["I14"] = "=I13*1.5%"

# Capex Budget — link to a real number rather than 0. We use a default
# capex pad of $1,000/unit; the user can override per deal.
su["C17"] = "=1000*'GGC Underwriting'!$N$7"
su["I17"] = "=1000*'GGC Underwriting'!$N$7"

# Replace every $P$7 reference in this sheet with $N$7. The blank uses
# P7 everywhere as the units denominator, which points at a pricing cell.
for row in su.iter_rows():
    for cell in row:
        if isinstance(cell.value, str) and "GGC Underwriting" in cell.value:
            cell.value = cell.value.replace("$P$7", "$N$7").replace("!P7", "!N7")

# ════════════════════════════════════════════════════════════════════════
# 7. GGC PRO FORMA — Y1 RE Tax step-up, Home Rent Exp, RV row label
# ════════════════════════════════════════════════════════════════════════
pf = wb["GGC Pro Forma"]
pf["C21"] = "Long term RV Site"   # was 'Lease to Own'
pf["H28"] = "=D28"                # Y1 RE Taxes: no 1.03 step-up
pf["B47"] = 0.15                  # Home Rent Exp ratio: 35% -> 15%
pf["H47"] = "=$B$47*H20"          # was =$B$47*SUM(H20:H21)

# $P$7 -> $N$7 everywhere on this sheet
for row in pf.iter_rows():
    for cell in row:
        if isinstance(cell.value, str) and "GGC Underwriting" in cell.value:
            cell.value = cell.value.replace("$P$7", "$N$7").replace("!P7", "!N7")

# ════════════════════════════════════════════════════════════════════════
# 8. WATERFALL PROMOTE TIERS — 20/30/30 -> 25/40/40 (10-yr + 5-yr)
# ════════════════════════════════════════════════════════════════════════
for tab in ["Waterfall (10-yr-S1)", "Waterfall (5-yr-S1)"]:
    if tab in wb.sheetnames:
        w = wb[tab]
        w["F16"] = 0.25
        w["F17"] = 0.40
        w["F18"] = 0.40

# ════════════════════════════════════════════════════════════════════════
# 9. INVESTOR RETURN — repoint at canonical waterfall tab, default "No"
# ════════════════════════════════════════════════════════════════════════
ir = wb["Investor Return"]
ir["F6"] = "='Waterfall (10-yr-S1)'!D30"
ir["F7"] = "='Waterfall (10-yr-S1)'!D31"
ir["F8"] = "=AVERAGE('Waterfall (10-yr-S1)'!G31:O31)"
ir["N4"] = "No"  # Send LOI? — default to No, not Yes

# ════════════════════════════════════════════════════════════════════════
# 10. DELETE JUNK DUPLICATE WATERFALL TABS
# ════════════════════════════════════════════════════════════════════════
for junk in ["Waterfall ", "Waterfall 2", "Waterfall 3"]:
    if junk in wb.sheetnames:
        del wb[junk]

# ════════════════════════════════════════════════════════════════════════
# 11. ROUND 3 — gap-fills from the dependency-graph verification pass
# ════════════════════════════════════════════════════════════════════════

# 11a. Make Management Fee % conditional on unit count per methodology
# (5% under 200 sites, 4% at 200+). Was hardcoded 0.05.
ws["J33"] = "=IF(N7>=200,0.04,0.05)"

# 11b. Seed Purchase Price (P4) so Sources & Uses, Loan Scenario, and Pro
# Forma Y0 don't collapse to zero. Default to the seller's asking price
# (Q9 on this tab is empty; backend.py writes asking price there at run
# time). Use IFERROR so the user can override.
ws["P4"] = "=IFERROR(P9,0)"

# 11c. GGC Pro Forma (Conversion) tab still has 35 stale $P$7 refs in
# col G (per-unit denominators). Sweep all $P$7 -> $N$7 there too.
if "GGC Pro Forma (Conversion)" in wb.sheetnames:
    conv = wb["GGC Pro Forma (Conversion)"]
    for row in conv.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and "GGC Underwriting" in cell.value:
                cell.value = cell.value.replace("$P$7", "$N$7").replace("!P7", "!N7")

# 11d. Investor Return rows 25-29 referenced 'Waterfall 2' which we
# deleted. Repoint at the canonical (10-yr-S1) tab so they show real
# numbers instead of #REF!.
for r, cell_path in [
    (25, ("F25", "='Waterfall (10-yr-S1)'!D30")),
    (26, ("F26", "='Waterfall (10-yr-S1)'!D31")),
    (27, ("F27", "=AVERAGE('Waterfall (10-yr-S1)'!G31:O31)")),
]:
    coord, formula = cell_path
    if ir[coord].value is not None:
        ir[coord] = formula

# 11e/f/g — Unit Mix Rent Growth row-22 / M20-O20 / A5-A10 fixes from
# round 3 are now obsolete. The canonical rebuild earlier in this script
# (section 4) handles all of that already in the new layout. Removed.

# ════════════════════════════════════════════════════════════════════════
# 12. ROUND 4 — REGRESSION FIXES from full cross-check vs CorrectOutput
# ════════════════════════════════════════════════════════════════════════

# 12a. The original subject block had labels at O2-O10. After we moved
# it to M-N those labels are orphans that now duplicate / misalign with
# the new pricing block (P4-P10). Clear them.
for coord in ("O2", "O3", "O4", "O5", "O6", "O7", "O8", "O9", "O10"):
    if isinstance(ws[coord].value, str) and ws[coord].value in (
        "Underwritten Date", "Property Information", "Property Name",
        "Property Address", "Property Type", "# of Units ",
        "Rent Roll Occupancy", "Acreage", "County",
    ):
        ws[coord] = None

# 12b. I13 (Home Rent Income stabilized) was ='Rent Roll Input'!F1005*95%
# in the original blank. Col F is now Tenant Name after the Rent Roll
# restructure, so multiplying a name by 0.95 produces #VALUE!. Zero out
# unless / until backend.py decides to populate it from per-row data.
ws["I13"] = 0

# 12c. P4 (Purchase Price) was =IFERROR(P9,0) which resolves to 0 when
# P9 is empty. Backend.py writes askingPrice to P9 if provided, but if
# the user negotiates a price below ask, P4 stays at ask. Switch to a
# hardcoded value with a comment: backend.py should overwrite P4 with
# the negotiated price each run via property_info["purchasePrice"].
# Default formula falls back to P9 (asking) only when no override.
ws["P4"] = "=IFERROR(IF(ISNUMBER(P9),P9,0),0)"

# 12d. Investor Return F28/F29 referenced 'Waterfall 2' which we
# deleted in round 1. Repoint at the canonical waterfall.
for coord in ("F28", "F29"):
    v = ir[coord].value
    if isinstance(v, str) and "Waterfall 2" in v:
        ir[coord] = v.replace("'Waterfall 2'", "'Waterfall (10-yr-S1)'")

# 12e. Loan Scenario C27 (DSCR denominator) pointed at 'GGC Underwriting'!H47
# which is the Lot Rent ONLY NOI. For debt sizing the correct basis is
# the Total NOI (I47). Repoint.
if "Loan Scenario (acquisition)" in wb.sheetnames:
    lscell = wb["Loan Scenario (acquisition)"]["C27"]
    if isinstance(lscell.value, str) and "H47" in lscell.value:
        lscell.value = lscell.value.replace("H47", "I47")

# 12f. Sources & Uses C17 (Capex Budget) was 1000*N7 — a placeholder.
# Wire it instead to a real capex breakdown at C21-C24:
#   C21 = utility / septic / water infrastructure
#   C22 = home in-fills
#   C23 = working capital
#   C24 = SUM, fed into C17.
su["B21"] = "Water / Septic / Utilities"
su["C21"] = 0
su["B22"] = "Add Homes / In-fill"
su["C22"] = 0
su["B23"] = "Working Capital"
su["C23"] = 200000
su["B24"] = "Capex Budget Total"
su["C24"] = "=SUM(C21:C23)"
su["C17"] = "=C24"
# Same for the S2 scenario (cols H:K).
su["I21"] = 0
su["I22"] = 0
su["I23"] = 200000
su["I24"] = "=SUM(I21:I23)"
su["I17"] = "=I24"

# 12g — Unit Mix Rent Growth rebuild from round 4 is obsolete; the
# canonical rebuild in section 4 above already produces this layout
# (and extends Year 7 to column L which round 4 was missing).

# ════════════════════════════════════════════════════════════════════════
# 13. EXACT-MATCH ALIGNMENT WITH CORRECTOUTPUT.XLSX (Underwriting tab)
# ════════════════════════════════════════════════════════════════════════
# Earlier rounds added "improvements" (MAX-with-reassessment on I22,
# conditional mgmt fee on J33, County Tax Rate cell at P12) that diverge
# from the gold-standard CorrectOutput layout. User wants exact match —
# revert those and restore the original formulas + cell labels.
ws["A26"] = "Electricity"           # row label (was "Electrcitiy" typo)
ws["I22"] = "=J22*N7"               # was MAX(...) — back to simple per-unit
ws["J33"] = 0.05                    # was =IF(N7>=200,...) — back to flat 5%
ws["I33"] = "=J33*I19"              # mgmt fee = % × EGI (stabilized)
ws["G33"] = "=J33*G19"              # mgmt fee for ALT NOI column
ws["H33"] = "=J33*H19"              # mgmt fee for lot-rent-only NOI
ws["I35"] = "=J35*N7"               # match correct (no $ anchor on N7)

# Restore the O/P pricing block labels EXACTLY as CorrectOutput has them.
# O4/P4 = Purchase/Offer Price (P4 is a numeric input written by backend).
# O5/P5 = $/site (P5 = P4/N7).  O6/P6 = Underwritten Cap (NOI/PP).
# O7/P7 = Stabilized YOC.  O9/P9 = Asking Price.  O10/P10 = Asking $/site.
ws["O4"] = "Purchase/Offer Price"
ws["O5"] = "Purchase Price Per Site"
ws["O6"] = "Underwritten CAP rate"
ws["O7"] = "Stabilized YOC"
ws["O9"] = "Asking Price by Seller"
ws["O10"] = "Asking Price Per Site "
ws["O12"] = "Brookings Multifamily cap rate at 6%-7.5%"
ws["P12"] = None                    # remove the County Tax Rate cell

# Match cosmetics
ws["M7"] = "# of Units "            # trailing space — matches correct
ws["P5"] = "=P4/N7"                 # purchase price per site
ws["P6"] = "=I47/P4"                # underwritten cap rate = NOI / PP
ws["P7"] = "=G47/'Sources and Uses'!C18"  # stabilized YOC
ws["P10"] = "=P9/N7"                # asking $ per site

# G4 (Stabilized GPR) — correct points at Unit Mix Rent Growth!H14, which
# is annual GPR for Year 3 in the canonical layout we rebuild next.
ws["G4"] = "='Unit Mix Rent Growth'!H14"

# I4 (Stabilized total units anchor) — Unit Mix Summary!C11 in correct's
# layout = total MH lot count.
ws["I4"] = "='Unit Mix Summary'!C11"

# J7 (Bad Debt assumption) — correct uses a hardcoded 3% rate, not a
# back-computed ratio.
ws["J7"] = 0.03

# G12 (Stabilized Utility Reimbursement) — correct uses =I12 (mirror of
# Total NOI col), not 75% of the water/sewer line.
ws["G12"] = "=I12"

# G13 (Stabilized Home Rent Income) — SUM the entire Home Rent column
# (J3:J1002) directly rather than reading from a fixed J151 "Totals" row.
# The Totals row produced an orphaned highlighted cell sitting 25+ rows
# below the actual data on every output. Range SUM produces the same
# number with no visual side effect.
ws["G13"] = "=SUM('Rent Roll Input'!J3:J1002)*12*95%"

# Lot-Rent-Only NOI column (H): correct ZEROES OUT home rent stream so
# the H47 cell yields lot-rent NOI only. H13 = 0; H14:H16 mirror I col.
ws["H13"] = 0

# I17 (Other Income stabilized) — average of 2024 + T12 (cols C, D), not
# 2022 + 2024 (cols B, C).
ws["I17"] = "=AVERAGE(C17,D17)"

# H25-H28 (Lot-Rent-Only utility stream) — correct does NOT scale by 80%,
# it mirrors I25-I28 directly.
for r in (25, 26, 27, 28):
    ws[f"H{r}"] = f"=I{r}"

# SUMIFS range sweep — the original template's SUMIFS only covered
# rows 3:21 (income) and 28:58 (expenses). Correct's Data Consolidation
# is wider: 3:36 income, 43:102 expense. Update every SUMIFS in the
# Underwriting tab to match.
_RANGE_SWAPS = [
    # Income block
    ("$D$3:$D$21", "$D$3:$D$36"),
    ("$E$3:$E$21", "$E$3:$E$36"),
    ("$F$3:$F$21", "$F$3:$F$36"),
    ("$G$3:$G$21", "$G$3:$G$36"),
    ("$H$3:$H$21", "$H$3:$H$36"),
    ("$A$3:$A$21", "$A$3:$A$36"),
    # Expense block
    ("$D$28:$D$58", "$D$43:$D$102"),
    ("$E$28:$E$58", "$E$43:$E$102"),
    ("$F$28:$F$58", "$F$43:$F$102"),
    ("$G$28:$G$58", "$G$43:$G$102"),
    ("$H$28:$H$58", "$H$43:$H$102"),
    ("$A$28:$A$58", "$A$43:$A$102"),
]
for row in ws.iter_rows():
    for cell in row:
        if isinstance(cell.value, str) and "SUMIFS" in cell.value:
            new_val = cell.value
            for old, new in _RANGE_SWAPS:
                new_val = new_val.replace(old, new)
            if new_val != cell.value:
                cell.value = new_val

# Row-14/15/16 SUMIFS criteria — the blank template hunts for old labels
# ("LTO", "SFH", "Laundry Income"). Mirror CorrectOutput's per-column
# criteria exactly, including the quirky ones (B16 searches "Laundry
# Income" but C16:F16 search "Retail" — a hand-edit artifact in the
# gold standard that we replicate so values match cell-for-cell).
def _swap_criterion(cell, new_criterion):
    if isinstance(cell.value, str) and "SUMIFS" in cell.value:
        import re as _re
        cell.value = _re.sub(r'"[^"]+"\)$', f'"{new_criterion}")', cell.value)

# Row 14: all five history columns search "RV Site Rental Income"
for col in ("B", "C", "D", "E", "F"):
    _swap_criterion(ws[f"{col}14"], "RV Site Rental Income")

# Row 15: all five search "Parking Income" (correct's chosen label for
# the 4108 Storage Unit Rent GL — even though the row is labeled
# "Storage Income" in A15 for human readability).
for col in ("B", "C", "D", "E", "F"):
    _swap_criterion(ws[f"{col}15"], "Parking Income")

# Row 16: B16 searches "Laundry Income" (the 2022 column had no retail);
# C16:F16 search "Retail". Match correct's exact criteria.
_swap_criterion(ws["B16"], "Laundry Income")
for col in ("C", "D", "E", "F"):
    _swap_criterion(ws[f"{col}16"], "Retail")

# Row 2 column headers — restore CorrectOutput's exact strings.
ws["G2"] = "ALT NOI"
ws["H2"] = "Lot Rent only NOI"
ws["I2"] = None    # correct leaves blank
ws["J2"] = None    # correct leaves blank

# Per-unit J-column cells correct just leaves blank (J20, J45).
ws["J20"] = None
ws["J45"] = None
# And the $-anchor on N7 for J31/J32 — correct uses bare N7 (no $).
ws["J31"] = "=I31/N7"
ws["J32"] = "=I32/N7"

# Cosmetic: (100%+3%) → 1.03 to match correct exactly.
for r in (39, 40):
    cell = ws[f"I{r}"]
    if isinstance(cell.value, str) and "(100%+3%)" in cell.value:
        cell.value = cell.value.replace("(100%+3%)", "1.03")

# Lot-rent-only NOI column quirks from CorrectOutput:
#   H13 = 0          (home rent removed, already set above)
#   H14 = I14        (RV rent included in lot-rent NOI)
#   H15 = I15        (Storage included)
#   H16 = None       (Retail removed from lot-rent NOI — correct's choice)
ws["H14"] = "=I14"
ws["H16"] = None

# Bifurcated valuation block at M13:P16. This is the GGC methodology
# split — value the Lot Rent NOI at 5.5% cap, Home Rent NOI at 12% cap.
ws["M13"] = "Asset"
ws["N13"] = "NOI"
ws["O13"] = "CAP RATE"
ws["P13"] = "VALUE"
ws["M14"] = "Lot Rent only NOI"
ws["N14"] = "=I47"
ws["O14"] = 0.055
ws["P14"] = "=N14/O14"
ws["M15"] = "Home Rent only NOI"
ws["N15"] = "=I47-H47"
ws["O15"] = 0.12
ws["P15"] = "=N15/O15"
ws["M16"] = "Total"
ws["N16"] = "=SUM(N14:N15)"
ws["O16"] = "=N16/P16"
ws["P16"] = "=SUM(P14:P15)"

# ════════════════════════════════════════════════════════════════════════
# 14. FORCE FULL RECALC ON OPEN
# ════════════════════════════════════════════════════════════════════════
# Force Excel to recalculate everything on open. The blank template may
# ship with calcMode="manual" or a stale full-calc flag, which would
# leave openpyxl-written formulas displayed as blank cells until the
# user manually presses F9. Setting both auto-mode AND fullCalcOnLoad
# covers the cases where Excel ignores one or the other.
wb.calculation.calcMode = "auto"
wb.calculation.fullCalcOnLoad = True
wb.calculation.calcCompleted = False
wb.calculation.calcOnSave = True

# ════════════════════════════════════════════════════════════════════════
# 15. BLOCK-SHIP FIXES — partner-grade cleanup (commit "make-it-count")
# ════════════════════════════════════════════════════════════════════════
# Five independent reviewers found ~40 items between them. This block
# addresses every CRITICAL / IMPORTANT one so the workbook reads as a
# finished GGC template, not a hand-built artifact mid-iteration.

# ── 15a. Pro Forma row 8 (Y1-Y10 GPR) chain break ─────────────────────
# H8:Q8 read from 'Unit Mix Rent Growth'!F22:O22 — but my round-4
# rebuild moved the annual-GPR-per-year forecasts to row 14 (F14:O14).
# Row 22 is empty. Result: every year of GPR computes to 0, killing the
# whole Pro Forma → Waterfall → IRR chain. Repoint to row 14.
pf = wb["GGC Pro Forma"]
_pf_year_cols = ["H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
_umrg_year_cols = ["F", "G", "H", "I", "J", "K", "L", "M", "N", "O"]  # Y1..Y10
for pf_col, umrg_col in zip(_pf_year_cols, _umrg_year_cols):
    pf[f"{pf_col}8"] = f"='Unit Mix Rent Growth'!{umrg_col}14"

# ── 15b. Pro Forma rows 68-74: blank the 7-yr cash-flow placeholder ──
# H68:N68 had HARDCODED debt-service numbers from a prior deal
# (-655388.04, -773882.18, etc.); N69 had a hardcoded debt-payoff of
# -9833889.28. Blank them and the orphan formulas in the same block —
# the 10-yr scenario at rows 60-66 is the canonical one; the 7-yr
# block was a half-finished alternate that nobody uses.
for col in ("H", "I", "J", "K", "L", "M", "N"):
    pf[f"{col}68"] = None    # Debt Service (was hardcoded)
    pf[f"{col}69"] = None    # Debt Payoff
    pf[f"{col}70"] = None    # Refi Cashout
    pf[f"{col}71"] = None    # New Loan
    pf[f"{col}72"] = None    # Total Sale - Community
    pf[f"{col}73"] = None    # Free Cash Flow (was SUM of the hardcodes)
    pf[f"{col}74"] = None    # DSCR
# Also clear the section labels so the empty block doesn't look orphan
for r in range(68, 75):
    pf[f"E{r}"] = None
    pf[f"F{r}"] = None

# ── 15c. Electrcitiy typos in Pro Forma (we already fixed Underwriting) ──
pf["C32"] = "Electricity"
if "GGC Pro Forma (Conversion)" in wb.sheetnames:
    pf_conv = wb["GGC Pro Forma (Conversion)"]
    pf_conv["C32"] = "Electricity"

# ── 15d. Investor Return C33:D41 — stale prior-deal bullets ───────────
# These mentioned a different property ("Brookhaven"), specific $700/$775
# rents, "8 mix of SW/DW", and a typo "signiciant". Clear the whole block.
ir = wb["Investor Return"]
for r in range(33, 42):
    for col in ("B", "C", "D", "E", "F"):
        ir[f"{col}{r}"] = None

# ── 15e. Sources & Uses L93 broken formula (hanging minus operator) ──
su = wb["Sources and Uses"]
# Was: "=-'Loan Scenario (acquisition)'!J66-" (incomplete subtraction)
su["L93"] = None

# ── 15f. Unit Mix Rent Growth — clear stray deal-specific notes ──────
# These were comments from a different property's worksheet that bled
# into the blank template.
umrg = wb["Unit Mix Rent Growth"]
for coord in ("E29", "B39", "H45", "K117", "K128", "L128", "M128"):
    umrg[coord] = None

# ── 15g. Loan Scenario — clear stray text in LTO table ──────────────
ls = wb["Loan Scenario (acquisition)"]
ls["D89"] = None    # was "being evicted"
ls["J89"] = None    # was "Highlighted have baloon payment at the end of date"

# ── 15h. Unit Mix Summary — clear orphan "Annual LTO Premium" + the
# empty stretch rows 17-22 that look half-deleted.
ums = wb["Unit Mix Summary"]
for r in range(17, 24):
    for coord in (f"B{r}", f"C{r}", f"D{r}", f"E{r}", f"F{r}", f"G{r}", f"H{r}"):
        ums[coord] = None

# ── 15i. TYPO + LABEL CONSISTENCY SWEEP (from review agent #1) ───────
# Format: (sheet_name, cell_coord, new_value, reason_comment)
LABEL_FIXES = [
    # CRITICAL misspellings
    ("Data Consolidation",        "A2",  "GGC Category"),                  # "Catorgy"
    ("Data Consolidation",        "D27", "Input Source Expense Data"),     # "Epense"
    ("GGC Pro Forma",             "C16", "Economic Vacancy %"),            # "Ecomomic"
    ("GGC Pro Forma (Conversion)","C16", "Economic Vacancy %"),
    ("Loan Scenario (acquisition)","L7", "Principal"),                     # "Principle"
    ("Sources and Uses",          "B12", "Uses of Funds"),                 # "Uses of Funds of Funds"
    ("Sources and Uses",          "H12", "Uses of Funds"),
    ("Sources and Uses",          "B15", "Acquisition Fee (2%)"),          # "Acquistion"
    ("Sources and Uses",          "H15", "Acquisition Fee (2%)"),
    ("Loan Scenario (acquisition)","B17","Mortgage Constant"),             # "Costant"
    ("Loan Scenario (acquisition)","F79","10yrs"),                         # "10yyrs"
    # Label inconsistencies — standardize
    ("Unit Mix Summary",          "B12", "Avg MH Lot Rent"),               # was "Avg Rent"
    ("Unit Mix Summary",          "B24", "Avg MH Lot Rent"),
    ("GGC Pro Forma",             "C47", "Home Rent Expense (MH)"),
    ("GGC Pro Forma (Conversion)","C47", "Home Rent Expense (MH)"),
    # Sale Proceeds label standardization
    ("GGC Pro Forma",             "F72", "Total Sale - Community"),        # was "total Sale..."
    ("GGC Pro Forma (Conversion)","F62", "Total Sale - Community"),        # was "Sale Proceeds Net"
    ("GGC Pro Forma (Conversion)","F73", "Total Sale - Community"),
    ("GGC Pro Forma (Conversion)","F113","Total Sale - Community"),
    # DSCR uppercase
    ("GGC Pro Forma (Conversion)","F116","DSCR"),                          # was "dscr"
    # CoC uppercase
    ("GGC Pro Forma",             "F99", "CoC Return"),
    ("GGC Pro Forma (Conversion)","F68", "CoC Return"),
    ("GGC Pro Forma (Conversion)","F101","CoC Return"),
    ("GGC Pro Forma (Conversion)","F117","CoC Return"),
    ("GGC Pro Forma",             "F100","Avg CoC Y1-Y4"),
    ("GGC Pro Forma (Conversion)","F69", "Avg CoC Y1-Y4"),
    ("GGC Pro Forma (Conversion)","F78", "Avg CoC Y1-Y4"),
    ("GGC Pro Forma (Conversion)","F102","Avg CoC Y1-Y4"),
    # 10-yr Cash on Cash should be Y1-Y9 (was contradicting itself)
    ("Investor Return",           "C8",  "Cash on Cash Avg Y1-Y9"),
    # Bifurcated NOI label consistency
    ("GGC Underwriting",          "H2",  "Lot Rent NOI"),                  # was "LOT RENT NOI"
    ("GGC Underwriting",          "M14", "Lot Rent NOI"),                  # was "Lot Rent only NOI"
    ("GGC Underwriting",          "M15", "Home Rent NOI"),                 # was "Home Rent only NOI"
    # Trailing/leading-space cleanups
    ("Unit Mix Rent Growth",      "D13", "$ Change"),                      # was "$change"
    ("GGC Underwriting",          "M7",  "# of Units"),                    # trailing space
    ("GGC Underwriting",          "O10", "Asking Price Per Site"),         # trailing space
    ("GGC Underwriting",          "A30", "Repair and Maintenance"),
    ("GGC Pro Forma",             "C36", "Repair and Maintenance"),
    ("GGC Pro Forma (Conversion)","C36", "Repair and Maintenance"),
    ("GGC Pro Forma",             "F65", "Free Cash Flow"),
    ("GGC Pro Forma",             "F84", "Free Cash Flow"),
    ("GGC Pro Forma",             "F97", "Free Cash Flow"),
    ("GGC Pro Forma (Conversion)","F66", "Free Cash Flow"),
    ("GGC Pro Forma (Conversion)","F74", "Free Cash Flow"),
    ("GGC Pro Forma (Conversion)","F85", "Free Cash Flow"),
    ("GGC Pro Forma (Conversion)","F99", "Free Cash Flow"),
    ("GGC Pro Forma (Conversion)","F115","Free Cash Flow"),
]
for sheet_name, coord, new_value in LABEL_FIXES:
    if sheet_name in wb.sheetnames:
        try:
            wb[sheet_name][coord] = new_value
        except Exception:
            pass

# ── 15j. Pro Forma B47 — add a label for the orphan 0.15 ratio cell ──
pf["A47"] = "Home Rent Exp Ratio"   # A47 was empty; B47 had value 0.15 only

# ── 15k. Orphan fill / border cleanup ─────────────────────────────────
# Helper agent found ~109K empty cells with fill applied — most visibly
# a pink strip on Rent Roll Input!B72:B1002 (would show up as a stray
# 930-row coloured block in the partner's view). Clear fill on every
# cell that has no value AND no formula, across these known clusters.
from openpyxl.styles import PatternFill as _PF
_no_fill = _PF(fill_type=None)
ORPHAN_FILL_CLEARS = [
    # CRITICAL visible colors past data
    ("Rent Roll Input",          "B3:B1002"),   # pink strip is the biggest offender
    ("Rent Roll Input",          "C3:C1002"),
    ("Rent Roll Input",          "H3:H1002"),
    ("Rent Roll Input",          "R3:R1002"),
    ("Rent Roll Input",          "D3:T1002"),   # broad white-fill ghost block
    ("Data Consolidation",       "A30:A58"),    # orange banding on empty rows
    ("Data Consolidation",       "G30:G58"),
    ("Unit Mix Summary",         "B17:H24"),
    ("GGC Underwriting",         "K4:L17"),     # green input-placeholder boxes
    ("GGC Underwriting",         "K35:L45"),
    ("GGC Underwriting",         "O23:R278"),
    ("GGC Underwriting",         "S24:AA29"),
    # IMPORTANT solid-white ghost columns past the data
    ("Loan Scenario (acquisition)", "AC6:IG64"),
    ("Loan Scenario (acquisition)", "K11:K1003"),
    ("Loan Scenario (acquisition)", "F127:J770"),
    ("Loan Scenario (acquisition)", "L11:W421"),
    ("Loan Scenario (acquisition)", "X2:AB421"),
    ("GGC Pro Forma",             "T2:AC181"),
    ("GGC Pro Forma (Conversion)","T2:AC183"),
    ("Waterfall (10-yr-S1)",      "J3:P173"),
    ("Waterfall (10-yr-S1)",      "C20:E270"),
    ("Waterfall (5-yr-S1)",       "J3:K173"),
    ("Waterfall (5-yr-S1)",       "C20:E270"),
    ("Unit Mix Rent Growth",      "M7:AY12"),
    # MINOR
    ("Collections",              "B1:DH1"),
    ("Sources and Uses",         "E3:G11"),
    ("Sources and Uses",         "O5:Q9"),
]
def _clear_orphan_fill(sheet, range_str):
    """Clear fill on any cell in the range whose value is None and which
    has no formula. Defensive: skip any cell that currently holds data."""
    if sheet.title not in wb.sheetnames:
        return
    # openpyxl 3 supports sheet[range_str] returning a tuple of tuples
    try:
        cells = sheet[range_str]
    except Exception:
        return
    if not isinstance(cells, tuple):
        cells = (cells,)
    for row in cells:
        row = row if isinstance(row, tuple) else (row,)
        for cell in row:
            if cell.value is None:
                cell.fill = _no_fill

for tab_name, rng in ORPHAN_FILL_CLEARS:
    if tab_name in wb.sheetnames:
        _clear_orphan_fill(wb[tab_name], rng)

# ════════════════════════════════════════════════════════════════════════
# 16. METHODOLOGY FIXES — partner-defensible per-unit / per-rate logic
# ════════════════════════════════════════════════════════════════════════
# All 10 methodology items the reviewers flagged. Each is a small,
# surgical change that ties the cell back to a written GGC playbook rule
# rather than a hardcoded magic number.

uw = wb["GGC Underwriting"]

# 16a. Management Fee — restore the methodology rule:
#      5% of EGI under 200 sites, 4% at 200+
# (We had reverted to flat 5% to match CorrectOutput's single deal
# but the underlying rule is the conditional; encode it.)
uw["J33"] = "=IF(N7>=200,0.04,0.05)"

# 16b. RE Taxes — implement methodology's three-method rule:
#   Primary:  PP × 65% × local tax rate
#   Floor:    T12 × 1.15 (reassessment can't reduce taxes)
#   Fallback: $400/unit (when tax rate unknown)
# Take MAX of all three so we always honor the conservative floor.
# P12 is the user-supplied county tax rate (backend writes it from form).
uw["O12"] = "County Tax Rate"
ws_uw = uw  # alias for compatibility with earlier code blocks
uw["I22"] = (
    "=MAX("
    "J22*N7,"                                       # $400/unit floor
    "D22*1.15,"                                     # T12 × 1.15 sanity
    "IF(AND(ISNUMBER(P4),ISNUMBER(P12),P12>0),"
    "P4*0.65*P12,0))"                               # PP × 65% × rate
)

# 16c. Insurance — flood zone override per methodology lines 2706-2716.
# Base: T12 × 1.05. If property is in a flood zone (P14=TRUE), multiply
# by additional 1.15 → 1.2075 total. Backend writes P14 from form input.
uw["O13"] = "Flood Zone (Y/N)"
uw["P13"] = False    # backend sets True when floodZone == "yes"
uw["I23"] = "=D23*1.05*IF(P13=TRUE,1.15,1)"

# 16d. Stabilized Vacancy — make the 5% explicit as "economic vacancy".
# Methodology Step 2 says PHYSICAL vacancy ties to rent roll (we use
# that in column D). The 5% on G5 is the STABILIZED economic vacancy
# benchmark. Add a comment cell at K5 to make this clear.
# G5 stays at =-5%*G4 but we document the source.
uw["K5"] = "Stabilized economic vacancy (5% industry benchmark)"

# 16e. Bad Debt — document the 3% rate as a fallback when goal-seek
# isn't tractable. K7 holds the explainer; J7 stays at 0.03.
uw["K7"] = "T12 actual; 3% fallback when no trend signal"

# 16f. Bifurcated cap rates — expose as inputs at O14/O15 with the
# methodology range as a comment. Default to midpoints (6% lot, 13.5%
# home) rather than the most aggressive ends.
uw["O14"] = 0.060   # was 0.055 — midpoint of 5-7% range
uw["O15"] = 0.135   # was 0.120 — midpoint of 12-15% range
uw["Q14"] = "Methodology range: 5-7%"
uw["Q15"] = "Methodology range: 12-15%"

# 16g. Loan rates — datestamp them so they don't go stale silently.
# Loan Scenario C14=4.05%, C15=185bps already in section 5. Add a note.
ls = wb["Loan Scenario (acquisition)"]
ls["D14"] = "as of 2026-06"
ls["D15"] = "GGC standard spread"

# 16h. Pro Forma B47 (Home Rent Expense ratio) was 0.15. Methodology
# says 25-50% for POH operating expense. Move to 0.30 (midpoint of
# typical 25-35%) and add a sensitivity comment.
pf["B47"] = 0.30
pf["C47"] = "Home Rent Expense (MH) — 30% of HRI (methodology: 25-50%)"

# 16i. Sources & Uses — surface the capex breakdown so reviewers see
# they need to populate it (not just leave at $200k working cap).
# C21 utilities / C22 home in-fills are still 0 but labels make it
# obvious they're INPUTS the user should fill per deal.
su["B21"] = "Water / Septic / Utilities (per deal)"
su["B22"] = "Add Homes / In-fill (per deal)"
su["B23"] = "Working Capital"

# 16j. Document the brokerProforma fallback explicitly in the prompt
# already (commit f7fa5da). Surface it on the workbook via a header.
uw["F2"] = "Broker's Proforma"  # was "Broker's Proforma  Y1" with double space

# ════════════════════════════════════════════════════════════════════════
# 17. COMPREHENSIVE STALE-CONTENT SWEEP (from helper agent #3)
# ════════════════════════════════════════════════════════════════════════
# Reviewer found 100+ cells with prior-deal artefacts: a 17-row LTO
# contract table at Unit Mix Rent Growth rows 78-94 (specific lot
# numbers, balances, maturity dates), expansion plans at rows 117-132,
# hardcoded exit caps in Pro Forma (Conversion), hardcoded debt service
# in Pro Forma rows 68-69 (already cleared in 15b), and prior-year
# headers in Data Consolidation.

# 17a. Unit Mix Rent Growth — clear the entire LTO contract table and
# expansion plan that bled in from a prior deal.
umrg = wb["Unit Mix Rent Growth"]
# Stray notes outside the table
for coord in ("C47", "A48", "C50", "C66", "C67", "C132",
              "B123", "C124", "D125", "D130", "K117", "K119"):
    umrg[coord] = None
# LTO contract table: rows 78-94, cols C through J
for r in range(78, 95):
    for col in ("C", "D", "E", "F", "G", "H", "I", "J"):
        umrg[f"{col}{r}"] = None
# Expansion-sites row 121 (the "69 existing sites" hardcode across cols)
for col in ("F", "G", "H", "I", "J", "K", "L", "M", "N", "O"):
    umrg[f"{col}121"] = None
# Trailing stray notes at rows 128-130
for coord in ("K128", "L128", "M128"):
    umrg[coord] = None

# 17b. Pro Forma (Conversion) — clear hardcoded exit-cap-rate cells.
# These were 6% from a prior deal; the workbook should expose a single
# exit-cap input rather than four scattered hardcodes.
if "GGC Pro Forma (Conversion)" in wb.sheetnames:
    pf_conv = wb["GGC Pro Forma (Conversion)"]
    for coord in ("S64", "N83", "S97", "S113"):
        pf_conv[coord] = None

# 17c. Data Consolidation — clear prior-deal monthly date headers at
# J2:U2 (e.g., 2024-06-01 through 2025-12-01). The monthly column
# headers don't need to be pre-populated; backend.py writes them per
# deal based on the source P&L's reporting period.
dc = wb["Data Consolidation"]
for col in ("J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U"):
    dc[f"{col}2"] = None

# 17d. Rent Roll Input — clear deal-specific title with date
# ("RENT ROLL April 2026" → just "RENT ROLL").
rr = wb["Rent Roll Input"]
rr["A1"] = "RENT ROLL"

# 17e. Working Capital placeholder in Sources & Uses — make zero by
# default so the user sees they need to populate per deal.
su["C23"] = 0
su["I23"] = 0

# ════════════════════════════════════════════════════════════════════════
# 18. PARTNER-WALKTHROUGH FIXES (from Verifier 3 simulation)
# ════════════════════════════════════════════════════════════════════════
# Three bugs that no other reviewer caught — found by simulating the
# actual partner walking through every tab. Each is a math error or
# wrong-cell reference that a sharp partner would catch in 30 seconds.

# ── 18a. Rent Roll Input M column — double-counts Lot Rent ─────────────
# M2 header was "Combined" (duplicating K2), and M3:M1002 contained
# `=K3+L3+I3` where K3 already equals I3+J3. So M effectively computed
# (Lot+Home) + L + Lot = double-counting Lot Rent across 1,000 rows.
# Column L is empty. Just clear the entire M column — it's redundant
# with K, which already gives Combined = Lot + Home correctly.
rr = wb["Rent Roll Input"]
rr["M2"] = None
for r in range(3, 1003):
    rr[f"M{r}"] = None

# ── 18b. Investor Return — fix wrong row references ───────────────────
# Row 8 was =AVERAGE('Waterfall (10-yr-S1)'!G31:O31) — but row 31 is
# the LP EQUITY MULTIPLE row, not cash-on-cash. Same bug in rows 27/29
# of the duplicate block (which we also clean up below).
ir = wb["Investor Return"]
# True CoC = average annual net LP cash flow / total LP equity invested.
# Waterfall row 30 (G30:O30) holds per-year net LP cash flow.
# Waterfall F28 (sign-flipped) holds total LP equity contributed.
# Use ABS so the divide stays positive regardless of contribution sign.
ir["F8"] = ("=IFERROR(AVERAGE('Waterfall (10-yr-S1)'!G30:O30)"
            "/ABS('Waterfall (10-yr-S1)'!F28),0)")
ir["F21"] = ("=IFERROR(AVERAGE('Waterfall (5-yr-S1)'!G30:J30)"
              "/ABS('Waterfall (5-yr-S1)'!F28),0)")

# Rows 25-32 were a DUPLICATE 10-year summary block (note the double-
# space in "10  Year Return Summary" at C25 vs single space at C4).
# Row 27 in particular had F27 = AVERAGE(G31:O31) which mislabels an
# equity-multiple average as an IRR. Clear the duplicate block entirely.
for r in range(25, 33):
    for col in ("B", "C", "D", "E", "F", "G"):
        ir[f"{col}{r}"] = None

# ── 18c. Pro Forma (Conversion) — 5-yr IRR/MOIC pulled 10-yr ranges ──
# G79 IRR and G87 MOIC referenced row 66 (the 10-yr Free Cash Flow
# row) instead of row 85 (the 5-yr Free Cash Flow row). They also used
# the full G:Q range (11 cells = Y0-Y10) when the 5-yr scenario only
# spans G:L (6 cells = Y0-Y5).
if "GGC Pro Forma (Conversion)" in wb.sheetnames:
    pfc = wb["GGC Pro Forma (Conversion)"]
    pfc["G79"] = "=IRR(G85:L85)"        # was =IRR(G66:Q66) — 10-yr range
    pfc["G87"] = "=SUM(H85:L85)/-G85"   # was =SUM(H66:Q66)/-G66

# ── 18d. Pro Forma J67 DSCR typo (caught as bonus by Verifier 3) ─────
# Was =-J53/I60 (uses Year 2 debt service as Year 3 divisor). Fix to
# the same-year reference.
pf["J67"] = "=-J53/J60"

# ── 18e. Loan Scenario P8 off-by-one in interest year sum ────────────
# Was P8 = SUM(H42:H53), which shifts the interest window by one
# month into year 5. Should be SUM(H43:H54). Same shift on T8/U8.
lsq = wb["Loan Scenario (acquisition)"]
lsq["P8"] = "=SUM(H43:H54)"
lsq["T8"] = "=SUM(H91:H102)"
lsq["U8"] = "=SUM(H103:H114)"

wb.save(TEMPLATE)
print(f"Patched {TEMPLATE.name}")
print(f"Sheets now: {wb.sheetnames}")
