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
# resolves to 0. We point at the rebuilt Unit Mix Summary (Total Units
# moves to E8 in the new 4-row layout).
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

# Headers (row 3)
ums["B2"] = "MH/RV Rent Roll"
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
ums["B9"]  = "Total MH Sites (TOH + POH)"
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
ums["B11"] = "Annual GPR (MH Lot Rent)"
ums["C11"] = "=H9*12"
ums["B12"] = "Avg MH Lot Rent"
ums["C12"] = "=IFERROR(H4/E4,0)"
ums["B13"] = "Occupancy %"
ums["C13"] = "=IFERROR(C8/E8,0)"
ums["B14"] = "POH %"
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

# ════════════════════════════════════════════════════════════════════════
# 4. UNIT MIX RENT GROWTH — replace 6 #REF! rows with 2 real rows
# ════════════════════════════════════════════════════════════════════════
umrg = wb["Unit Mix Rent Growth"]

# Clear the broken block (rows 13-22) before rewriting
for row in umrg.iter_rows(min_row=13, max_row=22, max_col=12):
    for cell in row:
        cell.value = None

umrg["B13"] = "Unit Type"
umrg["C13"] = "# of Units"
umrg["D13"] = "Avg Monthly Rent"
umrg["E13"] = "Year 1"
umrg["F13"] = "Year 2"
umrg["G13"] = "Year 3"
umrg["H13"] = "Year 4"
umrg["I13"] = "Year 5"

# Row 14: MH Lot Rent (pulls TOH MH counts/rent from Unit Mix Summary)
umrg["B14"] = "MH Lot Rent (TOH)"
umrg["C14"] = "='Unit Mix Summary'!E4"
umrg["D14"] = "='Unit Mix Summary'!C12"
umrg["E14"] = "=D14*(1+B$5)"
umrg["F14"] = "=E14*(1+C$5)"
umrg["G14"] = "=F14*(1+D$5)"
umrg["H14"] = "=G14*(1+E$5)"
umrg["I14"] = "=H14*(1+F$5)"

# Row 15: RV Lot Rent
umrg["B15"] = "Long Term RV Rent"
umrg["C15"] = "='Unit Mix Summary'!E6"
umrg["D15"] = "='Unit Mix Summary'!C16"
umrg["E15"] = "=D15*(1+B$6)"
umrg["F15"] = "=E15*(1+C$6)"
umrg["G15"] = "=F15*(1+D$6)"
umrg["H15"] = "=G15*(1+E$6)"
umrg["I15"] = "=H15*(1+F$6)"

# Weighted-average roll-up (used as stabilized GPR if desired)
umrg["B16"] = "Annual Stabilized GPR"
umrg["C16"] = "=C14+C15"
umrg["G16"] = "=(G14*C14+G15*C15)*12"  # Year 3 stabilized total

# Wire stabilized GPR on Underwriting to this cell
ws["G4"] = "='Unit Mix Rent Growth'!G16"

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

# 11e. Unit Mix Rent Growth row 22 was the year-by-year GPR projection
# but only M22:O22 (Y8-Y10) were set. Fill in F22:L22 (Y1-Y7) as a
# product of unit counts x year-N monthly rent x 12, and fix the
# weighted-average roll-up.
umrg = wb["Unit Mix Rent Growth"]
# Annual GPR per year = (MH count * MH year-N rent + RV count * RV year-N rent) * 12
gpr_year_formulas = {
    "F22": "=(E14*C14+E15*C15)*12",  # Y1
    "G22": "=(F14*C14+F15*C15)*12",  # Y2
    "H22": "=(G14*C14+G15*C15)*12",  # Y3 (mirrors G16)
    "I22": "=(H14*C14+H15*C15)*12",  # Y4
    "J22": "=(I14*C14+I15*C15)*12",  # Y5
    "K22": "=(J14*C14+J15*C15)*12",  # Y6
    "L22": "=(J14*(1+G$5)*C14+J15*(1+G$6)*C15)*12",  # Y7
    "M22": "=(J14*(1+G$5)*(1+H$5)*C14+J15*(1+G$6)*(1+H$6)*C15)*12",      # Y8
    "N22": "=(J14*(1+G$5)*(1+H$5)*(1+I$5)*C14+J15*(1+G$6)*(1+H$6)*(1+I$6)*C15)*12",  # Y9
    "O22": "=(J14*(1+G$5)*(1+H$5)*(1+I$5)*(1+J$5)*C14+J15*(1+G$6)*(1+H$6)*(1+I$6)*(1+J$6)*C15)*12",  # Y10
}
umrg["B22"] = "Annual GPR (lots only, by year)"
for coord, formula in gpr_year_formulas.items():
    umrg[coord] = formula

# 11f. Zero out the broken SUMPRODUCT placeholders at M20:O20 (they
# referenced B14:B19 which are text labels). Without these, downstream
# Pro Forma cells get clean year-by-year inputs from row 22 instead.
for coord in ("M20", "N20", "O20"):
    if isinstance(umrg[coord].value, str) and "SUMPRODUCT" in (umrg[coord].value or "").upper():
        umrg[coord] = 0

# 11g. Clean orphan label refs at A5:A10 (pointed at numeric formula
# cells C14:C19). Replace with the static unit-type labels.
for r, label in [(5, "MH Lot Rent (TOH)"), (6, "Long Term RV Rent"),
                 (7, ""), (8, ""), (9, ""), (10, "")]:
    umrg[f"A{r}"] = label or None

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

# 12g. The verification agent flagged Unit Mix Rent Growth lacks the
# canonical rows 9-17 (weighted-average + vacancy schedule). Build them.
# These feed Pro Forma's vacancy + GPR projections.
umrg["B9"] = "Rent Growth Vector (annual)"
umrg["B10"] = "MH Lot Rent (TOH)"
umrg["C10"] = "='Unit Mix Summary'!E4"  # unit count
umrg["B11"] = "Long Term RV Rent"
umrg["C11"] = "='Unit Mix Summary'!E6"
umrg["B12"] = "Weighted Avg Monthly Rent"
umrg["C12"] = ("=IFERROR((C10*'Unit Mix Summary'!C12 + "
               "C11*'Unit Mix Summary'!C16)/(C10+C11),0)")
umrg["B13"] = "Monthly $ Rent Change"
umrg["C13"] = "=C12-'Unit Mix Summary'!C12"
umrg["B14"] = "Annual GPR (combined lots)"
for col_idx, year_letter in enumerate(["E", "F", "G", "H", "I", "J"], start=1):
    # Year 1-6 GPR per year from row 22 (already populated above)
    umrg[f"{year_letter}14"] = f"={year_letter}22"
umrg["B15"] = "Vacant MH Lots"
umrg["C15"] = "='Unit Mix Summary'!D4"
umrg["B16"] = "Vacant RV / Other"
umrg["C16"] = "='Unit Mix Summary'!D6+'Unit Mix Summary'!D7"
umrg["B17"] = "Physical Vacancy %"
umrg["C17"] = "=IFERROR((C15+C16)/(C10+C11+C15+C16),0)"

# ════════════════════════════════════════════════════════════════════════
# 13. FORCE FULL RECALC ON OPEN
# ════════════════════════════════════════════════════════════════════════
wb.calculation.fullCalcOnLoad = True

wb.save(TEMPLATE)
print(f"Patched {TEMPLATE.name}")
print(f"Sheets now: {wb.sheetnames}")
