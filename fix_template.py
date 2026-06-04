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

# Total rows at row 151 — Underwriting!G13 (Stabilized Home Rent Income)
# reads J151 × 12 × 95% as the stabilized monthly POH rent total. Without
# this SUM the Home Rent Income line resolves to zero.
rr["B151"] = "Totals"
rr["I151"] = "=SUM(I3:I150)"     # Total monthly Lot Rent across active rows
rr["J151"] = "=SUM(J3:J150)"     # Total monthly Home Rent
rr["K151"] = "=SUM(K3:K150)"     # Total monthly Combined

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

# 12g. Rebuild Unit Mix Rent Growth rows 10-17 to mirror CorrectOutput
# EXACTLY. The correct layout drives Underwriting!G4 (stabilized GPR)
# which reads H14 = monthly weighted rent (Year 3) × total units × 12.
# Earlier rebuild used a different schema; replace it.

# Clear rows 9-22 first so the previous attempt's cells don't shadow.
for row in umrg.iter_rows(min_row=9, max_row=22, max_col=12):
    for cell in row:
        cell.value = None

# Per-type rent growth rows (10 = TOH MH, 11 = POH-Infilled).
# Columns: B=weight, C=unit-type label, D=units, E=current monthly rent,
# F-K = years 1-6 of compounded rent growth.
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

# Row 12 = weighted-average roll-up
umrg["B12"] = "=SUM(B10:B11)"
umrg["C12"] = "Total Weighted Average"
umrg["D12"] = "=SUM(D10:D11)"
for col in ("E", "F", "G", "H", "I", "J", "K"):
    umrg[f"{col}12"] = f"=SUMPRODUCT($B$10:$B$11,{col}10:{col}11)"

# Row 13 = $ change vs prior year
umrg["D13"] = "$change"
for prev, curr in (("E", "F"), ("F", "G"), ("G", "H"), ("H", "I"), ("I", "J"), ("J", "K")):
    umrg[f"{curr}13"] = f"={curr}12-{prev}12"

# Row 14 = annual GPR per year = weighted-avg monthly × total units × 12
# This is the cell Underwriting!G4 references for stabilized GPR.
umrg["D14"] = "GPR"
for col in ("E", "F", "G", "H", "I", "J", "K"):
    umrg[f"{col}14"] = f"={col}12*$D$12*12"

# Vacancy schedule (rows 15-17)
umrg["D15"] = "Vacant Lots"
umrg["E15"] = "='Unit Mix Summary'!D4"
umrg["C15"] = 0  # stabilization step target (0 by default)
umrg["F15"] = "=E15-C15"
umrg["D16"] = "Vacant Homes"
umrg["E16"] = "='Unit Mix Summary'!D5"
umrg["C16"] = 0
umrg["F16"] = "=E16-$C$16"
umrg["D17"] = "Vacancy"
for col in ("E", "F", "G", "H", "I", "J", "K"):
    umrg[f"{col}17"] = f"=SUM({col}15:{col}16)/$D$12"

# ════════════════════════════════════════════════════════════════════════
# 13. EXACT-MATCH ALIGNMENT WITH CORRECTOUTPUT.XLSX (Underwriting tab)
# ════════════════════════════════════════════════════════════════════════
# Earlier rounds added "improvements" (MAX-with-reassessment on I22,
# conditional mgmt fee on J33, County Tax Rate cell at P12) that diverge
# from the gold-standard CorrectOutput layout. User wants exact match —
# revert those and restore the original formulas + cell labels.
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

# G13 (Stabilized Home Rent Income) — correct pulls from the rent roll
# summary row at J151 with a 5% vacancy haircut.
ws["G13"] = "='Rent Roll Input'!J151*12*95%"

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
wb.calculation.fullCalcOnLoad = True

wb.save(TEMPLATE)
print(f"Patched {TEMPLATE.name}")
print(f"Sheets now: {wb.sheetnames}")
