"""
assumptions.py — SINGLE SOURCE OF TRUTH for every model input.

Both the Excel generator (phase_*.py) and the independent Python shadow model
(shadow.py) import their inputs from here, so no assumption is ever defined
twice.  Scalar inputs live in ``SCALARS`` (a flat name -> value dict whose keys
are the workbook named ranges).  Structured inputs (milestones, unit matrix,
budget lines) live in their own tables.

Nothing here computes model results — these are raw, hard-keyed assumptions
only.  Everything downstream is derived.
"""

# ---------------------------------------------------------------------------
# Horizon (kept consistent with common.py)
# ---------------------------------------------------------------------------
N_HOLD_MAX = 360
N_FWD = 12
N_MONTHS = N_HOLD_MAX + N_FWD

# ===========================================================================
# 1. SCALAR INPUTS  (name -> (value, number_format, unit/help note, section))
# ===========================================================================
# The name is the workbook named range.  Order within a section is preserved.

FMT = dict(
    usd="#,##0;(#,##0)", usd2="#,##0.00;(#,##0.00)", usdc="$#,##0;($#,##0)",
    pct1="0.0%", pct2="0.00%", pct3="0.000%", num0="#,##0", num1="#,##0.0",
    num2="#,##0.00", mult='0.00"x"', mo='#,##0" mo"', txt="@", date="mmm-yyyy",
    bps='#,##0" bps"',
)

# Each entry: key -> [value, fmt, note]
SECTIONS = []

def _sec(title, rows):
    SECTIONS.append((title, rows))

_sec("Property Details", [
    ("InvestmentType",  "Ground-Up Development", FMT["txt"], "Investment type"),
    ("PropertyName",    "Meridian Apartments",   FMT["txt"], "Deal name"),
    ("City",            "Austin",                FMT["txt"], ""),
    ("State",           "TX",                    FMT["txt"], ""),
    ("Market",          "Austin-Round Rock",     FMT["txt"], "MSA"),
    ("ProductType",     "Garden / Mid-Rise",     FMT["txt"], ""),
    ("Location",        "Suburban Infill",       FMT["txt"], ""),
    ("InceptionDate",   "2026-01-01",            FMT["date"], "Model period 1 = this month"),
])

_sec("Investment Details", [
    ("HoldPeriodMonths",   84,     FMT["mo"],   "Disposition month (1-360 toggle)"),
    ("EntryCapRate",       0.0475, FMT["pct2"], "Market / entry cap rate"),
    ("CapAdjLocation",     0.0000, FMT["pct2"], "Location adjustment (+/-)"),
    ("CapAdjSize",         0.0000, FMT["pct2"], "Size adjustment (+/-)"),
    ("CapAdjAging",        0.0025, FMT["pct2"], "Aging adjustment (+/-)"),
    ("ExitCapResidential", 0.0500, FMT["pct2"], "Exit cap - residential"),
    ("ExitCapRetail",      0.0650, FMT["pct2"], "Exit cap - retail"),
    ("GrowthRent",         0.0300, FMT["pct2"], "Annual market-rent growth"),
    ("GrowthLowIncRent",   0.0200, FMT["pct2"], "Annual low-income-rent growth"),
    ("GrowthExpense",      0.0300, FMT["pct2"], "Annual expense growth"),
    ("GrowthUtility",      0.0350, FMT["pct2"], "Annual utility growth"),
    ("GrowthTax",          0.0200, FMT["pct2"], "Annual tax growth"),
    ("GrowthOther",        0.0300, FMT["pct2"], "Annual other-income growth"),
])

_sec("Unit Delivery & Lease-Up", [
    ("LeaseUpIncomeOffset", 1,    FMT["num0"], "1=offset lease-up income vs equity, 0=no"),
    ("UnitsDeliveredPerMo", 30,   FMT["num0"], "Units delivered per month"),
    ("UnitsLeasedPerMo",    18,   FMT["num0"], "Units absorbed (leased) per month"),
    ("LeaseUpRampPct",      0.60, FMT["pct1"], "Variable-opex ramp: % of stabilized during lease-up"),
    ("LeaseUpElevationPct", 0.25, FMT["pct1"], "Elevated-opex uplift over stabilized during lease-up"),
    ("StabilizedOccupancy", 0.93, FMT["pct1"], "Physical occupancy at stabilization"),
])

_sec("Site Details", [
    ("LandCost",         14000000, FMT["usdc"], "Total land cost"),
    ("GrossAcres",       9.50,     FMT["num2"], "Gross site acres"),
    ("NetAcres",         8.00,     FMT["num2"], "Net developable acres"),
    ("FARRatio",         1.75,     FMT["num2"], "Floor-area ratio"),
    ("EfficiencyRatio",  0.85,     FMT["pct1"], "Net rentable / gross buildable"),
    ("ExistingStructureCost", 0,   FMT["usdc"], "Demo / existing structure"),
])

_sec("Property Taxes", [
    ("TaxLevyPct",        0.0185, FMT["pct2"], "Levy / mill rate (% of assessed)"),
    ("TaxAssessmentPct",  1.00,   FMT["pct1"], "Assessment ratio (% of value)"),
    ("TaxValueAdjFactor", 1.00,   FMT["num2"], "Value-adjustment factor"),
    ("TaxMethodValue",    1,      FMT["num0"], "Post-stab method: 1=value approach, 0=income approach"),
    ("TaxYear1",          650000, FMT["usdc"], "Year-1 stabilized tax (value approach)"),
    ("ReassessOnSale",    1,      FMT["num0"], "1=reassess buyer taxes on sale value, 0=in-place"),
    ("DispoTaxRatePct",   0.0185, FMT["pct2"], "Disposition-period effective tax rate on value"),
])

_sec("Capitalization & Funding", [
    ("FundingOrderPari",  0,      FMT["num0"], "1=pari-passu, 0=equity-first"),
    ("TargetLTC",         0.6500, FMT["pct2"], "Target construction loan-to-cost"),
    ("OpReserveCoverage", 1.10,   FMT["mult"], "Operating reserve coverage multiple on peak shortfall"),
    ("OpReserveMinMonths",3,      FMT["num0"], "Minimum months of stabilized opex floor for reserve"),
])

_sec("Construction Loan", [
    ("ConLoanTermMonths",  36,     FMT["mo"],   "Construction loan term"),
    ("ConRate",            0.0850, FMT["pct2"], "Construction interest rate (annual)"),
    ("ConActual360",       1,      FMT["num0"], "1=Actual/360, 0=30/360"),
    ("ConPointsPct",       0.0100, FMT["pct2"], "Origination points"),
    ("ConOtherCosts",      250000, FMT["usdc"], "Other construction loan costs"),
    ("ConMinDebtYield",    0.0700, FMT["pct2"], "Min stabilized debt yield (sizing)"),
    ("ConMinDSCR",         1.20,   FMT["mult"], "Min DSCR at target (sizing)"),
    ("LenderReqReserveOverride", 0, FMT["usdc"], "If >0, overrides computed interest reserve"),
])

_sec("Permanent Loan (Refinance)", [
    ("PermMaxLTV",        0.6500, FMT["pct2"], "Max permanent loan-to-value"),
    ("PermRate",          0.0625, FMT["pct2"], "Permanent interest rate (annual)"),
    ("PermTermMonths",    360,    FMT["mo"],   "Permanent loan term"),
    ("PermIOMonths",      24,     FMT["mo"],   "Permanent interest-only period"),
    ("PermAmortMonths",   360,    FMT["mo"],   "Permanent amortization period"),
    ("PermActual360",     0,      FMT["num0"], "1=Actual/360, 0=30/360"),
    ("PermPointsPct",     0.0075, FMT["pct2"], "Permanent loan points"),
    ("PermMinDebtYield",  0.0800, FMT["pct2"], "Min debt yield at refi (sizing)"),
    ("PermMinDSCR",       1.15,   FMT["mult"], "Min DSCR at refi (sizing)"),
    ("NoCashOut",         0,      FMT["num0"], "1=size perm so net cash at refi = 0"),
    ("RefiMonthOverride", 0,      FMT["num0"], "If >0, override refi month (else = con maturity)"),
])

_sec("Sponsor Fees", [
    ("DevFeePctDirect",     0.0400, FMT["pct2"], "Development fee % of direct costs"),
    ("DevFeeThirdPartySplit",0.0000,FMT["pct2"], "Portion of dev fee paid to 3rd party"),
    ("SponsorDispoFeePct",  0.0100, FMT["pct2"], "Sponsor disposition fee % of sale"),
    ("RecourseFeePct",      0.0025, FMT["pct2"], "Loan recourse / indemnity fee % of loan"),
])

_sec("Disposition & Sale", [
    ("BrokerCostPct",     0.0150, FMT["pct2"], "Broker / sale cost % of gross sale"),
    ("SalesTaxPct",       0.0000, FMT["pct2"], "Transfer / sales tax % of gross sale"),
    ("TaxProrationMonths",2,      FMT["mo"],   "Property-tax proration at sale (months)"),
])

_sec("Stabilized Revenue (other income)", [
    # Physical vacancy is derived from StabilizedOccupancy (1 - occupancy); the
    # items below are additional economic deductions and other income.
    ("BadDebtPct",        0.0050, FMT["pct2"], "Bad debt % of GPR"),
    ("ConcessionsPct",    0.0100, FMT["pct2"], "Concessions % of GPR"),
    ("LossToLeasePct",    0.0100, FMT["pct2"], "Gain/loss-to-lease % of GPR"),
    ("ModelUnits",        1,      FMT["num0"], "Down / model units (non-revenue)"),
    ("ParkingPerUnitMo",  35,     FMT["usdc"], "Parking & storage income $/unit/mo"),
    ("RetailIncomeMo",    18000,  FMT["usdc"], "Retail income $/mo (stabilized)"),
    ("UtilReimbPerUnitMo",45,     FMT["usdc"], "Utility reimbursement $/unit/mo"),
    ("OtherIncPerUnitMo", 25,     FMT["usdc"], "Other income $/unit/mo"),
])

_sec("Stabilized Operating Expenses (annual)", [
    ("ExpInsurancePerUnit",   425,  FMT["usdc"], "Insurance $/unit/yr"),
    ("ExpUtilitiesPerUnit",   900,  FMT["usdc"], "Utilities $/unit/yr (variable)"),
    ("ExpPayrollPerUnit",     1300, FMT["usdc"], "Payroll $/unit/yr (elevated in lease-up)"),
    ("ExpRMPerUnit",          650,  FMT["usdc"], "Repairs & maintenance $/unit/yr (variable)"),
    ("ExpTurnoverPerUnit",    250,  FMT["usdc"], "Turnover $/unit/yr (variable)"),
    ("ExpContractPerUnit",    500,  FMT["usdc"], "Contract services $/unit/yr (variable)"),
    ("ExpAdminPerUnit",       300,  FMT["usdc"], "Admin & legal $/unit/yr"),
    ("ExpAdvertisingPerUnit", 200,  FMT["usdc"], "Advertising $/unit/yr (elevated in lease-up)"),
    ("MgmtFeePct",            0.0300,FMT["pct2"],"Management fee % of EGI"),
    ("CapReservePerUnit",     250,  FMT["usdc"], "Capital reserves $/unit/yr (below NOI)"),
    ("TaxPrepPerUnit",        15,   FMT["usdc"], "Tax prep $/unit/yr (below NOI)"),
    ("AssetMgmtFeePct",       0.0100,FMT["pct2"],"Asset mgmt fee % of EGI (below NOI)"),
])

# Flat dict for the shadow model and for named-range registration.
SCALARS = {}
for _title, _rows in SECTIONS:
    for _k, _v, _f, _n in _rows:
        SCALARS[_k] = _v

# ===========================================================================
# 2. MILESTONE TABLE  (Milestone | Start month (1-based) | Duration months)
# ===========================================================================
# End month is derived = Start + Duration - 1.  Every downstream event keys off
# these, never a hardcoded date/month.
MILESTONES = [
    # name,                    start_month, duration_mo
    ("Open Escrow",                 1,   1),
    ("Due Diligence",               1,   3),
    ("Entitlement",                 2,   5),
    ("Pre-Construction",            5,   3),
    ("Initial Closing",             7,   1),
    ("Construction",                7,  20),
    ("1st Unit Delivery",          21,   1),
    ("Lease-Up",                   21,  12),
    ("Stabilization",              33,   1),
    ("Recapitalization/Refinance", 43,   1),   # default; engine caps at con maturity
    ("Disposition",                84,   1),   # tracks HoldPeriodMonths
]
# Named ranges emitted per milestone: MS_<slug>_Start, MS_<slug>_Dur, MS_<slug>_End
def ms_slug(name):
    keep = []
    for ch in name:
        if ch.isalnum():
            keep.append(ch)
        elif keep and keep[-1] != "_":
            keep.append("_")
    return "".join(keep).strip("_")

# ===========================================================================
# 3. UNIT MATRIX  (residential floorplans, market-rate + low-income; retail)
# ===========================================================================
# Each: plan, class(market/low), units, avg_sf, market_rent_mo, net_effective_mo
UNIT_PLANS = [
    # plan,     class,     units, avg_sf, mkt_rent, ne_rent
    ("Studio",  "Market",   40,   540,   1895,  1845),
    ("1BR-A",   "Market",   60,   720,   2245,  2195),
    ("1BR-B",   "Market",   30,   780,   2360,  2300),
    ("2BR-A",   "Market",   40,  1050,   3050,  2975),
    ("2BR-B",   "Market",   10,  1150,   3220,  3140),
    ("1BR-LI",  "Low",      12,   720,   1350,  1350),
    ("2BR-LI",  "Low",       8,  1050,   1625,  1625),
]
# Retail: suite, sf, rent_psf_yr
RETAIL_SUITES = [
    ("Retail-1", 4200, 34.0),
    ("Retail-2", 3100, 32.0),
]

# ===========================================================================
# 4. DEVELOPMENT BUDGET LINE ITEMS
# ===========================================================================
# basis: "abs" ($), "unit" ($/unit), "sf" ($/buildable NRSF)
# phase: milestone name governing timing window
# curve: Straight_Line | Bell_Curve | Development_Fee | Single_Point
# amount: interpreted per basis
# modifier: integer month offset applied to the phase window start (+/-)
BUDGET_LINES = [
    # label,                   group,   type,   basis,  amount,  phase,             curve,            modifier
    ("Land Acquisition",       "Land",  "Land", "abs",  None,    "Initial Closing",  "Single_Point",   0),
    ("Existing Structure/Demo","Land",  "Land", "abs",  None,    "Pre-Construction", "Straight_Line",  0),

    ("Due Diligence",          "Soft",  "Soft", "abs",  150000,  "Due Diligence",    "Straight_Line",  0),
    ("Entitlements",           "Soft",  "Soft", "abs",  600000,  "Entitlement",      "Straight_Line",  0),
    ("Architecture & Eng.",    "Soft",  "Soft", "sf",   6.50,    "Pre-Construction", "Straight_Line",  0),
    ("Consultants",            "Soft",  "Soft", "abs",  400000,  "Entitlement",      "Straight_Line",  0),
    ("Legal & Org.",           "Soft",  "Soft", "abs",  300000,  "Due Diligence",    "Straight_Line",  0),
    ("Marketing & Lease-Up",   "Soft",  "Soft", "unit", 1200,    "Lease-Up",         "Straight_Line",  0),
    ("Property Taxes (Dev)",   "Soft",  "Soft", "abs",  700000,  "Construction",     "Straight_Line",  0),
    ("Insurance (Dev)",        "Soft",  "Soft", "abs",  500000,  "Construction",     "Straight_Line",  0),

    ("GC Contract (Hard)",     "Hard",  "Hard", "sf",   200.0,   "Construction",     "Bell_Curve",     0),
    ("Off-Site Improvements",  "Hard",  "Hard", "abs",  1200000, "Construction",     "Straight_Line",  0),
    ("Utilities & Hookups",    "Hard",  "Hard", "abs",  900000,  "Construction",     "Straight_Line",  0),
    ("Builder's Risk",         "Hard",  "Hard", "abs",  350000,  "Construction",     "Straight_Line",  0),
    ("Hard Cost Contingency",  "Hard",  "Hard", "sf",   10.0,    "Construction",     "Bell_Curve",     0),
    ("GC Fee",                 "Hard",  "Hard", "sf",   12.0,    "Construction",     "Bell_Curve",     0),
]
# 12 user-definable "Other Soft/Hard Cost" lines (extensible: change N_USER_LINES).
N_USER_LINES = 12
USER_LINE_DEFAULTS = [
    # label, group/type, basis, amount, phase, curve, modifier  (blank amount => 0)
    ("Other Soft Cost 1 - FF&E",     "Soft", "unit", 800,   "Pre-Construction", "Straight_Line", 0),
    ("Other Soft Cost 2 - Permits",  "Soft", "abs",  250000,"Entitlement",      "Straight_Line", 0),
    ("Other Soft Cost 3 - Testing",  "Soft", "abs",  120000,"Construction",     "Straight_Line", 0),
    ("Other Soft Cost 4",            "Soft", "abs",  0,     "Construction",     "Straight_Line", 0),
    ("Other Soft Cost 5",            "Soft", "abs",  0,     "Construction",     "Straight_Line", 0),
    ("Other Soft Cost 6",            "Soft", "abs",  0,     "Construction",     "Straight_Line", 0),
    ("Other Hard Cost 1 - Landscape","Hard", "abs",  650000,"Construction",     "Bell_Curve",    0),
    ("Other Hard Cost 2 - Amenity",  "Hard", "abs",  900000,"Construction",     "Bell_Curve",    0),
    ("Other Hard Cost 3 - Signage",  "Hard", "abs",  150000,"Construction",     "Straight_Line", 0),
    ("Other Hard Cost 4",            "Hard", "abs",  0,     "Construction",     "Bell_Curve",    0),
    ("Other Hard Cost 5",            "Hard", "abs",  0,     "Construction",     "Bell_Curve",    0),
    ("Other Hard Cost 6",            "Hard", "abs",  0,     "Construction",     "Bell_Curve",    0),
]

# Financing lines (points, other loan costs, interest reserve, dev fee, op reserve)
# are computed by their respective engines, not hard-keyed here, but they occupy
# reserved budget rows so the budget foots to total capitalization.

def all_budget_lines():
    """Return standard + user lines as uniform dicts."""
    out = []
    for (label, group, typ, basis, amount, phase, curve, mod) in BUDGET_LINES:
        out.append(dict(label=label, group=group, type=typ, basis=basis,
                        amount=amount, phase=phase, curve=curve, modifier=mod,
                        user=False))
    for i in range(N_USER_LINES):
        if i < len(USER_LINE_DEFAULTS):
            label, typ, basis, amount, phase, curve, mod = USER_LINE_DEFAULTS[i]
        else:
            label, typ, basis, amount, phase, curve, mod = (
                f"User Line {i+1}", "Soft", "abs", 0, "Construction", "Straight_Line", 0)
        out.append(dict(label=label, group=typ, type=typ, basis=basis,
                        amount=amount, phase=phase, curve=curve, modifier=mod,
                        user=True))
    return out
