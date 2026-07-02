"""
Phase 07 — Stabilized operating model (monthly, full horizon).

Builds the monthly revenue and expense proforma with per-line growth, the
occupancy-driven variable-expense ramp and the elevated lease-up-expense uplift,
NOI before and after property tax, below-NOI items, the forward-looking 12-month
NOI aggregates (stabilized & Year-1), and the model-calculated operating reserve
sized to the peak cumulative lease-up cash shortfall.
Property tax here uses the value/income method; phase 08 adds the reassessment
and disposition-period detail.
"""

from common import (write_title, write_section, write_subhead, label, put_formula,
                    set_col_widths, FMT_USD0, FMT_PCT2, FMT_PCT1, FMT_NUM2,
                    FILL_TOTAL, N_MONTHS)
from layout import (SH_OPS, SH_LEASE, SH_TIME, SH_PERM, set_row, r, period_header,
                    date_header, grid_row, mref, FIRST_COL_LETTER, LAST_COL_LETTER)

YIDX = "INT(({p}-1)/12)"


def gf(rate, p):
    # Monthly compounding anchored at the Growth Start month (overridable date).
    return f"(1+{rate})^(MAX(0,{p}-GrowthStartMonth)/12)"


def build(wb, reg):
    ws = wb[SH_OPS]
    set_col_widths(ws, {"A": 32, "B": 12, "C": 16})
    ws.sheet_view.showGridLines = False
    write_title(ws, 1, "Stabilized Operating Model — Monthly Proforma")

    row = 3
    row = period_header(ws, row)
    row = date_header(ws, row)

    active = "IF({p}>=DeliverStart,{body},0)"
    occ = lambda p: mref(SH_LEASE, "occ", p)
    leased = lambda p: mref(SH_LEASE, "leased", p)

    # ---- Growth & factor helper rows -------------------------------------
    write_section(ws, row, "Growth & Lease-Up Factors", span=3)
    row += 1
    row = grid_row(ws, row, "Growth Factor — Rent",
                   lambda p, C, Cprev: gf("GrowthRent", p), fmt=FMT_NUM2,
                   total=None, key="gf_rent", sheet=SH_OPS)
    row = grid_row(ws, row, "Growth Factor — Low-Income Rent",
                   lambda p, C, Cprev: gf("GrowthLowIncRent", p), fmt=FMT_NUM2,
                   total=None, key="gf_low", sheet=SH_OPS)
    row = grid_row(ws, row, "Growth Factor — Expense",
                   lambda p, C, Cprev: gf("GrowthExpense", p), fmt=FMT_NUM2,
                   total=None, key="gf_exp", sheet=SH_OPS)
    row = grid_row(ws, row, "Growth Factor — Utility",
                   lambda p, C, Cprev: gf("GrowthUtility", p), fmt=FMT_NUM2,
                   total=None, key="gf_util", sheet=SH_OPS)
    row = grid_row(ws, row, "Growth Factor — Other Income",
                   lambda p, C, Cprev: gf("GrowthOther", p), fmt=FMT_NUM2,
                   total=None, key="gf_other", sheet=SH_OPS)
    row = grid_row(ws, row, "Var-Expense Ramp Factor",
                   lambda p, C, Cprev: (f"IF({p}<DeliverStart,LeaseUpRampPct,"
                                        f"MIN(1,LeaseUpRampPct+(1-LeaseUpRampPct)*{occ(p)}/StabilizedOccupancy))"),
                   fmt=FMT_NUM2, total=None, key="ramp", sheet=SH_OPS)
    row = grid_row(ws, row, "Elevated-Expense Factor",
                   lambda p, C, Cprev: (f"1+LeaseUpElevationPct*MAX(0,1-{occ(p)}/StabilizedOccupancy)"),
                   fmt=FMT_NUM2, total=None, key="elev", sheet=SH_OPS)

    Rgfr = r(SH_OPS, "gf_rent"); Rgfl = r(SH_OPS, "gf_low"); Rgfe = r(SH_OPS, "gf_exp")
    Rgfu = r(SH_OPS, "gf_util"); Rgfo = r(SH_OPS, "gf_other")
    Rramp = r(SH_OPS, "ramp"); Relev = r(SH_OPS, "elev")

    # ---- Revenue ----------------------------------------------------------
    write_section(ws, row, "Revenue", span=3); row += 1
    def A(body):  # active guard
        return f"IF({{p}}>=DeliverStart,{body},0)"

    row = grid_row(ws, row, "Gross Potential Rent",
                   lambda p, C, Cprev: A(f"BaseMarketRentMo*{C}${Rgfr}+BaseLowRentMo*{C}${Rgfl}").format(p=p),
                   fmt=FMT_USD0, key="gpr", sheet=SH_OPS)
    Rgpr = r(SH_OPS, "gpr")
    row = grid_row(ws, row, "Vacancy Loss",
                   lambda p, C, Cprev: f"-{C}${Rgpr}*(1-{occ(p)})",
                   fmt=FMT_USD0, key="vacancy", sheet=SH_OPS)
    row = grid_row(ws, row, "Gain/Loss to Lease",
                   lambda p, C, Cprev: f"-{C}${Rgpr}*LossToLeasePct",
                   fmt=FMT_USD0, key="ltl", sheet=SH_OPS)
    row = grid_row(ws, row, "Bad Debt",
                   lambda p, C, Cprev: f"-{C}${Rgpr}*BadDebtPct",
                   fmt=FMT_USD0, key="baddebt", sheet=SH_OPS)
    row = grid_row(ws, row, "Concessions",
                   lambda p, C, Cprev: f"-{C}${Rgpr}*ConcessionsPct",
                   fmt=FMT_USD0, key="concessions", sheet=SH_OPS)
    row = grid_row(ws, row, "Parking & Storage",
                   lambda p, C, Cprev: A(f"ParkingPerUnitMo*{leased(p)}*{C}${Rgfo}").format(p=p),
                   fmt=FMT_USD0, key="parking", sheet=SH_OPS)
    row = grid_row(ws, row, "Utility Reimbursement",
                   lambda p, C, Cprev: A(f"UtilReimbPerUnitMo*{leased(p)}*{C}${Rgfo}").format(p=p),
                   fmt=FMT_USD0, key="utilreimb", sheet=SH_OPS)
    row = grid_row(ws, row, "Other Income",
                   lambda p, C, Cprev: A(f"OtherIncPerUnitMo*{leased(p)}*{C}${Rgfo}").format(p=p),
                   fmt=FMT_USD0, key="otherinc", sheet=SH_OPS)
    row = grid_row(ws, row, "Retail Income",
                   lambda p, C, Cprev: A(f"RetailRentMo*{C}${Rgfo}*MIN(1,{occ(p)}/StabilizedOccupancy)").format(p=p),
                   fmt=FMT_USD0, key="retail", sheet=SH_OPS)
    Rvac = r(SH_OPS, "vacancy"); Rltl = r(SH_OPS, "ltl"); Rbd = r(SH_OPS, "baddebt")
    Rcon = r(SH_OPS, "concessions"); Rpk = r(SH_OPS, "parking")
    Rur = r(SH_OPS, "utilreimb"); Roi = r(SH_OPS, "otherinc"); Rret = r(SH_OPS, "retail")
    row = grid_row(ws, row, "Effective Gross Income",
                   lambda p, C, Cprev: (f"{C}${Rgpr}+{C}${Rvac}+{C}${Rltl}+{C}${Rbd}+{C}${Rcon}"
                                        f"+{C}${Rpk}+{C}${Rur}+{C}${Roi}+{C}${Rret}"),
                   fmt=FMT_USD0, bold=True, key="egi", sheet=SH_OPS)
    Regi = r(SH_OPS, "egi")

    # ---- Operating expenses ----------------------------------------------
    write_section(ws, row, "Operating Expenses", span=3); row += 1
    mo = "*TotalUnits/12"
    def ex(perunit, gfrow, factorrow=None):
        def f(p, C, Cprev):
            fac = f"*{C}${factorrow}" if factorrow else ""
            return A(f"{perunit}{mo}*{C}${gfrow}{fac}").format(p=p)
        return f
    row = grid_row(ws, row, "Insurance", ex("ExpInsurancePerUnit", Rgfe),
                   fmt=FMT_USD0, key="insurance", sheet=SH_OPS)
    row = grid_row(ws, row, "Admin & Legal", ex("ExpAdminPerUnit", Rgfe),
                   fmt=FMT_USD0, key="admin", sheet=SH_OPS)
    row = grid_row(ws, row, "Utilities (variable)", ex("ExpUtilitiesPerUnit", Rgfu, Rramp),
                   fmt=FMT_USD0, key="utilities", sheet=SH_OPS)
    row = grid_row(ws, row, "Repairs & Maintenance (variable)", ex("ExpRMPerUnit", Rgfe, Rramp),
                   fmt=FMT_USD0, key="rm", sheet=SH_OPS)
    row = grid_row(ws, row, "Turnover (variable)", ex("ExpTurnoverPerUnit", Rgfe, Rramp),
                   fmt=FMT_USD0, key="turnover", sheet=SH_OPS)
    row = grid_row(ws, row, "Contract Services (variable)", ex("ExpContractPerUnit", Rgfe, Rramp),
                   fmt=FMT_USD0, key="contract", sheet=SH_OPS)
    row = grid_row(ws, row, "Payroll (elevated)", ex("ExpPayrollPerUnit", Rgfe, Relev),
                   fmt=FMT_USD0, key="payroll", sheet=SH_OPS)
    row = grid_row(ws, row, "Advertising (elevated)", ex("ExpAdvertisingPerUnit", Rgfe, Relev),
                   fmt=FMT_USD0, key="advertising", sheet=SH_OPS)
    row = grid_row(ws, row, "Management Fee",
                   lambda p, C, Cprev: A(f"MgmtFeePct*{C}${Regi}").format(p=p),
                   fmt=FMT_USD0, key="mgmt", sheet=SH_OPS)
    Rins = r(SH_OPS, "insurance"); Radm = r(SH_OPS, "admin"); Rut = r(SH_OPS, "utilities")
    Rrm = r(SH_OPS, "rm"); Rturn = r(SH_OPS, "turnover"); Rcs = r(SH_OPS, "contract")
    Rpay = r(SH_OPS, "payroll"); Radv = r(SH_OPS, "advertising"); Rmg = r(SH_OPS, "mgmt")
    row = grid_row(ws, row, "OpEx before Property Tax",
                   lambda p, C, Cprev: (f"{C}${Rins}+{C}${Radm}+{C}${Rut}+{C}${Rrm}+{C}${Rturn}"
                                        f"+{C}${Rcs}+{C}${Rpay}+{C}${Radv}+{C}${Rmg}"),
                   fmt=FMT_USD0, bold=True, key="opex_bt", sheet=SH_OPS)
    Ropexbt = r(SH_OPS, "opex_bt")

    # NOI before property tax
    row = grid_row(ws, row, "NOI before Property Tax",
                   lambda p, C, Cprev: f"{C}${Regi}-{C}${Ropexbt}",
                   fmt=FMT_USD0, bold=True, key="noi_bt", sheet=SH_OPS)
    Rnoibt = r(SH_OPS, "noi_bt")

    # Property tax (value approach default; income approach non-circular via NOI-bt)
    def f_tax(p, C, Cprev):
        # Operating property tax applies only from stabilization; before that,
        # tax is capitalized as the construction-period tax (a dev cost).
        yrs = f"MAX(0,({p}-StabMonth)/12)"
        val_app = f"TaxYear1*(1+GrowthTax)^({yrs})/12"
        eff = "TaxLevyPct*TaxAssessmentPct*TaxValueAdjFactor"
        inc_app = f"MAX(0,{C}${Rnoibt}*12)/EntryCapRate*({eff})/12"
        return f"IF({p}<StabMonth,0,IF(TaxMethodValue>=1,{val_app},{inc_app}))"
    row = grid_row(ws, row, "Property Tax", f_tax, fmt=FMT_USD0, key="tax", sheet=SH_OPS)
    Rtax = r(SH_OPS, "tax")

    row = grid_row(ws, row, "Total Operating Expenses",
                   lambda p, C, Cprev: f"{C}${Ropexbt}+{C}${Rtax}",
                   fmt=FMT_USD0, bold=True, key="opex_total", sheet=SH_OPS)
    Ropext = r(SH_OPS, "opex_total")

    row = grid_row(ws, row, "Net Operating Income",
                   lambda p, C, Cprev: f"{C}${Rnoibt}-{C}${Rtax}",
                   fmt=FMT_USD0, bold=True, key="noi", sheet=SH_OPS)
    Rnoi = r(SH_OPS, "noi")

    # Below-NOI items
    write_section(ws, row, "Below-NOI Items", span=3); row += 1
    row = grid_row(ws, row, "Capital Reserves + Tax Prep + Asset Mgmt Fee",
                   lambda p, C, Cprev: A(f"CapReservePerUnit{mo}+TaxPrepPerUnit{mo}+AssetMgmtFeePct*{C}${Regi}").format(p=p),
                   fmt=FMT_USD0, key="below", sheet=SH_OPS)
    Rbelow = r(SH_OPS, "below")

    # ---- Cash flow rows (op CF pre-debt, and to-equity with perm DS) ------
    write_section(ws, row, "Operating Cash Flow", span=3); row += 1
    row = grid_row(ws, row, "Operating CF before Debt (property)",
                   lambda p, C, Cprev: f"{C}${Rnoi}-{C}${Rbelow}",
                   fmt=FMT_USD0, key="op_cf_pre", sheet=SH_OPS)
    Rpre = r(SH_OPS, "op_cf_pre")
    # cumulative op CF during lease-up window (for reserve sizing)
    def f_cum(p, C, Cprev):
        prev = "0" if Cprev is None else f"{Cprev}${r_cum}"
        return f"IF({p}<DeliverStart,0,{prev}+{C}${Rpre})"
    r_cum = row  # this row
    row = grid_row(ws, row, "Cumulative Op CF (lease-up)", f_cum, fmt=FMT_USD0,
                   total=None, key="cum_pre", sheet=SH_OPS)
    Rcum = r(SH_OPS, "cum_pre")

    # Levered operating CF (net of perm debt service) is assembled in the
    # CashFlow phase, where perm_ds is available (avoids a build-order cycle).
    row = grid_row(ws, row, "Unlevered Operating CF (property)",
                   lambda p, C, Cprev: f"IF({mref(SH_TIME,'OpsActive',p)}=1,{C}${Rpre},0)",
                   fmt=FMT_USD0, key="unlev_ops", sheet=SH_OPS)

    # ---- Aggregates & operating reserve ----------------------------------
    row += 1
    row = write_subhead(ws, row, "Operating Aggregates & Reserve", span=3)
    noirng = f"{FIRST_COL_LETTER}{Rnoi}:{LAST_COL_LETTER}{Rnoi}"
    cumrng = f"{FIRST_COL_LETTER}{Rcum}:{LAST_COL_LETTER}{Rcum}"
    opextrng = f"{FIRST_COL_LETTER}{Ropext}:{LAST_COL_LETTER}{Ropext}"

    def agg(lbl, formula, name, fmt=FMT_USD0):
        label(ws, row_ref[0], lbl, col=1)
        c = put_formula(ws, row_ref[0], 2, formula, fmt=fmt, link=True, bold=True)
        c.fill = FILL_TOTAL
        reg.add(name, SH_OPS, f"B{row_ref[0]}")
        set_row(SH_OPS, name, row_ref[0])
        row_ref[0] += 1
    row_ref = [row]
    noibtrng = f"{FIRST_COL_LETTER}{Rnoibt}:{LAST_COL_LETTER}{Rnoibt}"
    agg("Stabilized NOI before Tax (fwd 12mo)",
        f"SUM(INDEX({noibtrng},StabMonth):INDEX({noibtrng},StabMonth+11))", "StabNOIbt")
    agg("Stabilized NOI (fwd 12mo)",
        f"SUM(INDEX({noirng},StabMonth):INDEX({noirng},StabMonth+11))", "StabNOI")
    agg("Year-1 NOI (fwd 12mo from delivery)",
        f"SUM(INDEX({noirng},DeliverStart):INDEX({noirng},DeliverStart+11))", "Year1NOI")
    agg("Peak Cumulative Lease-Up Shortfall",
        f"MIN(0,MIN(INDEX({cumrng},DeliverStart):INDEX({cumrng},StabMonth)))", "PeakShort")
    agg("Operating Reserve (model-calculated)",
        f"MAX(MAX(0,-PeakShort)*OpReserveCoverage,OpReserveMinMonths*INDEX({opextrng},StabMonth))",
        "OpReserve")
    ws.freeze_panes = "D4"
