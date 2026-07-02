"""
shadow.py — Independent pure-Python re-computation of the entire model.

This is the AUTHORITATIVE numeric check.  It reads the same inputs as the Excel
workbook (from assumptions.py) and reproduces every mechanic with plain Python
arithmetic, month by month.  validate.py reconciles the LibreOffice-recalculated
Excel values against the arrays this module returns, to a tight tolerance.

Design mirrors the Excel exactly:
  * fully monthly over N_MONTHS (=372) columns
  * interest reserve resolved by a sequential row-by-row draw schedule (no circ.)
  * reassess-on-sale solved in closed form
  * refi-or-sell switch driven purely by hold length vs. construction maturity
"""

import datetime as _dt
from assumptions import (SCALARS, MILESTONES, UNIT_PLANS, RETAIL_SUITES,
                         all_budget_lines, ms_slug, N_MONTHS)


def _months_in(y, m):
    return (_dt.date(y + (m // 12), (m % 12) + 1, 1) - _dt.date(y, m, 1)).days


def days_in_month(inception_date, p):
    """Actual days in calendar month of period p (1-based)."""
    d0 = _dt.date.fromisoformat(inception_date)
    y = d0.year + (d0.month - 1 + p - 1) // 12
    m = (d0.month - 1 + p - 1) % 12 + 1
    return _months_in(y, m)


def year_index(p):
    """0-based year for growth stepping (annual step)."""
    return (p - 1) // 12


def compute(overrides=None, want_static=True):
    """Run the full model. Returns a dict of scalars and month-indexed lists
    (index 0 == period 1). ``overrides`` may replace any scalar input.
    ``want_static`` toggles the embedded no-growth reference (off for recursion)."""
    V = dict(SCALARS)
    if overrides:
        V.update(overrides)
    N = N_MONTHS

    # ---- Milestones -------------------------------------------------------
    ms = {}
    for name, start, dur in MILESTONES:
        ms[name] = dict(start=start, dur=dur, end=start + dur - 1)
    # Disposition tracks the hold toggle.
    hold = int(V["HoldPeriodMonths"])
    ms["Disposition"]["start"] = hold
    ms["Disposition"]["end"] = hold

    con_start = ms["Construction"]["start"]
    deliver_start = ms["1st Unit Delivery"]["start"]
    stab_month = ms["Stabilization"]["start"]
    con_term = int(V["ConLoanTermMonths"])
    con_maturity = con_start + con_term - 1

    dispo = hold
    refi = dispo > con_maturity
    if refi:
        ovr = int(V["RefiMonthOverride"])
        refi_month = min(con_maturity, ovr if ovr > 0 else con_maturity)
    else:
        refi_month = 0  # inactive

    # ---- Unit matrix ------------------------------------------------------
    total_units = sum(u[2] for u in UNIT_PLANS)
    market_units = sum(u[2] for u in UNIT_PLANS if u[1] == "Market")
    low_units = sum(u[2] for u in UNIT_PLANS if u[1] == "Low")
    nrsf = sum(u[2] * u[3] for u in UNIT_PLANS)
    base_market_rent_mo = sum(u[2] * u[4] for u in UNIT_PLANS if u[1] == "Market")
    base_low_rent_mo = sum(u[2] * u[4] for u in UNIT_PLANS if u[1] == "Low")
    retail_rent_mo = sum(s[1] * s[2] / 12.0 for s in RETAIL_SUITES)

    buildable_sf = nrsf / V["EfficiencyRatio"]

    # ---- Development budget: modeled totals -------------------------------
    lines = all_budget_lines()

    def modeled_total(ln):
        if ln["label"] == "Land Acquisition":
            return V["LandCost"]
        if ln["label"] == "Existing Structure/Demo":
            return V["ExistingStructureCost"]
        a = ln["amount"] or 0
        if ln["basis"] == "abs":
            return a
        if ln["basis"] == "unit":
            return a * total_units
        if ln["basis"] == "sf":
            return a * nrsf
        return a

    for ln in lines:
        ln["total"] = modeled_total(ln)

    direct_costs = sum(ln["total"] for ln in lines)
    land_total = sum(ln["total"] for ln in lines if ln["group"] == "Land")
    soft_total = sum(ln["total"] for ln in lines if ln["group"] == "Soft")
    hard_total = sum(ln["total"] for ln in lines if ln["group"] == "Hard")

    # Development fee (sponsor) on direct costs.
    dev_fee = V["DevFeePctDirect"] * direct_costs

    # ---- Cost curve engine: spread each line monthly ----------------------
    def window(ln):
        m = ms[ln["phase"]]
        s = m["start"] + ln["modifier"]
        e = m["end"] + ln["modifier"]
        s = max(1, min(N, s))
        e = max(s, min(N, e))
        return s, e

    def curve_weights(curve, s, e):
        n = e - s + 1
        if n <= 0:
            return {}
        if curve == "Single_Point":
            return {s: 1.0}
        if curve == "Bell_Curve":
            raw = [(i + 1) * (n - i) for i in range(n)]  # parabolic bell
            tot = float(sum(raw))
            return {s + i: raw[i] / tot for i in range(n)}
        # Straight_Line / Development_Fee
        return {s + i: 1.0 / n for i in range(n)}

    spend = [0.0] * (N + 1)          # total monthly development spend (direct)
    line_spend = {}
    for ln in lines:
        s, e = window(ln)
        w = curve_weights(ln["curve"], s, e)
        arr = [0.0] * (N + 1)
        for p, wt in w.items():
            arr[p] += ln["total"] * wt
        line_spend[ln["label"]] = arr
        for p in range(1, N + 1):
            spend[p] += arr[p]

    # Dev fee spread straight-line across Initial Closing..Stabilization.
    fee_s = ms["Initial Closing"]["start"]
    fee_e = ms["Stabilization"]["start"]
    fee_arr = [0.0] * (N + 1)
    nfee = fee_e - fee_s + 1
    for p in range(fee_s, fee_e + 1):
        fee_arr[p] = dev_fee / nfee

    # ---- Operating model (needed for loan sizing) -------------------------
    # Absorption / occupancy
    stab_units = round(V["StabilizedOccupancy"] * total_units)
    delivered = [0.0] * (N + 1)
    leased = [0.0] * (N + 1)
    occ = [0.0] * (N + 1)
    for p in range(1, N + 1):
        k = p - deliver_start + 1
        if k <= 0:
            d = l = 0
        else:
            d = min(total_units, k * V["UnitsDeliveredPerMo"])
            l = min(stab_units, d, k * V["UnitsLeasedPerMo"])
        delivered[p] = d
        leased[p] = l
        occ[p] = l / total_units if total_units else 0.0

    def gf(rate, p):
        return (1.0 + rate) ** year_index(p)

    # Revenue / expense / NOI (before property tax and after)
    gpr = [0.0] * (N + 1)
    egi = [0.0] * (N + 1)
    noi_bt = [0.0] * (N + 1)      # NOI before property tax
    noi = [0.0] * (N + 1)         # NOI after property tax
    below = [0.0] * (N + 1)       # below-NOI items (cap res, tax prep, AM fee)
    op_tax = [0.0] * (N + 1)
    opex_total = [0.0] * (N + 1)
    retail_rev = [0.0] * (N + 1)

    perunit = dict(
        insurance=V["ExpInsurancePerUnit"], utilities=V["ExpUtilitiesPerUnit"],
        payroll=V["ExpPayrollPerUnit"], rm=V["ExpRMPerUnit"],
        turnover=V["ExpTurnoverPerUnit"], contract=V["ExpContractPerUnit"],
        admin=V["ExpAdminPerUnit"], advertising=V["ExpAdvertisingPerUnit"])
    stab_occ = V["StabilizedOccupancy"]

    for p in range(1, N + 1):
        if p < deliver_start:
            continue
        occp = occ[p]
        ramp = V["LeaseUpRampPct"] + (1 - V["LeaseUpRampPct"]) * (occp / stab_occ if stab_occ else 1)
        ramp = min(1.0, ramp)
        elev = 1.0 + V["LeaseUpElevationPct"] * max(0.0, 1 - (occp / stab_occ if stab_occ else 1))
        # Revenue
        gpr_p = base_market_rent_mo * gf(V["GrowthRent"], p) + base_low_rent_mo * gf(V["GrowthLowIncRent"], p)
        gpr[p] = gpr_p
        vacancy = gpr_p * (1 - occp)
        loss_to_lease = gpr_p * V["LossToLeasePct"]
        bad_debt = gpr_p * V["BadDebtPct"]
        concessions = gpr_p * V["ConcessionsPct"]
        occ_units = leased[p]
        parking = V["ParkingPerUnitMo"] * occ_units * gf(V["GrowthOther"], p)
        util_reimb = V["UtilReimbPerUnitMo"] * occ_units * gf(V["GrowthOther"], p)
        other_inc = V["OtherIncPerUnitMo"] * occ_units * gf(V["GrowthOther"], p)
        retail = retail_rent_mo * (occp / stab_occ if stab_occ else 1) * gf(V["GrowthOther"], p)
        retail = min(retail, retail_rent_mo * gf(V["GrowthOther"], p))
        retail_rev[p] = retail
        egi_p = (gpr_p - vacancy - loss_to_lease - bad_debt - concessions
                 + parking + util_reimb + other_inc + retail)
        egi[p] = egi_p
        # Expenses (monthly). Variable scale by ramp; elevated by elev.
        ge = gf(V["GrowthExpense"], p)
        gu = gf(V["GrowthUtility"], p)
        mo = lambda a: a * total_units / 12.0
        insurance = mo(perunit["insurance"]) * ge
        admin = mo(perunit["admin"]) * ge
        utilities = mo(perunit["utilities"]) * gu * ramp
        rm = mo(perunit["rm"]) * ge * ramp
        turnover = mo(perunit["turnover"]) * ge * ramp
        contract = mo(perunit["contract"]) * ge * ramp
        payroll = mo(perunit["payroll"]) * ge * elev
        advertising = mo(perunit["advertising"]) * ge * elev
        mgmt_fee = V["MgmtFeePct"] * egi_p
        # Property tax
        yrs_from_stab = max(0, year_index(p) - year_index(stab_month))
        if V["TaxMethodValue"] >= 1:
            tax_annual = V["TaxYear1"] * ((1 + V["GrowthTax"]) ** yrs_from_stab)
        else:
            eff = V["TaxLevyPct"] * V["TaxAssessmentPct"] * V["TaxValueAdjFactor"]
            noibt_annual_ref = (gpr_p * stab_occ - vacancy) * 12  # rough proxy, non-circular
            val = max(0.0, noibt_annual_ref) / V["EntryCapRate"] if V["EntryCapRate"] else 0
            tax_annual = val * eff
        tax_m = tax_annual / 12.0
        op_tax[p] = tax_m
        opex_noturn = (insurance + admin + utilities + rm + turnover + contract
                       + payroll + advertising + mgmt_fee)
        opex_total[p] = opex_noturn + tax_m
        noi_bt[p] = egi_p - opex_noturn            # before property tax
        noi[p] = noi_bt[p] - tax_m                 # after property tax
        below[p] = (V["CapReservePerUnit"] * total_units / 12.0
                    + V["TaxPrepPerUnit"] * total_units / 12.0
                    + V["AssetMgmtFeePct"] * egi_p)

    def fwd12(arr, start):
        return sum(arr[p] for p in range(start + 1, start + 13) if 1 <= p <= N)

    stab_noi_annual = fwd12(noi, stab_month - 1)   # 12 months from stabilization
    year1_noi_annual = fwd12(noi, deliver_start - 1)

    # ---- Construction loan sizing (non-circular) --------------------------
    fin_points_placeholder = 0.0  # points depend on loan; resolve below
    recourse_fee_placeholder = 0.0
    # Cost basis for LTC excludes interest reserve (loan-funded on top).
    # Iterate points/recourse once analytically: they scale with loan, but for
    # sizing we use cost basis excluding those, then add them as uses.
    cost_ex_int = direct_costs + dev_fee + V["ConOtherCosts"]  # + points+recourse+opreserve added below
    # Operating reserve (needs op cash flow); compute a first pass w/o reserve.

    # Peak cumulative operating cash shortfall during lease-up (op CF pre-debt)
    op_cf_pre = [0.0] * (N + 1)
    for p in range(1, N + 1):
        op_cf_pre[p] = noi[p] - below[p]
    cum = 0.0
    peak_short = 0.0
    for p in range(deliver_start, stab_month + 1):
        cum += op_cf_pre[p]
        peak_short = min(peak_short, cum)
    op_reserve = max(0.0, -peak_short) * V["OpReserveCoverage"]
    op_reserve = max(op_reserve, V["OpReserveMinMonths"] * (sum(opex_total[stab_month:stab_month + 1]) or 0))

    cost_ex_int += op_reserve

    # Max loan constraints (points/recourse solved via closed form for LTC test)
    max_by_ltc = V["TargetLTC"] * cost_ex_int
    max_by_dy = stab_noi_annual / V["ConMinDebtYield"] if V["ConMinDebtYield"] else 1e18
    max_by_dscr = stab_noi_annual / (V["ConMinDSCR"] * V["ConRate"]) if V["ConRate"] else 1e18
    con_cost_commit = min(max_by_ltc, max_by_dy, max_by_dscr)

    con_points = V["ConPointsPct"] * con_cost_commit
    recourse_fee = V["RecourseFeePct"] * con_cost_commit
    # add financing costs to uses (equity/loan funds them); they are part of cost_ex_int
    cost_ex_int_final = direct_costs + dev_fee + V["ConOtherCosts"] + con_points + recourse_fee + op_reserve
    debt_share = con_cost_commit / cost_ex_int_final if cost_ex_int_final else 0.0

    # Monthly fundable uses (direct spend + dev fee + financing at con start + op reserve at stab)
    uses_m = [0.0] * (N + 1)
    for p in range(1, N + 1):
        uses_m[p] += spend[p] + fee_arr[p]
    close_m = ms["Initial Closing"]["start"]
    uses_m[close_m] += V["ConOtherCosts"] + con_points + recourse_fee
    uses_m[stab_month] += op_reserve

    # ---- Funding order + sequential interest reserve ----------------------
    equity_budget = cost_ex_int_final - con_cost_commit
    con_bal = [0.0] * (N + 1)       # construction loan balance (incl cap. interest)
    debt_draw = [0.0] * (N + 1)
    eq_draw = [0.0] * (N + 1)
    con_int = [0.0] * (N + 1)
    cum_eq = 0.0
    pari = V["FundingOrderPari"] >= 1
    for p in range(1, N + 1):
        # interest on prior balance while loan active (con_start..payoff)
        payoff_month = refi_month if refi else dispo
        if con_start <= p <= payoff_month:
            if V["ConActual360"] >= 1:
                rate_m = V["ConRate"] * days_in_month(V["InceptionDate"], p) / 360.0
            else:
                rate_m = V["ConRate"] / 12.0
        else:
            rate_m = 0.0
        interest = rate_m * con_bal[p - 1]
        # split this month's uses
        u = uses_m[p]
        if pari:
            dd = debt_share * u
            ed = u - dd
        else:  # equity-first
            if cum_eq + u <= equity_budget:
                ed = u
                dd = 0.0
            elif cum_eq >= equity_budget:
                ed = 0.0
                dd = u
            else:
                ed = equity_budget - cum_eq
                dd = u - ed
            cum_eq += ed
        debt_draw[p] = dd
        eq_draw[p] = ed
        con_int[p] = interest
        con_bal[p] = con_bal[p - 1] + dd + interest
    accrued_int = sum(con_int)
    interest_reserve = V["LenderReqReserveOverride"] if V["LenderReqReserveOverride"] > 0 else accrued_int
    con_loan_total = con_cost_commit + interest_reserve
    con_payoff_bal = con_bal[refi_month] if refi else con_bal[dispo]

    tdc = cost_ex_int_final + interest_reserve   # total dev cost incl interest reserve
    equity_total = tdc - con_loan_total

    # ---- Permanent loan / refinance ---------------------------------------
    perm_loan = 0.0
    perm_points = 0.0
    refi_cash_out = 0.0
    value_at_refi = 0.0
    perm_bal = [0.0] * (N + 1)
    perm_ds = [0.0] * (N + 1)
    if refi:
        stab_noi_refi = fwd12(noi, refi_month - 1)  # forward 12mo NOI from refi (incl refi month)
        value_at_refi = stab_noi_refi / V["EntryCapRate"] if V["EntryCapRate"] else 0
        max_ltv = V["PermMaxLTV"] * value_at_refi
        max_dy = stab_noi_refi / V["PermMinDebtYield"] if V["PermMinDebtYield"] else 1e18
        r = V["PermRate"] / 12.0
        nA = int(V["PermAmortMonths"])
        mort_const = (r / (1 - (1 + r) ** (-nA))) * 12 if r > 0 else 12.0 / nA
        max_dscr = stab_noi_refi / (V["PermMinDSCR"] * mort_const) if mort_const else 1e18
        perm_sized = min(max_ltv, max_dy, max_dscr)
        other_refi_costs = 0.0
        if V["NoCashOut"] >= 1:
            perm_need = (con_payoff_bal + other_refi_costs) / (1 - V["PermPointsPct"])
            perm_loan = min(perm_sized, perm_need)
        else:
            perm_loan = perm_sized
        perm_points = V["PermPointsPct"] * perm_loan
        refi_cash_out = perm_loan - con_payoff_bal - perm_points - other_refi_costs
        # amortization schedule from refi_month(+1) to dispo
        bal = perm_loan
        io_end = refi_month + int(V["PermIOMonths"])
        for p in range(refi_month + 1, N + 1):
            if p > dispo:
                break
            rr = V["PermRate"] / 12.0
            if V["PermActual360"] >= 1:
                rr = V["PermRate"] * days_in_month(V["InceptionDate"], p) / 360.0
            interest = rr * bal
            if p <= io_end:
                principal = 0.0
            else:
                pmt = perm_loan * (r / (1 - (1 + r) ** (-nA))) if r > 0 else perm_loan / nA
                principal = max(0.0, pmt - interest)
                principal = min(principal, bal)
            perm_ds[p] = interest + principal
            bal -= principal
            perm_bal[p] = bal
    perm_payoff_bal = perm_bal[dispo] if refi else 0.0

    # ---- Operating cash flow with debt service ----------------------------
    op_cf = [0.0] * (N + 1)     # to equity, after perm debt service (con int capitalized)
    for p in range(1, N + 1):
        ds = perm_ds[p] if refi else 0.0
        # operations only after delivery start and up to disposition
        if deliver_start <= p <= dispo:
            op_cf[p] = noi[p] - below[p] - ds
        else:
            op_cf[p] = 0.0

    # ---- Disposition ------------------------------------------------------
    fwd_noi_bt = fwd12(noi_bt, dispo)        # forward 12mo NOI before property tax
    fwd_noi_after_inplace = fwd12(noi, dispo)
    fwd_egi = fwd12(egi, dispo)
    fwd_retail = fwd12(retail_rev, dispo)
    retail_share = (fwd_retail / fwd_egi) if fwd_egi else 0.0
    # blended exit cap from resi/retail split of forward NOI (before tax basis)
    resi_noi = fwd_noi_bt * (1 - retail_share)
    retail_noi = fwd_noi_bt * retail_share
    denom = (resi_noi / V["ExitCapResidential"] + retail_noi / V["ExitCapRetail"]) if fwd_noi_bt else 0
    blended_cap = fwd_noi_bt / denom if denom else V["ExitCapResidential"]

    if V["ReassessOnSale"] >= 1:
        eff_tax_on_value = V["DispoTaxRatePct"]
        exit_value = fwd_noi_bt / (blended_cap + eff_tax_on_value) if (blended_cap + eff_tax_on_value) else 0
    else:
        exit_value = fwd_noi_after_inplace / blended_cap if blended_cap else 0

    broker = V["BrokerCostPct"] * exit_value
    sales_tax = V["SalesTaxPct"] * exit_value
    sponsor_fee = V["SponsorDispoFeePct"] * exit_value
    tax_proration = op_tax[dispo] * V["TaxProrationMonths"]
    active_payoff = perm_payoff_bal if refi else con_bal[dispo]
    net_sale = exit_value - broker - sales_tax - sponsor_fee - tax_proration - active_payoff

    # ---- Levered & unlevered cash flows -----------------------------------
    lev = [0.0] * (N + 1)
    unlev = [0.0] * (N + 1)
    # unlevered uses = direct + dev fee + op reserve (exclude financing)
    unlev_uses = [0.0] * (N + 1)
    for p in range(1, N + 1):
        unlev_uses[p] = spend[p] + fee_arr[p]
    unlev_uses[stab_month] += op_reserve
    for p in range(1, N + 1):
        lev[p] = -eq_draw[p] + op_cf[p]
        unlev[p] = -unlev_uses[p] + (noi[p] - below[p] if deliver_start <= p <= dispo else 0.0)
    if refi:
        lev[refi_month] += refi_cash_out
    lev[dispo] += net_sale
    # unlevered disposition = exit value net of sale costs (no loan payoff)
    unlev[dispo] += exit_value - broker - sales_tax - sponsor_fee - tax_proration

    # ---- Returns ----------------------------------------------------------
    def irr_monthly(cf):
        # cf: list periods 1..N ; solve monthly rate via bisection; annualize.
        # Multiply-by-discount-factor form avoids float division-by-zero, and a
        # safe bracket [-0.20, 0.50]/mo keeps the discount powers finite.
        last = max((p for p in range(1, N + 1) if abs(cf[p]) > 1e-9), default=0)
        if last == 0:
            return None
        vals = [cf[p] for p in range(1, last + 1)]

        def npv(r):
            d = 1.0 / (1.0 + r)
            s = 0.0
            f = 1.0
            for v in vals:
                s += v * f
                f *= d
            return s
        lo, hi = -0.20, 0.50
        flo, fhi = npv(lo), npv(hi)
        if flo != flo or fhi != fhi or flo * fhi > 0:
            return None
        mid = (lo + hi) / 2
        for _ in range(200):
            mid = (lo + hi) / 2
            fm = npv(mid)
            if abs(fm) < 1e-6:
                break
            if flo * fm < 0:
                hi = mid
            else:
                lo = mid
                flo = fm
        return (1 + mid) ** 12 - 1

    def moic(cf):
        inflow = sum(cf[p] for p in range(1, N + 1) if cf[p] > 0)
        outflow = -sum(cf[p] for p in range(1, N + 1) if cf[p] < 0)
        return inflow / outflow if outflow else None

    lev_irr = irr_monthly(lev)
    unlev_irr = irr_monthly(unlev)
    lev_moic = moic(lev)
    unlev_moic = moic(unlev)

    yield_on_cost = stab_noi_annual / tdc if tdc else 0
    dev_spread = yield_on_cost - V["ExitCapResidential"]
    roc_current = year1_noi_annual / tdc if tdc else 0
    roc_stab = stab_noi_annual / tdc if tdc else 0

    # ---- Static (no-growth) unlevered reference ---------------------------
    static = None
    if want_static:
        ov = dict(overrides or {})
        for k in ("GrowthRent", "GrowthLowIncRent", "GrowthExpense",
                  "GrowthUtility", "GrowthTax", "GrowthOther"):
            ov[k] = 0.0
        s = compute(ov, want_static=False)
        static = dict(unlev_irr=s["unlev_irr"], unlev_moic=s["unlev_moic"],
                      lev_irr=s["lev_irr"], lev_moic=s["lev_moic"],
                      yield_on_cost=s["yield_on_cost"])

    return dict(
        # timing
        ms=ms, con_start=con_start, con_maturity=con_maturity, deliver_start=deliver_start,
        stab_month=stab_month, dispo=dispo, refi=refi, refi_month=refi_month,
        # units
        total_units=total_units, market_units=market_units, low_units=low_units,
        nrsf=nrsf, buildable_sf=buildable_sf,
        # budget
        land_total=land_total, soft_total=soft_total, hard_total=hard_total,
        direct_costs=direct_costs, dev_fee=dev_fee, line_spend=line_spend,
        spend=spend, fee_arr=fee_arr,
        # loan sizing
        con_cost_commit=con_cost_commit, debt_share=debt_share, con_points=con_points,
        recourse_fee=recourse_fee, op_reserve=op_reserve, peak_short=peak_short,
        max_by_ltc=max_by_ltc, max_by_dy=max_by_dy, max_by_dscr=max_by_dscr,
        # interest reserve
        con_bal=con_bal, con_int=con_int, debt_draw=debt_draw, eq_draw=eq_draw,
        accrued_int=accrued_int, interest_reserve=interest_reserve,
        con_loan_total=con_loan_total, con_payoff_bal=con_payoff_bal,
        tdc=tdc, equity_total=equity_total,
        # operating
        delivered=delivered, leased=leased, occ=occ, gpr=gpr, egi=egi,
        noi=noi, noi_bt=noi_bt, below=below, op_tax=op_tax, op_cf=op_cf,
        stab_noi_annual=stab_noi_annual, year1_noi_annual=year1_noi_annual,
        # perm
        value_at_refi=value_at_refi, perm_loan=perm_loan, perm_points=perm_points,
        refi_cash_out=refi_cash_out, perm_bal=perm_bal, perm_ds=perm_ds,
        perm_payoff_bal=perm_payoff_bal,
        # disposition
        fwd_noi_bt=fwd_noi_bt, fwd_noi_after_inplace=fwd_noi_after_inplace,
        blended_cap=blended_cap, exit_value=exit_value, broker=broker,
        sales_tax=sales_tax, sponsor_fee=sponsor_fee, tax_proration=tax_proration,
        active_payoff=active_payoff, net_sale=net_sale,
        # cash flows / returns
        lev=lev, unlev=unlev, lev_irr=lev_irr, unlev_irr=unlev_irr,
        lev_moic=lev_moic, unlev_moic=unlev_moic,
        yield_on_cost=yield_on_cost, dev_spread=dev_spread,
        roc_current=roc_current, roc_stab=roc_stab,
        tdc_per_unit=tdc / total_units if total_units else 0,
        tdc_per_sf=tdc / nrsf if nrsf else 0,
        static=static,
        V=V,
    )


if __name__ == "__main__":
    import json
    r = compute()
    keys = ["total_units", "nrsf", "direct_costs", "dev_fee", "land_total",
            "soft_total", "hard_total", "con_cost_commit", "debt_share",
            "interest_reserve", "accrued_int", "con_loan_total", "con_payoff_bal",
            "op_reserve", "peak_short", "tdc", "equity_total", "stab_noi_annual",
            "year1_noi_annual", "value_at_refi", "perm_loan", "refi_cash_out",
            "perm_payoff_bal", "exit_value", "blended_cap", "net_sale",
            "lev_irr", "unlev_irr", "lev_moic", "unlev_moic",
            "yield_on_cost", "dev_spread", "roc_current", "roc_stab",
            "tdc_per_unit", "tdc_per_sf", "refi", "refi_month", "dispo",
            "con_maturity", "stab_month", "deliver_start"]
    for k in keys:
        v = r[k]
        if isinstance(v, float):
            print(f"{k:22s} {v:,.4f}")
        else:
            print(f"{k:22s} {v}")
    print("static:", {k: (round(v, 4) if isinstance(v, float) else v)
                       for k, v in r["static"].items()})
