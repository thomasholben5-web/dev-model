"""
Phase 13 — Return calculations & static reference.

Levered/unlevered IRR (Excel IRR on the monthly series, annualized) and MOIC,
gross and net of sponsor fees; development spread (yield-on-cost vs exit cap),
return on cost (current & stabilized), total development cost ($, $/unit, $/SF),
current/exit basis per unit, LTC and debt yield.  A static no-growth reference
recomputes the unlevered NOI with all growth factors = 1 and reports its
IRR/MOIC.
"""

from common import (write_title, write_section, write_subhead, label, put_formula,
                    put_text, set_col_widths, FMT_USD0, FMT_USD0_DOLLAR, FMT_PCT2,
                    FMT_MULT, FMT_BPS, FMT_NUM0, FILL_TOTAL)
from layout import (SH_RET, SH_OPS, SH_CF, SH_LEASE, SH_TIME, set_row, r,
                    period_header, grid_row, mref, FIRST_COL_LETTER, LAST_COL_LETTER)


def build(wb, reg):
    _static_noi_rows(wb, reg)
    _returns(wb, reg)


# ---------------------------------------------------------------------------
def _static_noi_rows(wb, reg):
    """Append no-growth (static) NOI rows to the Operating sheet."""
    ws = wb[SH_OPS]
    Rramp = r(SH_OPS, "ramp"); Relev = r(SH_OPS, "elev")
    row = r(SH_OPS, "SumNOI") + 2
    row = write_section(ws, row, "Static (No-Growth) Reference NOI", span=3)
    row = period_header(ws, row)
    mo = "TotalUnits/12"

    def egi_expr(p, C):
        occ = mref(SH_LEASE, "occ", p); leased = mref(SH_LEASE, "leased", p)
        return (f"(BaseMarketRentMo+BaseLowRentMo)*({occ}-LossToLeasePct-BadDebtPct-ConcessionsPct)"
                f"+{leased}*(ParkingPerUnitMo+UtilReimbPerUnitMo+OtherIncPerUnitMo)"
                f"+RetailRentMo*MIN(1,{occ}/StabilizedOccupancy)")

    def f_noibt(p, C, Cprev):
        E = egi_expr(p, C)
        return (f"IF({p}>=DeliverStart,({E})*(1-MgmtFeePct)"
                f"-{mo}*(ExpInsurancePerUnit+ExpAdminPerUnit)"
                f"-{mo}*{C}${Rramp}*(ExpUtilitiesPerUnit+ExpRMPerUnit+ExpTurnoverPerUnit+ExpContractPerUnit)"
                f"-{mo}*{C}${Relev}*(ExpPayrollPerUnit+ExpAdvertisingPerUnit),0)")
    row = grid_row(ws, row, "Static NOI before Tax", f_noibt, fmt=FMT_USD0,
                   key="s_noibt", sheet=SH_OPS)
    Rsnoibt = r(SH_OPS, "s_noibt")

    def f_stax(p, C, Cprev):
        eff = "TaxLevyPct*TaxAssessmentPct*TaxValueAdjFactor"
        return (f"IF({p}>=DeliverStart,IF(TaxMethodValue>=1,TaxYear1/12,"
                f"MAX(0,{C}${Rsnoibt}*12)/EntryCapRate*({eff})/12),0)")
    row = grid_row(ws, row, "Static Property Tax", f_stax, fmt=FMT_USD0,
                   key="s_tax", sheet=SH_OPS)
    Rstax = r(SH_OPS, "s_tax")

    row = grid_row(ws, row, "Static NOI",
                   lambda p, C, Cprev: f"{C}${Rsnoibt}-{C}${Rstax}",
                   fmt=FMT_USD0, key="s_noi", sheet=SH_OPS)
    Rsnoi = r(SH_OPS, "s_noi")

    def f_sbelow(p, C, Cprev):
        E = egi_expr(p, C)
        return f"IF({p}>=DeliverStart,{mo}*(CapReservePerUnit+TaxPrepPerUnit)+AssetMgmtFeePct*({E}),0)"
    row = grid_row(ws, row, "Static Below-NOI", f_sbelow, fmt=FMT_USD0,
                   key="s_below", sheet=SH_OPS)
    Rsbelow = r(SH_OPS, "s_below")

    row = grid_row(ws, row, "Static Unlevered Operating CF",
                   lambda p, C, Cprev: f"IF({mref(SH_TIME,'OpsActive',p)}=1,{C}${Rsnoi}-{C}${Rsbelow},0)",
                   fmt=FMT_USD0, key="s_unlev_ops", sheet=SH_OPS)


# ---------------------------------------------------------------------------
def _returns(wb, reg):
    ws = wb[SH_RET]
    set_col_widths(ws, {"A": 40, "B": 18, "C": 20})
    ws.sheet_view.showGridLines = False
    write_title(ws, 1, "Return Calculations")

    Rlev = r(SH_CF, "lev"); Runlev = r(SH_CF, "unlev")
    levrng = f"'{SH_CF}'!{FIRST_COL_LETTER}${Rlev}:'{SH_CF}'!{LAST_COL_LETTER}${Rlev}"
    unlrng = f"'{SH_CF}'!{FIRST_COL_LETTER}${Runlev}:'{SH_CF}'!{LAST_COL_LETTER}${Runlev}"

    row = 3
    row = write_section(ws, row, "Levered & Unlevered Returns (net of fees)", span=3)

    def out(lbl, formula, name, fmt=FMT_MULT):
        label(ws, row_ref[0], lbl, col=1)
        c = put_formula(ws, row_ref[0], 2, formula, fmt=fmt, link=True, bold=True)
        reg.add(name, SH_RET, f"B{row_ref[0]}")
        set_row(SH_RET, name, row_ref[0])
        row_ref[0] += 1
    row_ref = [row]
    out("Levered IRR", f'IFERROR((1+IRR({levrng},0.01))^12-1,"n/a")', "LevIRR", fmt=FMT_PCT2)
    out("Unlevered IRR", f'IFERROR((1+IRR({unlrng},0.01))^12-1,"n/a")', "UnlevIRR", fmt=FMT_PCT2)
    out("Levered MOIC", "IF(LevEquityIn=0,0,LevDistros/LevEquityIn)", "LevMOIC")
    out("Unlevered MOIC", "IF(UnlevInvested=0,0,UnlevDistros/UnlevInvested)", "UnlevMOIC")

    # Gross of sponsor fees (dev fee + disposition fee added back at disposition)
    row = row_ref[0] + 1
    row = write_subhead(ws, row, "Gross of Sponsor Fees", span=3)
    row = period_header(ws, row)
    def f_glev(p, C, Cprev):
        return f"{mref(SH_CF,'lev',p)}+IF({p}=DispoMonth,DevFee+SponsorDispoFeePct*ExitValue,0)"
    row = grid_row(ws, row, "Levered CF (gross of fees)", f_glev, fmt=FMT_USD0,
                   key="glev", sheet=SH_RET)
    def f_gunlev(p, C, Cprev):
        return f"{mref(SH_CF,'unlev',p)}+IF({p}=DispoMonth,DevFee+SponsorDispoFeePct*ExitValue,0)"
    row = grid_row(ws, row, "Unlevered CF (gross of fees)", f_gunlev, fmt=FMT_USD0,
                   key="gunlev", sheet=SH_RET)
    Rglev = r(SH_RET, "glev"); Rgunlev = r(SH_RET, "gunlev")
    glevrng = f"{FIRST_COL_LETTER}{Rglev}:{LAST_COL_LETTER}{Rglev}"
    gunlrng = f"{FIRST_COL_LETTER}{Rgunlev}:{LAST_COL_LETTER}{Rgunlev}"
    row += 1
    row_ref2 = [row]
    def out2(lbl, formula, name, fmt=FMT_MULT):
        label(ws, row_ref2[0], lbl, col=1)
        c = put_formula(ws, row_ref2[0], 2, formula, fmt=fmt, link=True, bold=True)
        reg.add(name, SH_RET, f"B{row_ref2[0]}")
        set_row(SH_RET, name, row_ref2[0])
        row_ref2[0] += 1
    out2("Levered IRR (gross)", f'IFERROR((1+IRR({glevrng},0.01))^12-1,"n/a")', "LevIRRgross", fmt=FMT_PCT2)
    out2("Unlevered IRR (gross)", f'IFERROR((1+IRR({gunlrng},0.01))^12-1,"n/a")', "UnlevIRRgross", fmt=FMT_PCT2)
    out2("Levered MOIC (gross)", f'IF(LevEquityIn=0,0,(SUMIF({glevrng},">0"))/LevEquityIn)', "LevMOICgross")
    out2("Unlevered MOIC (gross)", f'IF(UnlevInvested=0,0,(SUMIF({gunlrng},">0"))/UnlevInvested)', "UnlevMOICgross")

    # Development metrics
    row = row_ref2[0] + 1
    row = write_section(ws, row, "Development Metrics", span=3)
    row_ref3 = [row]
    def out3(lbl, formula, name, fmt=FMT_PCT2):
        label(ws, row_ref3[0], lbl, col=1)
        c = put_formula(ws, row_ref3[0], 2, formula, fmt=fmt, link=True, bold=True)
        reg.add(name, SH_RET, f"B{row_ref3[0]}")
        set_row(SH_RET, name, row_ref3[0])
        row_ref3[0] += 1
    out3("Yield on Cost (stabilized)", "IF(TDC_Total=0,0,StabNOI/TDC_Total)", "YieldOnCost")
    out3("Development Spread (%)", "YieldOnCost-ExitCapResidential", "DevSpreadPct")
    out3("Development Spread (bps)", "(YieldOnCost-ExitCapResidential)*10000", "DevSpreadBps", fmt=FMT_BPS)
    out3("Return on Cost — Current (Yr-1)", "IF(TDC_Total=0,0,Year1NOI/TDC_Total)", "ROCcurrent")
    out3("Return on Cost — Stabilized", "IF(TDC_Total=0,0,StabNOI/TDC_Total)", "ROCstab")
    out3("Debt Yield — Stabilized (con)", "IF(ConLoanTotal=0,0,StabNOI/ConLoanTotal)", "DebtYieldStab")
    out3("Total Development Cost", "TDC_Total", "TDCout", fmt=FMT_USD0)
    out3("TDC per Unit", "IF(TotalUnits=0,0,TDC_Total/TotalUnits)", "TDCperUnit", fmt=FMT_USD0_DOLLAR)
    out3("TDC per NRSF", "IF(NRSF=0,0,TDC_Total/NRSF)", "TDCperSF", fmt=FMT_USD0_DOLLAR)
    out3("Current Basis per Unit", "IF(TotalUnits=0,0,TDC_Total/TotalUnits)", "BasisPerUnitCur", fmt=FMT_USD0_DOLLAR)
    out3("Exit Basis per Unit", "IF(TotalUnits=0,0,ExitValue/TotalUnits)", "BasisPerUnitExit", fmt=FMT_USD0_DOLLAR)

    # ---- Static (no-growth) reference ------------------------------------
    row = row_ref3[0] + 1
    row = write_section(ws, row, "Static (No-Growth) Reference — Unlevered", span=3)
    row = period_header(ws, row)
    Rsnoibt = r(SH_OPS, "s_noibt"); Rsnoi = r(SH_OPS, "s_noi"); Rstax = r(SH_OPS, "s_tax")
    snoibt_rng = f"'{SH_OPS}'!{FIRST_COL_LETTER}${Rsnoibt}:'{SH_OPS}'!{LAST_COL_LETTER}${Rsnoibt}"
    snoi_rng = f"'{SH_OPS}'!{FIRST_COL_LETTER}${Rsnoi}:'{SH_OPS}'!{LAST_COL_LETTER}${Rsnoi}"
    stax_rng = f"'{SH_OPS}'!{FIRST_COL_LETTER}${Rstax}:'{SH_OPS}'!{LAST_COL_LETTER}${Rstax}"

    # static forward NOI & exit value scalars
    row_ref4 = [row]
    def sset(lbl, formula, name, fmt=FMT_USD0):
        label(ws, row_ref4[0], lbl, col=1)
        c = put_formula(ws, row_ref4[0], 2, formula, fmt=fmt, link=True)
        reg.add(name, SH_RET, f"B{row_ref4[0]}")
        set_row(SH_RET, name, row_ref4[0])
        row_ref4[0] += 1
    sset("Static Fwd NOI before Tax",
         f"SUM(INDEX({snoibt_rng},DispoMonth+1):INDEX({snoibt_rng},DispoMonth+12))", "StaticFwdNOIbt")
    sset("Static Fwd NOI after Tax",
         f"SUM(INDEX({snoi_rng},DispoMonth+1):INDEX({snoi_rng},DispoMonth+12))", "StaticFwdNOI")
    sset("Static Exit Value",
         ("IF(ReassessOnSale>=1,IF((BlendedExitCap+DispoEffTaxRate)=0,0,"
          "StaticFwdNOIbt/(BlendedExitCap+DispoEffTaxRate)),"
          "IF(BlendedExitCap=0,0,StaticFwdNOI/BlendedExitCap))"), "StaticExitValue")
    sset("Static Net Sale (unlevered)",
         (f"StaticExitValue*(1-BrokerCostPct-SalesTaxPct-SponsorDispoFeePct)"
          f"-INDEX({stax_rng},MAX(1,DispoMonth))*TaxProrationMonths"), "StaticNetSaleUnlev")

    row = row_ref4[0]
    row = period_header(ws, row)
    def f_sunlev(p, C, Cprev):
        return (f"-{mref(SH_CF,'unlev_uses',p)}+{mref(SH_OPS,'s_unlev_ops',p)}"
                f"+IF({p}=DispoMonth,StaticNetSaleUnlev,0)")
    row = grid_row(ws, row, "Static Unlevered CF", f_sunlev, fmt=FMT_USD0,
                   key="s_unlev", sheet=SH_RET)
    Rsu = r(SH_RET, "s_unlev")
    surng = f"{FIRST_COL_LETTER}{Rsu}:{LAST_COL_LETTER}{Rsu}"
    row += 1
    label(ws, row, "Static Unlevered IRR", col=1)
    put_formula(ws, row, 2, f'IFERROR((1+IRR({surng},0.01))^12-1,"n/a")', fmt=FMT_PCT2, link=True, bold=True)
    reg.add("StaticUnlevIRR", SH_RET, f"B{row}"); set_row(SH_RET, "StaticUnlevIRR", row); row += 1
    label(ws, row, "Static Unlevered MOIC", col=1)
    put_formula(ws, row, 2, f'IF(SUMIF({surng},"<0")=0,0,SUMIF({surng},">0")/-SUMIF({surng},"<0"))',
                fmt=FMT_MULT, link=True, bold=True)
    reg.add("StaticUnlevMOIC", SH_RET, f"B{row}"); set_row(SH_RET, "StaticUnlevMOIC", row)
    ws.freeze_panes = "A3"
