"""
Phase 04 — Construction loan sizing & monthly draw schedule.

Sizes the construction loan to MIN(target LTC, min debt yield, min DSCR) on a
cost basis excluding capitalized interest (so the interest reserve, resolved
sequentially in phase 05, never creates circularity).  Then draws monthly
against the fundable-uses curve under the funding-order toggle (equity-first or
pari-passu), rolling the balance forward.
"""

from common import (write_title, write_section, write_subhead, label, put_formula,
                    put_text, set_col_widths, FMT_USD0, FMT_PCT2, FMT_NUM0,
                    FMT_MULT, FMT_MONTHS, font_subhead, FILL_TOTAL)
from layout import (SH_CON, SH_SPEND, SH_TIME, set_row, r, period_header,
                    date_header, grid_row, mref)


def build(wb, reg):
    ws = wb[SH_CON]
    set_col_widths(ws, {"A": 34, "B": 18, "C": 16})
    ws.sheet_view.showGridLines = False
    write_title(ws, 1, "Construction Loan — Sizing & Monthly Draw Schedule")

    # ---- Sizing block -----------------------------------------------------
    row = 3
    row = write_subhead(ws, row, "Loan Sizing (min of LTC / Debt Yield / DSCR)", span=3)
    def sc(lbl, formula, name, fmt=FMT_USD0, note=""):
        label(ws, row_ref[0], lbl, col=1)
        put_formula(ws, row_ref[0], 2, formula, fmt=fmt, link=True, bold=True)
        if note:
            put_text(ws, row_ref[0], 3, note, italic=True)
        reg.add(name, SH_CON, f"B{row_ref[0]}")
        set_row(SH_CON, name, row_ref[0])
        row_ref[0] += 1
    row_ref = [row]
    sc("Cost Basis for LTC (excl. cap. int.)",
       "DirectCosts+DevFee+ConOtherCosts+OpReserve", "ConCostBasis")
    sc("Max Loan by Target LTC", "TargetLTC*ConCostBasis", "MaxByLTC")
    sc("Max Loan by Min Debt Yield", "StabNOI/ConMinDebtYield", "MaxByDebtYield")
    sc("Max Loan by Min DSCR (I/O)", "StabNOI/(ConMinDSCR*ConRate)", "MaxByDSCR")
    sc("Construction Loan Commitment", "MIN(MaxByLTC,MaxByDebtYield,MaxByDSCR)", "ConCostCommit")
    sc("Origination Points", "ConPointsPct*ConCostCommit", "ConPoints")
    sc("Loan Recourse / Indemnity Fee", "RecourseFeePct*ConCostCommit", "RecourseFee")
    sc("Total Uses ex-Interest (final)",
       "DirectCosts+DevFee+ConOtherCosts+ConPoints+RecourseFee+OpReserve", "ConCostExInt")
    sc("Debt Share of Uses", "ConCostCommit/ConCostExInt", "DebtShare", fmt=FMT_PCT2)
    sc("Equity Budget (ex-Interest)", "ConCostExInt-ConCostCommit", "EquityBudget")

    # Sizing tests pass/fail
    row = row_ref[0] + 1
    row = write_subhead(ws, row, "Construction Sizing Tests", span=3)
    def test(lbl, formula, name, fmt=FMT_MULT):
        label(ws, row_ref2[0], lbl, col=1)
        put_formula(ws, row_ref2[0], 2, formula, fmt=fmt)
        reg.add(name, SH_CON, f"B{row_ref2[0]}")
        set_row(SH_CON, name, row_ref2[0])
        row_ref2[0] += 1
    row_ref2 = [row]
    test("Debt Yield at Commitment", "StabNOI/ConCostCommit", "ConDYtest", fmt=FMT_PCT2)
    test("DSCR at Commitment (I/O)", "StabNOI/(ConRate*ConCostCommit)", "ConDSCRtest")
    test("Realized LTC (incl. int.)", "ConLoanTotal/TDC_Total", "ConRealizedLTC", fmt=FMT_PCT2)
    label(ws, row_ref2[0], "Debt Yield Test Pass", col=1)
    put_formula(ws, row_ref2[0], 2, 'IF(ConDYtest>=ConMinDebtYield,"PASS","FAIL")', fmt="@")
    row_ref2[0] += 1
    label(ws, row_ref2[0], "DSCR Test Pass", col=1)
    put_formula(ws, row_ref2[0], 2, 'IF(ConDSCRtest>=ConMinDSCR,"PASS","FAIL")', fmt="@")
    row_ref2[0] += 1

    # ---- Monthly draw schedule -------------------------------------------
    row = row_ref2[0] + 1
    row = write_section(ws, row, "Monthly Draw Schedule", span=3)
    R_hdr = row
    row = period_header(ws, row)
    row = date_header(ws, row)
    # Pre-assign monthly row numbers for forward references
    R_days = row
    R_uses = row + 1
    R_eq = row + 2
    R_cumeq = row + 3
    R_debt = row + 4
    R_ratem = row + 5
    R_int = row + 6
    R_bal = row + 7

    grid_row(ws, R_days, "Days in Month",
             lambda p, C, Cprev: f"DAY(EOMONTH(EDATE(InceptionDate,{p}-1),0))",
             fmt=FMT_NUM0, total=None, key="days", sheet=SH_CON)

    def f_uses(p, C, Cprev):
        return (f"{mref(SH_SPEND,'spend_total',p)}+{mref(SH_SPEND,'fee_spend',p)}"
                f"+IF({p}=MS_Initial_Closing_Start,ConOtherCosts+ConPoints+RecourseFee,0)"
                f"+IF({p}=StabMonth,OpReserve,0)")
    grid_row(ws, R_uses, "Fundable Uses", f_uses, fmt=FMT_USD0, key="uses", sheet=SH_CON)

    def f_eq(p, C, Cprev):
        cumprev = "0" if Cprev is None else f"{Cprev}${R_cumeq}"
        u = f"{C}${R_uses}"
        return (f"IF(FundingOrderPari=1,(1-DebtShare)*{u},"
                f"MIN({u},MAX(0,EquityBudget-({cumprev}))))")
    grid_row(ws, R_eq, "Equity Draw", f_eq, fmt=FMT_USD0, key="eq_draw", sheet=SH_CON)

    def f_cumeq(p, C, Cprev):
        prev = "0" if Cprev is None else f"{Cprev}${R_cumeq}"
        return f"{prev}+{C}${R_eq}"
    grid_row(ws, R_cumeq, "Cumulative Equity", f_cumeq, fmt=FMT_USD0, total=None,
             key="cum_eq", sheet=SH_CON)

    grid_row(ws, R_debt, "Debt Draw",
             lambda p, C, Cprev: f"{C}${R_uses}-{C}${R_eq}",
             fmt=FMT_USD0, key="debt_draw", sheet=SH_CON)

    def f_ratem(p, C, Cprev):
        active = mref(SH_TIME, "ConActive", p)
        return (f"IF({active}=1,IF(ConActual360=1,ConRate*{C}${R_days}/360,ConRate/12),0)")
    grid_row(ws, R_ratem, "Monthly Interest Rate", f_ratem, fmt="0.0000%",
             total=None, key="ratem", sheet=SH_CON)

    def f_int(p, C, Cprev):
        balprev = "0" if Cprev is None else f"{Cprev}${R_bal}"
        return f"{C}${R_ratem}*({balprev})"
    grid_row(ws, R_int, "Interest Accrued (capitalized)", f_int, fmt=FMT_USD0,
             key="con_int", sheet=SH_CON)

    def f_bal(p, C, Cprev):
        balprev = "0" if Cprev is None else f"{Cprev}${R_bal}"
        return f"IF({p}>ConPayoffMonth,0,{balprev}+{C}${R_debt}+{C}${R_int})"
    grid_row(ws, R_bal, "Construction Loan Balance", f_bal, fmt=FMT_USD0,
             total=None, key="con_bal", sheet=SH_CON)

    ws.freeze_panes = "D3"
