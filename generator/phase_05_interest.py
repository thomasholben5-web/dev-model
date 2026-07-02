"""
Phase 05 — Interest reserve & construction-loan totals.

The monthly interest was accrued sequentially in phase 04 (each month on the
prior month's ending balance — no circularity).  Here we total it into the
interest reserve (with a lender-required override), roll up the construction
loan total, total development cost and equity, and post the reconciliation
checks (debt/equity draws, ending balance drawn to zero, reserve funded vs
accrued).
"""

from common import (write_section, write_subhead, label, put_formula, put_text,
                    FMT_USD0, FMT_PCT2, font_subhead, FILL_TOTAL)
from layout import (SH_CON, set_row, r, tref, FIRST_COL_LETTER, LAST_COL_LETTER)


def build(wb, reg):
    ws = wb[SH_CON]
    R_bal = r(SH_CON, "con_bal")
    row = R_bal + 2
    row = write_section(ws, row, "Interest Reserve & Loan Totals", span=3)

    conbal_rng = f"{FIRST_COL_LETTER}{R_bal}:{LAST_COL_LETTER}{R_bal}"

    def sc(lbl, formula, name, fmt=FMT_USD0):
        label(ws, row_ref[0], lbl, col=1)
        c = put_formula(ws, row_ref[0], 2, formula, fmt=fmt, link=True, bold=True)
        c.fill = FILL_TOTAL
        reg.add(name, SH_CON, f"B{row_ref[0]}")
        set_row(SH_CON, name, row_ref[0])
        row_ref[0] += 1
    row_ref = [row]
    sc("Interest Reserve — Accrued", tref(SH_CON, "con_int").lstrip("="), "InterestReserveAccrued")
    sc("Interest Reserve — Funded (w/ override)",
       "IF(LenderReqReserveOverride>0,LenderReqReserveOverride,InterestReserveAccrued)",
       "InterestReserveFunded")
    sc("Construction Loan Total", "ConCostCommit+InterestReserveFunded", "ConLoanTotal")
    sc("Construction Loan Payoff Balance", f"INDEX({conbal_rng},ConPayoffMonth)", "ConPayoffBal")
    sc("Total Development Cost (TDC)", "ConCostExInt+InterestReserveFunded", "TDC_Total")
    sc("Total Equity", "TDC_Total-ConLoanTotal", "EquityTotal")
    sc("Loan-to-Cost (realized)", "ConLoanTotal/TDC_Total", "LTC_Realized", fmt=FMT_PCT2)

    # Reconciliation checks
    row = row_ref[0] + 1
    row = write_subhead(ws, row, "Construction Reconciliation", span=3)
    def chk(lbl, formula, name):
        label(ws, row_ref2[0], lbl, col=1)
        put_formula(ws, row_ref2[0], 2, formula, fmt=FMT_USD0)
        reg.add(name, SH_CON, f"B{row_ref2[0]}")
        set_row(SH_CON, name, row_ref2[0])
        row_ref2[0] += 1
    row_ref2 = [row]
    chk("Σ Debt Draws − Commitment", f"{tref(SH_CON,'debt_draw')}-ConCostCommit", "Chk_DebtDraws")
    chk("Σ Equity Draws − Equity Budget", f"{tref(SH_CON,'eq_draw')}-EquityBudget", "Chk_EquityDraws")
    chk("Σ Uses − Total Uses ex-Int", f"{tref(SH_CON,'uses')}-ConCostExInt", "Chk_Uses")
    chk("Ending Balance (must be 0)", f"INDEX({conbal_rng},{_N()})", "Chk_ConBalEnd")
    chk("Reserve Funded − Accrued (>=0)", "InterestReserveFunded-InterestReserveAccrued", "Chk_IntReserve")


def _N():
    from common import N_MONTHS
    return N_MONTHS
