"""
Phase 12 — Levered & unlevered monthly cash flows.

Levered (to equity): −equity draws + operating CF to equity + refi cash
distribution at the refi month + net sale proceeds at disposition.
Unlevered (property): −project uses (ex-financing) + unlevered operating CF +
unlevered net sale proceeds at disposition.
"""

from common import (write_title, write_section, write_subhead, label, put_formula,
                    set_col_widths, FMT_USD0, FILL_TOTAL)
from layout import (SH_CF, SH_CON, SH_OPS, SH_SPEND, SH_TIME, SH_PERM, set_row,
                    r, period_header, grid_row, mref, tref)


def build(wb, reg):
    ws = wb[SH_CF]
    set_col_widths(ws, {"A": 32, "B": 12, "C": 16})
    ws.sheet_view.showGridLines = False
    write_title(ws, 1, "Levered & Unlevered Cash Flows (monthly)")

    row = 3
    row = period_header(ws, row)

    # Unlevered project uses (exclude financing costs & interest reserve)
    def f_unlev_uses(p, C, Cprev):
        return (f"{mref(SH_SPEND,'spend_total',p)}+{mref(SH_SPEND,'fee_spend',p)}"
                f"+IF({p}=StabMonth,OpReserve,0)")
    row = grid_row(ws, row, "Unlevered Project Uses", f_unlev_uses, fmt=FMT_USD0,
                   key="unlev_uses", sheet=SH_CF)

    # Levered equity cash flow (op CF net of perm debt service assembled here)
    def f_lev(p, C, Cprev):
        opcf = (f"IF({mref(SH_TIME,'OpsActive',p)}=1,"
                f"{mref(SH_OPS,'op_cf_pre',p)}-{mref(SH_PERM,'perm_ds',p)},0)")
        return (f"-{mref(SH_CON,'eq_draw',p)}+{opcf}"
                f"+IF({p}=RefiMonth,RefiCashOut,0)+IF({p}=DispoMonth,NetSale,0)")
    row = grid_row(ws, row, "Levered Cash Flow (to equity)", f_lev, fmt=FMT_USD0,
                   bold=True, key="lev", sheet=SH_CF)
    Runl = r(SH_CF, "unlev_uses")

    # Unlevered property cash flow
    def f_unlev(p, C, Cprev):
        return (f"-{C}${Runl}+{mref(SH_OPS,'unlev_ops',p)}"
                f"+IF({p}=DispoMonth,NetSaleUnlev,0)")
    row = grid_row(ws, row, "Unlevered Cash Flow (property)", f_unlev, fmt=FMT_USD0,
                   bold=True, key="unlev", sheet=SH_CF)

    # ---- Reconciliation aggregates ---------------------------------------
    row += 1
    row = write_subhead(ws, row, "Cash Flow Reconciliation", span=3)
    Rlev = r(SH_CF, "lev"); Runlev = r(SH_CF, "unlev")
    from layout import FIRST_COL_LETTER, LAST_COL_LETTER
    levrng = f"{FIRST_COL_LETTER}{Rlev}:{LAST_COL_LETTER}{Rlev}"
    unlrng = f"{FIRST_COL_LETTER}{Runlev}:{LAST_COL_LETTER}{Runlev}"

    def agg(lbl, formula, name):
        label(ws, row_ref[0], lbl, col=1)
        c = put_formula(ws, row_ref[0], 2, formula, fmt=FMT_USD0, link=True, bold=True)
        c.fill = FILL_TOTAL
        reg.add(name, SH_CF, f"B{row_ref[0]}")
        set_row(SH_CF, name, row_ref[0])
        row_ref[0] += 1
    row_ref = [row]
    agg("Levered Equity Invested (Σ outflows)", f'-SUMIF({levrng},"<0")', "LevEquityIn")
    agg("Levered Distributions (Σ inflows)", f'SUMIF({levrng},">0")', "LevDistros")
    agg("Levered Net Profit", f"SUM({levrng})", "LevNetProfit")
    agg("Unlevered Invested (Σ outflows)", f'-SUMIF({unlrng},"<0")', "UnlevInvested")
    agg("Unlevered Distributions (Σ inflows)", f'SUMIF({unlrng},">0")', "UnlevDistros")
    agg("Unlevered Net Profit", f"SUM({unlrng})", "UnlevNetProfit")
    ws.freeze_panes = "D4"
