"""
Phase 09 — Permanent loan / refinance engine.

Active only when the hold extends past construction maturity (RefiFlag=1).
Sizes the takeout to MIN(max LTV, min debt yield, min DSCR) on the refi-date
forward NOI, honours the No-Cash-Out toggle (solve loan so net cash at refi =
0 via closed form), pays off the construction loan, and amortizes monthly
(interest-only period then level amortization) through disposition.
"""

from common import (write_title, write_section, write_subhead, label, put_formula,
                    put_text, set_col_widths, FMT_USD0, FMT_PCT2, FMT_MULT,
                    FMT_MONTHS, FMT_NUM0, FILL_TOTAL, N_MONTHS)
from layout import (SH_PERM, SH_OPS, SH_TIME, set_row, r, period_header,
                    date_header, grid_row, mref, FIRST_COL_LETTER, LAST_COL_LETTER)


def build(wb, reg):
    ws = wb[SH_PERM]
    set_col_widths(ws, {"A": 34, "B": 18, "C": 16})
    ws.sheet_view.showGridLines = False
    write_title(ws, 1, "Permanent Loan / Refinance Engine")

    noirng = f"'{SH_OPS}'!{FIRST_COL_LETTER}${r(SH_OPS,'noi')}:'{SH_OPS}'!{LAST_COL_LETTER}${r(SH_OPS,'noi')}"

    row = 3
    row = write_subhead(ws, row, "Refinance Sizing (min of LTV / Debt Yield / DSCR)", span=3)
    def sc(lbl, formula, name, fmt=FMT_USD0, note=""):
        label(ws, row_ref[0], lbl, col=1)
        c = put_formula(ws, row_ref[0], 2, formula, fmt=fmt, link=True, bold=True)
        if note:
            put_text(ws, row_ref[0], 3, note, italic=True)
        reg.add(name, SH_PERM, f"B{row_ref[0]}")
        set_row(SH_PERM, name, row_ref[0])
        row_ref[0] += 1
    row_ref = [row]
    sc("Refi Forward NOI (12mo)",
       f"RefiFlag*SUM(INDEX({noirng},MAX(1,RefiMonth)):INDEX({noirng},MAX(1,RefiMonth)+11))",
       "StabNOIrefi")
    sc("Property Value at Refi", "IF(EntryCapRate=0,0,StabNOIrefi/EntryCapRate)", "ValueAtRefi")
    sc("Mortgage Constant (annual)",
       "IF(PermRate=0,12/PermAmortMonths,(PermRate/12)/(1-(1+PermRate/12)^(-PermAmortMonths))*12)",
       "MortConst", fmt=FMT_PCT2)
    sc("Max Loan by LTV", "PermMaxLTV*ValueAtRefi", "MaxLTVperm")
    sc("Max Loan by Debt Yield", "IF(PermMinDebtYield=0,0,StabNOIrefi/PermMinDebtYield)", "MaxDYperm")
    sc("Max Loan by DSCR", "IF(MortConst=0,0,StabNOIrefi/(PermMinDSCR*MortConst))", "MaxDSCRperm")
    sc("Perm Loan — Max Sized", "MIN(MaxLTVperm,MaxDYperm,MaxDSCRperm)", "PermSized")
    sc("Perm Loan — No-Cash-Out Need", "ConPayoffBal/(1-PermPointsPct)", "PermNeedNCO")
    sc("Permanent Loan Amount",
       "RefiFlag*IF(NoCashOut>=1,MIN(PermSized,PermNeedNCO),PermSized)", "PermLoan")
    sc("Permanent Loan Points", "PermPointsPct*PermLoan", "PermPoints")
    sc("Net Cash Distribution at Refi", "RefiFlag*(PermLoan-ConPayoffBal-PermPoints)", "RefiCashOut")
    sc("Level Monthly Payment", "PermLoan*MortConst/12", "PermPmt")

    # Perm tests
    row = row_ref[0] + 1
    row = write_subhead(ws, row, "Permanent Sizing Tests", span=3)
    def tst(lbl, formula, name, fmt=FMT_MULT):
        label(ws, row_ref2[0], lbl, col=1)
        put_formula(ws, row_ref2[0], 2, formula, fmt=fmt)
        reg.add(name, SH_PERM, f"B{row_ref2[0]}")
        set_row(SH_PERM, name, row_ref2[0])
        row_ref2[0] += 1
    row_ref2 = [row]
    tst("LTV at Loan", "IF(ValueAtRefi=0,0,PermLoan/ValueAtRefi)", "PermLTVtest", fmt=FMT_PCT2)
    tst("Debt Yield at Loan", "IF(PermLoan=0,0,StabNOIrefi/PermLoan)", "PermDYtest", fmt=FMT_PCT2)
    tst("DSCR at Loan", "IF(PermLoan=0,0,StabNOIrefi/(PermPmt*12))", "PermDSCRtest")

    # ---- Monthly amortization --------------------------------------------
    row = row_ref2[0] + 1
    row = write_section(ws, row, "Monthly Amortization Schedule", span=3)
    row = period_header(ws, row)
    row = date_header(ws, row)
    R_days = row
    R_ratem = row + 1
    R_int = row + 2
    R_prin = row + 3
    R_ds = row + 4
    R_bal = row + 5

    grid_row(ws, R_days, "Days in Month",
             lambda p, C, Cprev: f"DAY(EOMONTH(EDATE(InceptionDate,{p}-1),0))",
             fmt=FMT_NUM0, total=None, key="pdays", sheet=SH_PERM)

    def f_ratem(p, C, Cprev):
        active = mref(SH_TIME, "PermActive", p)
        return f"IF({active}=1,IF(PermActual360=1,PermRate*{C}${R_days}/360,PermRate/12),0)"
    grid_row(ws, R_ratem, "Monthly Interest Rate", f_ratem, fmt="0.0000%",
             total=None, key="pratem", sheet=SH_PERM)

    def f_int(p, C, Cprev):
        balprev = "0" if Cprev is None else f"{Cprev}${R_bal}"
        return f"{C}${R_ratem}*({balprev})"
    grid_row(ws, R_int, "Interest", f_int, fmt=FMT_USD0, key="pint", sheet=SH_PERM)

    def f_prin(p, C, Cprev):
        balprev = "0" if Cprev is None else f"{Cprev}${R_bal}"
        active = mref(SH_TIME, "PermActive", p)
        io = "IF(OR({p}<=RefiMonth+PermIOMonths,{active}=0),0,MIN({bp},PermPmt-{C}${R_int}))"
        return io.format(p=p, active=active, bp=balprev, C=C, R_int=R_int)
    grid_row(ws, R_prin, "Principal", f_prin, fmt=FMT_USD0, key="pprin", sheet=SH_PERM)

    def f_ds(p, C, Cprev):
        active = mref(SH_TIME, "PermActive", p)
        return f"IF({active}=1,{C}${R_int}+{C}${R_prin},0)"
    grid_row(ws, R_ds, "Debt Service", f_ds, fmt=FMT_USD0, key="perm_ds", sheet=SH_PERM)

    def f_bal(p, C, Cprev):
        balprev = "0" if Cprev is None else f"{Cprev}${R_bal}"
        return (f"IF(RefiFlag=0,0,IF({p}<RefiMonth,0,IF({p}=RefiMonth,PermLoan,"
                f"IF({p}>DispoMonth,0,{balprev}-{C}${R_prin}))))")
    grid_row(ws, R_bal, "Permanent Loan Balance", f_bal, fmt=FMT_USD0,
             total=None, key="perm_bal", sheet=SH_PERM)

    # Payoff balance at disposition
    row = R_bal + 2
    balrng = f"{FIRST_COL_LETTER}{R_bal}:{LAST_COL_LETTER}{R_bal}"
    label(ws, row, "Permanent Loan Payoff at Disposition", col=1)
    c = put_formula(ws, row, 2, f"RefiFlag*INDEX({balrng},DispoMonth)", fmt=FMT_USD0, link=True, bold=True)
    c.fill = FILL_TOTAL
    reg.add("PermPayoffBal", SH_PERM, f"B{row}")
    set_row(SH_PERM, "PermPayoffBal", row)
    row += 1
    # Refi sources (perm loan) must equal uses (con payoff + points + cash-out)
    label(ws, row, "Refi sources − uses (must reconcile)", col=1)
    put_formula(ws, row, 2, "RefiFlag*(PermLoan-PermPoints-RefiCashOut-ConPayoffBal)", fmt=FMT_USD0)
    reg.add("Chk_RefiBalance", SH_PERM, f"B{row}")
    set_row(SH_PERM, "Chk_RefiBalance", row)
    ws.freeze_panes = "D3"
