"""
Phase 15 — Diagnostics & validation.

A plain-English integrity layer: each check computes a value and a PASS/FAIL,
covering formula errors, budget allocation, cost-curve/draw/interest/loan
reconciliation, the refi-or-sell switch, the No-Cash-Out and reassess-on-sale
toggles, lease-up, reserves, disposition and cash-flow identities.  A single
"total errors found" count aggregates every failure.
"""

from common import (write_title, write_section, label, put_formula, put_text,
                    set_col_widths, FMT_USD0, FMT_NUM0, font_subhead,
                    FILL_CHECK_OK, FILL_CHECK_BAD, FILL_SUBHEAD, ALIGN_L, ALIGN_C)
from layout import (SH_DIAG, SH_CF, SH_CON, SH_OPS, set_row, r, tref,
                    FIRST_COL_LETTER, LAST_COL_LETTER)


def build(wb, reg):
    ws = wb[SH_DIAG]
    set_col_widths(ws, {"A": 48, "B": 18, "C": 12})
    ws.sheet_view.showGridLines = False
    write_title(ws, 1, "Diagnostics & Validation")

    Rlev = r(SH_CF, "lev"); Runlev = r(SH_CF, "unlev")
    Rcbal = r(SH_CON, "con_bal"); Rnoi = r(SH_OPS, "noi")
    levrng = f"'{SH_CF}'!{FIRST_COL_LETTER}${Rlev}:'{SH_CF}'!{LAST_COL_LETTER}${Rlev}"
    unlrng = f"'{SH_CF}'!{FIRST_COL_LETTER}${Runlev}:'{SH_CF}'!{LAST_COL_LETTER}${Runlev}"
    cbalrng = f"'{SH_CON}'!{FIRST_COL_LETTER}${Rcbal}:'{SH_CON}'!{LAST_COL_LETTER}${Rcbal}"
    noirng = f"'{SH_OPS}'!{FIRST_COL_LETTER}${Rnoi}:'{SH_OPS}'!{LAST_COL_LETTER}${Rnoi}"
    spend_total = tref(SH_SPEND := "DevSpend", "spend_total")

    err = ("SUMPRODUCT(--ISERROR(" + levrng + "))+SUMPRODUCT(--ISERROR(" + unlrng + "))"
           "+SUMPRODUCT(--ISERROR(" + cbalrng + "))+SUMPRODUCT(--ISERROR(" + noirng + "))")

    # (description, value formula, pass formula)
    checks = [
        ("Total formula errors (key ranges = 0)", err, f"({err})=0"),
        ("Development budget fully allocated", f"{spend_total}-DirectCosts", f"ABS({spend_total}-DirectCosts)<1"),
        ("Cost curves sum to line totals", "Chk_CostCurve", "ABS(Chk_CostCurve)<1"),
        ("Dev fee curve sums to fee", "Chk_DevFeeCurve", "ABS(Chk_DevFeeCurve)<1"),
        ("Construction debt draws reconcile", "Chk_DebtDraws", "ABS(Chk_DebtDraws)<1"),
        ("Equity draws reconcile", "Chk_EquityDraws", "ABS(Chk_EquityDraws)<1"),
        ("Fundable uses reconcile", "Chk_Uses", "ABS(Chk_Uses)<1"),
        ("Interest reserve funded >= accrued", "Chk_IntReserve", "Chk_IntReserve>=-1"),
        ("Construction balance drawn to zero", "Chk_ConBalEnd", "ABS(Chk_ConBalEnd)<1"),
        ("Lease-up reaches stabilized occupancy", "Chk_LeaseUpReached", "ABS(Chk_LeaseUpReached)<1"),
        ("Units delivered reach total by stabilization", "Chk_UnitsDelivered-TotalUnits", "Chk_UnitsDelivered>=TotalUnits-1"),
        ("Operating reserve >= peak lease-up shortfall", "OpReserve+PeakShort", "OpReserve>=-PeakShort-1"),
        ("Refi-or-sell switch behaves correctly", "PermLoan",
         "AND(IF(RefiFlag=0,PermLoan=0,TRUE()),IF(RefiFlag=1,PermLoan>0,TRUE()))"),
        ("Permanent loan retires construction", "Chk_LoanRetired", "Chk_LoanRetired>=-1"),
        ("Refinance sources = uses", "Chk_RefiBalance", "ABS(Chk_RefiBalance)<1"),
        ("No-Cash-Out toggle functions (no cash distributed)", "RefiCashOut", "IF(NoCashOut=1,RefiCashOut<=1,TRUE())"),
        ("Reassess-on-sale closed form reconciles", "Chk_Reassess", "ABS(Chk_Reassess)<1"),
        ("Property value reconciles to NOI & cap",
         "ExitValue-IF(ReassessOnSale>=1,FwdNOIbt/(BlendedExitCap+DispoEffTaxRate),FwdNOIafter/BlendedExitCap)",
         "ABS(ExitValue-IF(ReassessOnSale>=1,FwdNOIbt/(BlendedExitCap+DispoEffTaxRate),FwdNOIafter/BlendedExitCap))<1"),
        ("Levered cash-flow identity", "LevNetProfit-(LevDistros-LevEquityIn)", "ABS(LevNetProfit-(LevDistros-LevEquityIn))<1"),
        ("Unlevered cash-flow identity", "UnlevNetProfit-(UnlevDistros-UnlevInvested)", "ABS(UnlevNetProfit-(UnlevDistros-UnlevInvested))<1"),
        ("Perm term not expiring before disposition", "IF(RefiFlag=1,RefiMonth+PermTermMonths-DispoMonth,0)",
         "IF(RefiFlag=1,RefiMonth+PermTermMonths>=DispoMonth,TRUE())"),
        ("Hold period within 1..360 months", "HoldPeriodMonths", "AND(HoldPeriodMonths>=1,HoldPeriodMonths<=360)"),
    ]

    hr = 3
    for j, h in enumerate(["Check", "Value", "Status"], start=1):
        c = ws.cell(row=hr, column=j, value=h); c.font = font_subhead()
        c.fill = FILL_SUBHEAD; c.alignment = ALIGN_L if j == 1 else ALIGN_C
    row = hr + 1
    first = row
    for desc, valf, passf in checks:
        put_text(ws, row, 1, desc)
        put_formula(ws, row, 2, valf, fmt=FMT_USD0)
        c = put_formula(ws, row, 3, f'IF({passf},"PASS","FAIL")', fmt="@")
        c.alignment = ALIGN_C
        row += 1
    last = row - 1

    row += 1
    write_section(ws, row, "Summary", span=3); row += 1
    label(ws, row, "TOTAL ERRORS FOUND", col=1, bold=True)
    c = put_formula(ws, row, 2, f'COUNTIF(C{first}:C{last},"FAIL")', fmt=FMT_NUM0, bold=True)
    reg.add("ErrorsFound", SH_DIAG, f"B{row}")
    set_row(SH_DIAG, "ErrorsFound", row)
    err_row = row
    row += 1
    label(ws, row, "MODEL STATUS", col=1, bold=True)
    put_formula(ws, row, 2, f'IF(B{err_row}=0,"ALL CHECKS PASS","REVIEW FAILURES")', fmt="@", bold=True)

    # conditional formatting for PASS/FAIL
    from openpyxl.formatting.rule import CellIsRule
    ws.conditional_formatting.add(
        f"C{first}:C{last}",
        CellIsRule(operator="equal", formula=['"PASS"'], fill=FILL_CHECK_OK))
    ws.conditional_formatting.add(
        f"C{first}:C{last}",
        CellIsRule(operator="equal", formula=['"FAIL"'], fill=FILL_CHECK_BAD))
    ws.freeze_panes = "A4"
