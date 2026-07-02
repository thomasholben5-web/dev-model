"""
Phase 14 — Dashboard.

A single summary view that updates automatically from the inputs: key returns
(levered/unlevered IRR & MOIC, gross & net, static reference), development
metrics, sources & uses, timeline, loan sizing tests, the refi-or-sell status,
and the diagnostics error count.
"""

from common import (write_title, write_section, write_subhead, label, put_formula,
                    put_text, set_col_widths, FMT_USD0, FMT_USD0_DOLLAR, FMT_PCT2,
                    FMT_MULT, FMT_BPS, FMT_NUM0, FMT_MONTHS, FMT_DATE, FMT_TEXT,
                    font_title, FILL_TOTAL, FILL_CHECK_OK, ALIGN_R)
from layout import SH_DASH, set_row


def build(wb, reg):
    ws = wb[SH_DASH]
    set_col_widths(ws, {"A": 30, "B": 18, "C": 6, "D": 30, "E": 18})
    ws.sheet_view.showGridLines = False
    write_title(ws, 1, "Institutional Multifamily Development Model — Dashboard")
    put_formula(ws, 2, 1, '"" & PropertyName & "  •  " & City & ", " & State & "  •  " & ProductType',
                fmt=FMT_TEXT, link=True)

    # ---- Deal snapshot ----------------------------------------------------
    row = 4
    write_section(ws, row, "Deal Snapshot", span=2, col=1); row += 1
    snap = [
        ("Total Units", "TotalUnits", FMT_NUM0),
        ("Net Rentable SF", "NRSF", FMT_NUM0),
        ("Hold Period (months)", "HoldPeriodMonths", FMT_MONTHS),
        ("Refi or Sell", 'IF(RefiFlag=1,"Refinance then hold","Sell during construction")', FMT_TEXT),
        ("Construction Maturity (mo)", "ConMaturity", FMT_MONTHS),
        ("Refi Month", "RefiMonth", FMT_MONTHS),
        ("Disposition Month", "DispoMonth", FMT_MONTHS),
    ]
    for lbl, f, fmt in snap:
        label(ws, row, lbl, col=1)
        put_formula(ws, row, 2, f, fmt=fmt, link=True)
        row += 1

    # ---- Key returns ------------------------------------------------------
    r2 = 4
    write_section(ws, r2, "Key Returns", span=2, col=4); r2 += 1
    rets = [
        ("Levered IRR (net)", "LevIRR", FMT_PCT2),
        ("Levered IRR (gross)", "LevIRRgross", FMT_PCT2),
        ("Unlevered IRR (net)", "UnlevIRR", FMT_PCT2),
        ("Levered MOIC (net)", "LevMOIC", FMT_MULT),
        ("Unlevered MOIC (net)", "UnlevMOIC", FMT_MULT),
        ("Static Unlevered IRR", "StaticUnlevIRR", FMT_PCT2),
        ("Static Unlevered MOIC", "StaticUnlevMOIC", FMT_MULT),
    ]
    for lbl, f, fmt in rets:
        label(ws, r2, lbl, col=4)
        c = put_formula(ws, r2, 5, f, fmt=fmt, link=True, bold=True)
        r2 += 1

    # ---- Development metrics ---------------------------------------------
    row = 13
    write_section(ws, row, "Development Metrics", span=2, col=1); row += 1
    dm = [
        ("Yield on Cost (stab)", "YieldOnCost", FMT_PCT2),
        ("Development Spread", "DevSpreadBps", FMT_BPS),
        ("Return on Cost (current)", "ROCcurrent", FMT_PCT2),
        ("Return on Cost (stab)", "ROCstab", FMT_PCT2),
        ("Debt Yield (stab)", "DebtYieldStab", FMT_PCT2),
        ("TDC", "TDC_Total", FMT_USD0),
        ("TDC / Unit", "TDCperUnit", FMT_USD0_DOLLAR),
        ("TDC / NRSF", "TDCperSF", FMT_USD0_DOLLAR),
        ("Exit Value", "ExitValue", FMT_USD0),
        ("Exit Basis / Unit", "BasisPerUnitExit", FMT_USD0_DOLLAR),
    ]
    for lbl, f, fmt in dm:
        label(ws, row, lbl, col=1)
        put_formula(ws, row, 2, f, fmt=fmt, link=True)
        row += 1

    # ---- Sources & Uses ---------------------------------------------------
    r3 = 13
    write_section(ws, r3, "Sources & Uses", span=2, col=4); r3 += 1
    uses = [
        ("Land", "BudgetLand"),
        ("Soft Costs", "BudgetSoft"),
        ("Hard Costs", "BudgetHard"),
        ("Development Fee", "DevFee"),
        ("Financing Costs", "ConPoints+RecourseFee+ConOtherCosts"),
        ("Interest Reserve", "InterestReserveFunded"),
        ("Operating Reserve", "OpReserve"),
    ]
    for lbl, f in uses:
        label(ws, r3, lbl, col=4)
        put_formula(ws, r3, 5, f, fmt=FMT_USD0, link=True)
        r3 += 1
    label(ws, r3, "Total Uses (TDC)", col=4, bold=True)
    c = put_formula(ws, r3, 5, "TDC_Total", fmt=FMT_USD0, link=True, bold=True); c.fill = FILL_TOTAL
    r3 += 2
    label(ws, r3, "Construction Loan", col=4)
    put_formula(ws, r3, 5, "ConLoanTotal", fmt=FMT_USD0, link=True); r3 += 1
    label(ws, r3, "Equity", col=4)
    put_formula(ws, r3, 5, "EquityTotal", fmt=FMT_USD0, link=True); r3 += 1
    label(ws, r3, "Total Sources", col=4, bold=True)
    c = put_formula(ws, r3, 5, "ConLoanTotal+EquityTotal", fmt=FMT_USD0, link=True, bold=True)
    c.fill = FILL_TOTAL
    r3 += 1
    label(ws, r3, "Sources − Uses (check)", col=4)
    put_formula(ws, r3, 5, "ConLoanTotal+EquityTotal-TDC_Total", fmt=FMT_USD0, link=True)

    # ---- Loan sizing & status --------------------------------------------
    row += 1
    write_section(ws, row, "Loan Sizing & Status", span=2, col=1); row += 1
    ls = [
        ("Construction Loan", "ConLoanTotal", FMT_USD0),
        ("LTC (realized)", "LTC_Realized", FMT_PCT2),
        ("Con Debt Yield Test", "ConDYtest", FMT_PCT2),
        ("Con DSCR Test", "ConDSCRtest", FMT_MULT),
        ("Permanent Loan", "PermLoan", FMT_USD0),
        ("Perm LTV Test", "PermLTVtest", FMT_PCT2),
        ("Perm DSCR Test", "PermDSCRtest", FMT_MULT),
        ("Refi Cash Distribution", "RefiCashOut", FMT_USD0),
    ]
    for lbl, f, fmt in ls:
        label(ws, row, lbl, col=1)
        put_formula(ws, row, 2, f, fmt=fmt, link=True)
        row += 1

    # ---- Validation banner ------------------------------------------------
    r4 = row if row > 27 else 27
    write_section(ws, r4, "Validation", span=2, col=4); r4 += 1
    label(ws, r4, "Total Errors Found", col=4, bold=True)
    c = put_formula(ws, r4, 5, "ErrorsFound", fmt=FMT_NUM0, link=True, bold=True); r4 += 1
    label(ws, r4, "Model Status", col=4, bold=True)
    c = put_formula(ws, r4, 5, 'IF(ErrorsFound=0,"ALL CHECKS PASS","REVIEW FAILURES")',
                    fmt=FMT_TEXT, link=True, bold=True)
    c.fill = FILL_CHECK_OK
    ws.freeze_panes = "A3"
