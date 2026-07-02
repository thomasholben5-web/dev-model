"""
Phase 11 — Disposition engine.

Exit value = forward-looking 12-month NOI / exit cap.  Reassess-on-sale is
solved in closed form (Exit = Fwd NOI-before-tax / (exit cap + tax rate on
value)) to avoid circularity; otherwise in-place forward NOI / exit cap.  A
blended residential/retail exit cap is derived from the forward NOI split.  Net
proceeds deduct broker, sales tax, sponsor fee, tax proration and the payoff of
the active loan (construction or permanent per the refi-or-sell switch).
"""

from common import (write_title, write_section, write_subhead, label, put_formula,
                    put_text, set_col_widths, FMT_USD0, FMT_PCT2, FILL_TOTAL)
from layout import (SH_DISPO, SH_OPS, set_row, r, FIRST_COL_LETTER, LAST_COL_LETTER)


def build(wb, reg):
    ws = wb[SH_DISPO]
    set_col_widths(ws, {"A": 40, "B": 18, "C": 30})
    ws.sheet_view.showGridLines = False
    write_title(ws, 1, "Disposition Engine")

    def rng(key):
        rr = r(SH_OPS, key)
        return f"'{SH_OPS}'!{FIRST_COL_LETTER}${rr}:'{SH_OPS}'!{LAST_COL_LETTER}${rr}"
    noibt = rng("noi_bt"); noi = rng("noi"); egi = rng("egi"); retail = rng("retail")
    tax = rng("tax")

    row = 3
    row = write_section(ws, row, "Exit Valuation (forward 12-month NOI)", span=3)

    def sc(lbl, formula, name, fmt=FMT_USD0, note=""):
        label(ws, row_ref[0], lbl, col=1)
        c = put_formula(ws, row_ref[0], 2, formula, fmt=fmt, link=True, bold=True)
        if note:
            put_text(ws, row_ref[0], 3, note, italic=True)
        reg.add(name, SH_DISPO, f"B{row_ref[0]}")
        set_row(SH_DISPO, name, row_ref[0])
        row_ref[0] += 1
    row_ref = [row]
    sc("Forward NOI before Property Tax (12mo)",
       f"SUM(INDEX({noibt},DispoMonth+1):INDEX({noibt},DispoMonth+12))", "FwdNOIbt")
    sc("Forward NOI after In-Place Tax (12mo)",
       f"SUM(INDEX({noi},DispoMonth+1):INDEX({noi},DispoMonth+12))", "FwdNOIafter")
    sc("Forward EGI (12mo)",
       f"SUM(INDEX({egi},DispoMonth+1):INDEX({egi},DispoMonth+12))", "FwdEGI")
    sc("Forward Retail Income (12mo)",
       f"SUM(INDEX({retail},DispoMonth+1):INDEX({retail},DispoMonth+12))", "FwdRetail")
    sc("Retail Revenue Share", "IF(FwdEGI=0,0,FwdRetail/FwdEGI)", "RetailShare", fmt=FMT_PCT2)
    sc("Blended Exit Cap Rate",
       ("IF(FwdNOIbt=0,ExitCapResidential,"
        "IF((FwdNOIbt*(1-RetailShare)/ExitCapResidential+FwdNOIbt*RetailShare/ExitCapRetail)=0,"
        "ExitCapResidential,"
        "FwdNOIbt/(FwdNOIbt*(1-RetailShare)/ExitCapResidential+FwdNOIbt*RetailShare/ExitCapRetail)))"),
       "BlendedExitCap", fmt=FMT_PCT2)
    sc("Exit Value (reassess-aware, closed form)",
       ("IF(ReassessOnSale>=1,"
        "IF((BlendedExitCap+DispoEffTaxRate)=0,0,FwdNOIbt/(BlendedExitCap+DispoEffTaxRate)),"
        "IF(BlendedExitCap=0,0,FwdNOIafter/BlendedExitCap))"),
       "ExitValue")

    # ---- Sale costs & net proceeds ---------------------------------------
    row = row_ref[0] + 1
    row = write_section(ws, row, "Sale Costs & Net Proceeds", span=3)
    row_ref2 = [row]
    def sc2(lbl, formula, name, fmt=FMT_USD0):
        label(ws, row_ref2[0], lbl, col=1)
        c = put_formula(ws, row_ref2[0], 2, formula, fmt=fmt, link=True)
        reg.add(name, SH_DISPO, f"B{row_ref2[0]}")
        set_row(SH_DISPO, name, row_ref2[0])
        row_ref2[0] += 1
    sc2("Gross Sale Price", "ExitValue", "GrossSale")
    sc2("Broker / Sale Costs", "-BrokerCostPct*ExitValue", "BrokerCost")
    sc2("Transfer / Sales Tax", "-SalesTaxPct*ExitValue", "SalesTaxAmt")
    sc2("Sponsor Disposition Fee", "-SponsorDispoFeePct*ExitValue", "SponsorDispoFee")
    sc2("Property Tax Proration", f"-INDEX({tax},MAX(1,DispoMonth))*TaxProrationMonths", "TaxProration")
    sc2("Active Loan Payoff", "-IF(RefiFlag=1,PermPayoffBal,ConPayoffBal)", "ActivePayoff")
    label(ws, row_ref2[0], "Net Sale Proceeds (levered)", col=1, bold=True)
    c = put_formula(ws, row_ref2[0], 2,
                    "GrossSale+BrokerCost+SalesTaxAmt+SponsorDispoFee+TaxProration+ActivePayoff",
                    fmt=FMT_USD0, link=True, bold=True)
    c.fill = FILL_TOTAL
    reg.add("NetSale", SH_DISPO, f"B{row_ref2[0]}")
    set_row(SH_DISPO, "NetSale", row_ref2[0])
    row_ref2[0] += 1
    label(ws, row_ref2[0], "Net Sale Proceeds (unlevered)", col=1, bold=True)
    c = put_formula(ws, row_ref2[0], 2,
                    "GrossSale+BrokerCost+SalesTaxAmt+SponsorDispoFee+TaxProration",
                    fmt=FMT_USD0, link=True, bold=True)
    c.fill = FILL_TOTAL
    reg.add("NetSaleUnlev", SH_DISPO, f"B{row_ref2[0]}")
    set_row(SH_DISPO, "NetSaleUnlev", row_ref2[0])
    row_ref2[0] += 1

    # Gain & checks
    row = row_ref2[0] + 1
    row = write_subhead(ws, row, "Disposition Reconciliation", span=3)
    label(ws, row, "Development Gain (Exit − TDC)", col=1)
    put_formula(ws, row, 2, "ExitValue-TDC_Total", fmt=FMT_USD0); row += 1
    label(ws, row, "Loan retired by proceeds (>=0)", col=1)
    put_formula(ws, row, 2,
                "GrossSale+BrokerCost+SalesTaxAmt+SponsorDispoFee+TaxProration-IF(RefiFlag=1,PermPayoffBal,ConPayoffBal)",
                fmt=FMT_USD0)
    reg.add("Chk_LoanRetired", SH_DISPO, f"B{row}")
    set_row(SH_DISPO, "Chk_LoanRetired", row)
    row += 1
    label(ws, row, "Reassess closed-form residual (=0)", col=1)
    put_formula(ws, row, 2,
                "IF(ReassessOnSale>=1,FwdNOIbt-ExitValue*(BlendedExitCap+DispoEffTaxRate),0)",
                fmt=FMT_USD0)
    reg.add("Chk_Reassess", SH_DISPO, f"B{row}")
    set_row(SH_DISPO, "Chk_Reassess", row)
    ws.freeze_panes = "A3"
