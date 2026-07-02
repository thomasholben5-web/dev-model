"""
Phase 08 — Property tax engine.

Operating-period taxes are computed monthly in phase 07 (value or income
approach, both non-circular).  Here we expose the effective tax-rate-on-value
used by the reassess-on-sale closed form, the disposition-period rate, and a
reconciliation of the stabilized tax.
"""

from common import (write_section, write_subhead, label, put_formula,
                    FMT_USD0, FMT_PCT3, FMT_PCT2, FILL_TOTAL)
from layout import SH_OPS, set_row, r, FIRST_COL_LETTER, LAST_COL_LETTER


def build(wb, reg):
    ws = wb[SH_OPS]
    row = r(SH_OPS, "OpReserve") + 2
    row = write_section(ws, row, "Property Tax Engine", span=3)

    def sc(lbl, formula, name, fmt=FMT_PCT3):
        label(ws, row_ref[0], lbl, col=1)
        c = put_formula(ws, row_ref[0], 2, formula, fmt=fmt, link=True)
        reg.add(name, SH_OPS, f"B{row_ref[0]}")
        set_row(SH_OPS, name, row_ref[0])
        row_ref[0] += 1
    row_ref = [row]
    sc("Effective Tax Rate on Value (operating)",
       "TaxLevyPct*TaxAssessmentPct*TaxValueAdjFactor", "EffTaxRateOnValue")
    sc("Disposition-Period Tax Rate on Value", "DispoTaxRatePct", "DispoEffTaxRate")

    Rtax = r(SH_OPS, "tax")
    taxrng = f"{FIRST_COL_LETTER}{Rtax}:{LAST_COL_LETTER}{Rtax}"
    sc("Stabilized Annual Property Tax (fwd 12mo)",
       f"SUM(INDEX({taxrng},StabMonth):INDEX({taxrng},StabMonth+11))", "StabTaxAnnual", fmt=FMT_USD0)

    # Reconciliation: value approach year-1 tax equals the input when method=value
    label(ws, row_ref[0], "Tax Method (1=value / 0=income)", col=1)
    put_formula(ws, row_ref[0], 2, "TaxMethodValue", fmt="#,##0")
    reg.add("Chk_TaxMethod", SH_OPS, f"B{row_ref[0]}")
    set_row(SH_OPS, "Chk_TaxMethod", row_ref[0])
