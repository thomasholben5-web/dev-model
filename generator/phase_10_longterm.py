"""
Phase 10 — Long-term operating cash flow.

The operating grid in phase 07 already runs the full horizon (up to 360 months
plus the 12 forward months for exit NOI), with perm debt service applied on the
active loan.  This phase rolls the operating cash flow up over the hold and
registers the totals used by the cash-flow and dashboard phases.
"""

from common import (write_subhead, label, put_formula, FMT_USD0, FILL_TOTAL)
from layout import SH_OPS, set_row, r, tref


def build(wb, reg):
    ws = wb[SH_OPS]
    row = r(SH_OPS, "Chk_TaxMethod") + 2
    row = write_subhead(ws, row, "Long-Term Operating Cash Flow (over hold)", span=3)

    def agg(lbl, ref, name):
        label(ws, row_ref[0], lbl, col=1)
        c = put_formula(ws, row_ref[0], 2, ref.lstrip("="), fmt=FMT_USD0, link=True, bold=True)
        c.fill = FILL_TOTAL
        reg.add(name, SH_OPS, f"B{row_ref[0]}")
        set_row(SH_OPS, name, row_ref[0])
        row_ref[0] += 1
    row_ref = [row]
    agg("Σ Operating CF before Debt (property)", tref(SH_OPS, "op_cf_pre"), "SumOpCFpre")
    agg("Σ Unlevered Operating CF (property)", tref(SH_OPS, "unlev_ops"), "SumUnlevOps")
    agg("Σ NOI over horizon", tref(SH_OPS, "noi"), "SumNOI")
