"""
Phase 06 — Unit delivery & lease-up engine.

Monthly cumulative unit deliveries and absorption from the 1st Unit Delivery
milestone, capped by total units and the stabilized-occupancy leased target.
Physical occupancy = leased / total units and drives revenue and the occupancy-
based operating-expense ramp/elevation in the operating model.
"""

from common import (write_title, write_section, write_subhead, label, put_formula,
                    set_col_widths, FMT_NUM0, FMT_PCT1, FILL_TOTAL)
from layout import (SH_LEASE, SH_TIME, set_row, r, period_header, grid_row,
                    FIRST_COL_LETTER, LAST_COL_LETTER)


def build(wb, reg):
    ws = wb[SH_LEASE]
    set_col_widths(ws, {"A": 30, "B": 10, "C": 14})
    ws.sheet_view.showGridLines = False
    write_title(ws, 1, "Unit Delivery & Lease-Up Engine")

    row = 3
    label(ws, row, "Stabilized Leased Target (units)", col=1)
    put_formula(ws, row, 2, "ROUND(StabilizedOccupancy*TotalUnits,0)", fmt=FMT_NUM0, link=True)
    reg.add("StabLeasedTarget", SH_LEASE, f"B{row}")
    set_row(SH_LEASE, "StabLeasedTarget", row)
    row += 2

    write_section(ws, row, "Monthly Delivery & Absorption", span=3); row += 1
    row = period_header(ws, row)

    grid_row(ws, row, "Cumulative Units Delivered",
             lambda p, C, Cprev: f"IF({p}<DeliverStart,0,MIN(TotalUnits,({p}-DeliverStart+1)*UnitsDeliveredPerMo))",
             fmt=FMT_NUM0, total=None, key="delivered", sheet=SH_LEASE)
    row += 1
    R_del = r(SH_LEASE, "delivered")
    grid_row(ws, row, "Cumulative Units Leased",
             lambda p, C, Cprev: (f"IF({p}<DeliverStart,0,MIN(StabLeasedTarget,{C}${R_del},"
                                  f"({p}-DeliverStart+1)*UnitsLeasedPerMo))"),
             fmt=FMT_NUM0, total=None, key="leased", sheet=SH_LEASE)
    row += 1
    R_leased = r(SH_LEASE, "leased")
    grid_row(ws, row, "Physical Occupancy",
             lambda p, C, Cprev: f"IF(TotalUnits=0,0,{C}${R_leased}/TotalUnits)",
             fmt=FMT_PCT1, total=None, key="occ", sheet=SH_LEASE)
    row += 2

    # Reconciliation
    write_subhead(ws, row, "Lease-Up Reconciliation", span=3); row += 1
    leased_rng = f"{FIRST_COL_LETTER}{R_leased}:{LAST_COL_LETTER}{R_leased}"
    del_rng = f"{FIRST_COL_LETTER}{R_del}:{LAST_COL_LETTER}{R_del}"
    label(ws, row, "Occupancy at Stabilization − target", col=1)
    put_formula(ws, row, 2, f"INDEX({leased_rng},StabMonth)-StabLeasedTarget", fmt=FMT_NUM0)
    reg.add("Chk_LeaseUpReached", SH_LEASE, f"B{row}")
    set_row(SH_LEASE, "Chk_LeaseUpReached", row)
    row += 1
    label(ws, row, "Units delivered at Stabilization", col=1)
    put_formula(ws, row, 2, f"INDEX({del_rng},StabMonth)", fmt=FMT_NUM0)
    reg.add("Chk_UnitsDelivered", SH_LEASE, f"B{row}")
    set_row(SH_LEASE, "Chk_UnitsDelivered", row)
    ws.freeze_panes = "D3"
