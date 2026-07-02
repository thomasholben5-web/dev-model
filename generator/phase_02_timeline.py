"""
Phase 02 — Timeline engine.

Builds the milestone table (Start | Duration | End, every event milestone-
relative and never a hardcoded date), the key timing derivations (construction
start / maturity, delivery, stabilization, disposition, refi-or-sell switch),
and the monthly period grid with per-milestone phase flags.
"""

from common import (write_title, write_section, write_subhead, label, put_input,
                    put_formula, put_text, set_col_widths, FMT_NUM0, FMT_MONTHS,
                    FMT_DATE, font_subhead, ALIGN_R, FILL_TOTAL)
from assumptions import MILESTONES, ms_slug
from layout import (SH_TIME, set_row, r, period_header, date_header, grid_row,
                    mref)


def build(wb, reg):
    ws = wb[SH_TIME]
    set_col_widths(ws, {"A": 30, "B": 12, "C": 12, "D": 12})
    ws.sheet_view.showGridLines = False
    write_title(ws, 1, "Timeline Engine — Milestones & Monthly Period Grid")

    # ---- Milestone table --------------------------------------------------
    hr = 3
    for j, h in enumerate(["Milestone", "Start (mo)", "Duration (mo)", "End (mo)"], start=1):
        c = ws.cell(row=hr, column=j, value=h)
        c.font = font_subhead(); c.alignment = ALIGN_R if j > 1 else None
    first = hr + 1
    row = first
    slug_row = {}
    for (name, start, dur) in MILESTONES:
        slug = ms_slug(name)
        put_text(ws, row, 1, name)
        if name == "Disposition":
            put_formula(ws, row, 2, "HoldPeriodMonths", fmt=FMT_MONTHS, link=True)
            put_input(ws, row, 3, dur, fmt=FMT_MONTHS)
        else:
            put_input(ws, row, 2, start, fmt=FMT_MONTHS)
            put_input(ws, row, 3, dur, fmt=FMT_MONTHS)
        put_formula(ws, row, 4, f"B{row}+C{row}-1", fmt=FMT_MONTHS)
        reg.add(f"MS_{slug}_Start", SH_TIME, f"B{row}")
        reg.add(f"MS_{slug}_End", SH_TIME, f"D{row}")
        set_row(SH_TIME, f"MS_{slug}_Start", row)
        slug_row[slug] = row
        row += 1
    last = row - 1
    # Ranges for INDEX/MATCH lookup by milestone name
    reg.add("MilestoneNames", SH_TIME, f"$A${first}:$A${last}")
    reg.add("MilestoneStart", SH_TIME, f"$B${first}:$B${last}")
    reg.add("MilestoneEnd", SH_TIME, f"$D${first}:$D${last}")

    # ---- Key timing derivations ------------------------------------------
    row += 1
    row = write_subhead(ws, row, "Key Timing (derived)", span=2)
    def derive(lbl, formula, name, fmt=FMT_MONTHS):
        label(ws, row_ref[0], lbl, col=1)
        put_formula(ws, row_ref[0], 2, formula, fmt=fmt, link=True, bold=True)
        reg.add(name, SH_TIME, f"B{row_ref[0]}")
        set_row(SH_TIME, name, row_ref[0])
        row_ref[0] += 1
    row_ref = [row]
    derive("Construction Start", "MS_Construction_Start", "ConStart")
    derive("Construction Maturity", "ConStart+ConLoanTermMonths-1", "ConMaturity")
    derive("1st Unit Delivery", "MS_1st_Unit_Delivery_Start", "DeliverStart")
    derive("Stabilization Month", "MS_Stabilization_Start", "StabMonth")
    derive("Disposition Month", "HoldPeriodMonths", "DispoMonth")
    derive("Growth Start Month",
           "(YEAR(GrowthStartDate)-YEAR(InceptionDate))*12+MONTH(GrowthStartDate)-MONTH(InceptionDate)+1",
           "GrowthStartMonth")
    derive("Refi Flag (1=refi,0=sell)", "IF(DispoMonth>ConMaturity,1,0)", "RefiFlag", fmt=FMT_NUM0)
    derive("Refi Month",
           "IF(RefiFlag=1,MIN(ConMaturity,IF(RefiMonthOverride>0,RefiMonthOverride,ConMaturity)),0)",
           "RefiMonth")
    derive("Con Payoff Month", "IF(RefiFlag=1,RefiMonth,DispoMonth)", "ConPayoffMonth")

    # ---- Monthly period grid ---------------------------------------------
    row = row_ref[0] + 1
    row = write_section(ws, row, "Monthly Period Grid", span=3)
    row = period_header(ws, row)
    set_row(SH_TIME, "hdr", row - 1)
    row = date_header(ws, row)
    set_row(SH_TIME, "date", row - 1)

    # Per-milestone flag rows
    for (name, start, dur) in MILESTONES:
        slug = ms_slug(name)
        srow = r(SH_TIME, f"MS_{slug}_Start")
        # flag = 1 if start<=p<=end
        def make(srow):
            return lambda p, C, Cprev: f"IF(AND({p}>=$B${srow},{p}<=$D${srow}),1,0)"
        row = grid_row(ws, row, f"Flag: {name}", make(srow), fmt=FMT_NUM0,
                       total=None, key=f"flag_{slug}", sheet=SH_TIME)

    # Operations-active flag (delivery .. disposition)
    row = grid_row(ws, row, "Flag: Operations Active",
                   lambda p, C, Cprev: f"IF(AND({p}>=DeliverStart,{p}<=DispoMonth),1,0)",
                   fmt=FMT_NUM0, total=None, key="OpsActive", sheet=SH_TIME)
    # Construction-loan-active flag (con start .. payoff month)
    row = grid_row(ws, row, "Flag: Con Loan Active",
                   lambda p, C, Cprev: f"IF(AND({p}>=ConStart,{p}<=ConPayoffMonth),1,0)",
                   fmt=FMT_NUM0, total=None, key="ConActive", sheet=SH_TIME)
    # Perm-loan-active flag (refi+1 .. disposition, only if refi)
    row = grid_row(ws, row, "Flag: Perm Loan Active",
                   lambda p, C, Cprev: f"IF(AND(RefiFlag=1,{p}>RefiMonth,{p}<=DispoMonth),1,0)",
                   fmt=FMT_NUM0, total=None, key="PermActive", sheet=SH_TIME)
    ws.freeze_panes = "D3"
