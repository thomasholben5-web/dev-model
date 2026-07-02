"""
build.py — orchestrates the phased construction of the workbook.

Creates the workbook skeleton, runs each phase module in order (each fills its
section of the connected calculation graph and registers its rows / named
ranges), saves a versioned snapshot per phase, and writes the final workbook.

Run:  python3 build.py            # build all phases, write model.xlsx
      python3 build.py --snap     # also save build/model_phaseNN.xlsx per phase
"""

import os
import sys
import importlib

from openpyxl import Workbook
from common import NameRegistry, set_col_widths
from layout import SHEET_ORDER

BUILD_DIR = os.path.join(os.path.dirname(__file__), "..", "build")
OUT_XLSX = os.path.join(os.path.dirname(__file__), "..", "model.xlsx")

PHASES = [
    ("phase_01_architecture", "Workbook architecture, inputs, unit matrix, budget"),
    ("phase_02_timeline",     "Timeline engine"),
    ("phase_03_costcurve",    "Development budget & cost-curve engine"),
    ("phase_04_conloan",      "Construction loan draws & sizing"),
    ("phase_05_interest",     "Interest reserve (sequential)"),
    ("phase_06_leaseup",      "Unit delivery & lease-up engine"),
    ("phase_07_operating",    "Stabilized operating model"),
    ("phase_08_proptax",      "Property tax engine"),
    ("phase_09_permloan",     "Permanent loan / refinance engine"),
    ("phase_10_longterm",     "Long-term operating cash flow"),
    ("phase_11_disposition",  "Disposition engine"),
    ("phase_12_cashflows",    "Levered & unlevered cash flows"),
    ("phase_13_returns",      "Return calculations"),
    ("phase_14_dashboard",    "Dashboard"),
    ("phase_15_diagnostics",  "Diagnostics & validation"),
]


def new_workbook():
    wb = Workbook()
    # create sheets in canonical order; remove default
    default = wb.active
    for name in SHEET_ORDER:
        wb.create_sheet(title=name)
    wb.remove(default)
    return wb


def main():
    snap = "--snap" in sys.argv
    os.makedirs(BUILD_DIR, exist_ok=True)
    wb = new_workbook()
    reg = NameRegistry()

    for i, (mod_name, desc) in enumerate(PHASES, start=1):
        mod = importlib.import_module(mod_name)
        mod.build(wb, reg)
        reg.apply(wb)                      # (re)apply names after each phase
        if snap:
            path = os.path.join(BUILD_DIR, f"model_phase{i:02d}.xlsx")
            wb.save(path)
        print(f"  phase {i:02d} built: {desc}")

    reg.apply(wb)
    wb.save(OUT_XLSX)
    # dump named ranges for the audit log
    with open(os.path.join(BUILD_DIR, "named_ranges.txt"), "w") as fh:
        for name, ref in reg.dump().items():
            fh.write(f"{name}\t{ref}\n")
    print(f"Workbook written: {OUT_XLSX}  ({len(reg.dump())} named ranges)")


if __name__ == "__main__":
    main()
