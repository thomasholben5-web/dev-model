"""
scenario_test.py — recompute the Excel workbook under different input scenarios
(hold length either side of construction maturity, No-Cash-Out, reassess-on-sale,
funding order) using the `formulas` engine's input override, and reconcile each
against the shadow model with the same overrides.  Confirms zero error cells and
matching key outputs in every scenario.
"""

import os
import sys
import warnings

import numpy as np
import openpyxl
import formulas
import shadow

warnings.filterwarnings("ignore")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "model.xlsx")
BOOK = "model.xlsx"
ERRS = ("#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NUM!", "#NULL!")

OUT_KEYS = ["RefiFlag", "ExitValue", "NetSale", "PermLoan", "ConPayoffBal",
            "LevMOIC", "UnlevMOIC", "LevIRR", "UnlevIRR", "RefiCashOut", "ErrorsFound"]
SHADOW_MAP = {
    "RefiFlag": lambda s: 1 if s["refi"] else 0,
    "ExitValue": lambda s: s["exit_value"],
    "NetSale": lambda s: s["net_sale"],
    "PermLoan": lambda s: s["perm_loan"],
    "ConPayoffBal": lambda s: s["con_payoff_bal"],
    "LevMOIC": lambda s: s["lev_moic"],
    "UnlevMOIC": lambda s: s["unlev_moic"],
    "LevIRR": lambda s: s["lev_irr"],
    "UnlevIRR": lambda s: s["unlev_irr"],
    "RefiCashOut": lambda s: s["refi_cash_out"],
    "ErrorsFound": lambda s: 0,
}
TOL = {"ExitValue": 2000, "NetSale": 2000, "PermLoan": 1000, "ConPayoffBal": 100,
       "RefiCashOut": 1000, "LevMOIC": 3e-3, "UnlevMOIC": 3e-3, "LevIRR": 2e-3,
       "UnlevIRR": 2e-3, "RefiFlag": 0.1, "ErrorsFound": 0.5}

SCENARIOS = [
    ("Base (hold 84, refi)", {}),
    ("Sell during construction (hold 30)", {"HoldPeriodMonths": 30}),
    ("Sell at maturity (hold 42)", {"HoldPeriodMonths": 42}),
    ("Refi just past maturity (hold 43)", {"HoldPeriodMonths": 43}),
    ("Long hold (hold 180)", {"HoldPeriodMonths": 180}),
    ("Max hold (hold 360)", {"HoldPeriodMonths": 360}),
    ("No-Cash-Out on", {"NoCashOut": 1}),
    ("Reassess-on-sale off", {"ReassessOnSale": 0}),
    ("Pari-passu funding", {"FundingOrderPari": 1}),
    ("Lender reserve override", {"LenderReqReserveOverride": 9000000}),
]


def cell_key(wb, name):
    dn = wb.defined_names.get(name)
    for title, coord in dn.destinations:
        return f"'[{BOOK}]{title.upper()}'!{coord.replace('$','')}"


def sval(v):
    try:
        v = v.value
    except AttributeError:
        pass
    if isinstance(v, np.ndarray):
        f = v.ravel()
        v = f[0] if f.size else None
    return v


def main():
    print("Compiling workbook once ...")
    xl = formulas.ExcelModel().loads(SRC).finish()
    wb = openpyxl.load_workbook(SRC)
    input_keys = {k: cell_key(wb, k) for k in
                  set().union(*[d.keys() for _, d in SCENARIOS]) if k}
    out_keys = {k: cell_key(wb, k) for k in OUT_KEYS}

    all_ok = True
    for name, ov in SCENARIOS:
        inputs = {input_keys[k]: v for k, v in ov.items()}
        sol = xl.calculate(inputs=inputs) if inputs else xl.calculate()
        # error scan
        nerr = 0
        for key, v in sol.items():
            x = sval(v)
            if isinstance(x, str) and any(e in x for e in ERRS):
                nerr += 1
        s = shadow.compute(ov)
        line_ok = (nerr == 0)
        details = []
        for k in OUT_KEYS:
            xlv = sval(sol.get(out_keys[k]))
            shv = SHADOW_MAP[k](s)
            if isinstance(xlv, str):          # e.g. IRR "n/a"
                ok = shv is None
            elif shv is None:
                ok = True
            else:
                ok = abs(float(xlv) - float(shv)) <= TOL.get(k, 1.0)
            if not ok:
                details.append(f"{k}(xl={xlv},sh={shv})")
            line_ok = line_ok and ok
        all_ok = all_ok and line_ok
        status = "PASS" if line_ok else "FAIL"
        extra = "" if line_ok else "  << " + "; ".join(details)
        print(f"  [{status}] {name:38s} errcells={nerr}{extra}")
    print("\nALL SCENARIOS PASS" if all_ok else "\nSOME SCENARIOS FAILED")
    return all_ok


if __name__ == "__main__":
    sys.exit(0 if main() else 1)
