"""
validate.py — recalculate the workbook with the pure-Python `formulas` engine,
scan every calculated cell for Excel error strings, and reconcile key named-range
outputs against the independent Python shadow model.

(LibreOffice headless conversion is non-functional in this container — it fails
to load any source file — so `formulas`, a pure-Python Excel calculator, is used
as the recalculation engine.  The shadow model remains the authoritative numeric
check.)

Usage: python3 validate.py [path_to_xlsx]
"""

import os
import sys
import warnings

import numpy as np
import openpyxl
import formulas
import shadow

warnings.filterwarnings("ignore")
ERROR_STRINGS = ("#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NUM!", "#NULL!")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _scalar(v):
    """Extract a python scalar from a formulas solution value."""
    try:
        val = v.value
    except AttributeError:
        val = v
    if isinstance(val, np.ndarray):
        flat = val.ravel()
        val = flat[0] if flat.size else None
    return val


def calc(path):
    xl = formulas.ExcelModel().loads(path).finish()
    sol = xl.calculate()
    return sol


def scan_errors(sol):
    hits = []
    for key, v in sol.items():
        if "!" not in key:
            continue
        val = _scalar(v)
        if isinstance(val, str):
            for e in ERROR_STRINGS:
                if e in val:
                    hits.append((key, val))
                    break
    return hits


def named_key(wb, name, bookname):
    dn = wb.defined_names.get(name)
    if dn is None:
        return None
    for title, coord in dn.destinations:
        cell = coord.replace("$", "")
        return f"'[{bookname}]{title.upper()}'!{cell}"
    return None


def reconcile(sol, wb, s, bookname):
    checks = [
        ("TotalUnits", s["total_units"], 0.5),
        ("NRSF", s["nrsf"], 0.5),
        ("DirectCosts", s["direct_costs"], 1.0),
        ("DevFee", s["dev_fee"], 1.0),
        ("ConTaxTotal", s["con_tax_total"], 5.0),
        ("BudgetLand", s["land_total"], 1.0),
        ("BudgetSoft", s["soft_total"], 1.0),
        ("BudgetHard", s["hard_total"], 1.0),
        ("ConCostCommit", s["con_cost_commit"], 5.0),
        ("InterestReserveAccrued", s["accrued_int"], 50.0),
        ("ConLoanTotal", s["con_loan_total"], 50.0),
        ("ConPayoffBal", s["con_payoff_bal"], 50.0),
        ("OpReserve", s["op_reserve"], 100.0),
        ("TDC_Total", s["tdc"], 100.0),
        ("EquityTotal", s["equity_total"], 100.0),
        ("StabNOI", s["stab_noi_annual"], 100.0),
        ("StabNOIbt", s["stab_noi_bt"], 100.0),
        ("TaxYear1", s["tax_year1"], 50.0),
        ("AssessedValueStab", s["assessed_value_stab"], 500.0),
        ("Year1NOI", s["year1_noi_annual"], 100.0),
        ("ValueAtRefi", s["value_at_refi"], 1000.0),
        ("PermLoan", s["perm_loan"], 500.0),
        ("RefiCashOut", s["refi_cash_out"], 500.0),
        ("PermPayoffBal", s["perm_payoff_bal"], 500.0),
        ("ExitValue", s["exit_value"], 1000.0),
        ("BlendedExitCap", s["blended_cap"], 1e-4),
        ("NetSale", s["net_sale"], 1000.0),
        ("YieldOnCost", s["yield_on_cost"], 1e-4),
        ("LevMOIC", s["lev_moic"], 2e-3),
        ("UnlevMOIC", s["unlev_moic"], 2e-3),
        ("LevIRR", s["lev_irr"], 1e-3),
        ("UnlevIRR", s["unlev_irr"], 1e-3),
        ("ErrorsFound", 0, 0.5),
    ]
    results = []
    for name, sval, tol in checks:
        key = named_key(wb, name, bookname)
        xl = _scalar(sol.get(key)) if key else None
        ok, detail = False, ""
        if key is None:
            detail = "named range missing"
        elif xl is None:
            detail = "no solution value"
        elif isinstance(xl, str):
            detail = f"excel={xl!r} shadow={sval}"
            ok = (sval is None)
        elif sval is None:
            ok, detail = True, f"excel={float(xl):,.4f} shadow=n/a"
        else:
            diff = abs(float(xl) - float(sval))
            ok = diff <= tol
            detail = f"excel={float(xl):,.4f} shadow={float(sval):,.4f} diff={diff:,.4f}"
        results.append((name, ok, detail))
    return results


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "model.xlsx")
    bookname = os.path.basename(src)
    print(f"Recalculating {src} with the `formulas` engine ...")
    sol = calc(src)
    print(f"Calculated {len(sol)} cells.")

    hits = scan_errors(sol)
    print(f"\nError-string scan: {len(hits)} cell(s) with Excel errors")
    for key, v in hits[:50]:
        print(f"   {key}: {v}")

    wb = openpyxl.load_workbook(src)
    s = shadow.compute()
    print("\nShadow reconciliation:")
    results = reconcile(sol, wb, s, bookname)
    npass = sum(1 for _, ok, _ in results if ok)
    for name, ok, detail in results:
        print(f"   [{'OK  ' if ok else 'FAIL'}] {name:24s} {detail}")
    print(f"\n{npass}/{len(results)} reconciliations pass; {len(hits)} error cells")
    return len(hits) == 0 and npass == len(results)


if __name__ == "__main__":
    sys.exit(0 if main() else 1)
