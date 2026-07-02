# Institutional Multifamily Ground-Up Development Model

A single, self-contained Excel workbook (`model.xlsx`) that underwrites a
ground-up multifamily development. Fully monthly, hold period toggleable from
**1 to 360 months**, with construction financing that automatically either
refinances into permanent debt or is retired by a sale depending on hold length.

The workbook is generated programmatically (Python + `openpyxl`), recalculated
headlessly, error-scanned, and reconciled cell-for-cell against an independent
pure-Python "shadow" model.

---

## 1. How it is built and proven correct

| Script | Purpose |
|---|---|
| `generator/assumptions.py` | **Single source of truth** for every input (scalars, milestones, unit matrix, budget lines). Imported by both the Excel generator and the shadow model, so no input is defined twice. |
| `generator/common.py` | Styling / number-format / named-range primitives (blue input · black formula · green link conventions). |
| `generator/layout.py` | Shared row registry + the monthly-grid writer + cross-sheet reference helpers. |
| `generator/phase_01…15_*.py` | One module per build phase. Each fills its section of the connected calculation graph and registers its rows / named ranges. |
| `generator/build.py` | Orchestrates the phases in order, writes `model.xlsx`, saves per-phase snapshots (`build/model_phaseNN.xlsx`), dumps the named-range table. |
| `generator/shadow.py` | **Authoritative independent re-computation** in pure Python — every mechanic, month by month. |
| `generator/validate.py` | Recalculates `model.xlsx`, scans every cell for Excel error strings, reconciles ~30 headline named ranges to the shadow to a tight tolerance. |
| `generator/scenario_test.py` | Recomputes the workbook under 10 input scenarios (hold either side of maturity, No-Cash-Out, reassess off, funding order, reserve override) and reconciles each to the shadow. |

### Recalculation engine note
The build container's **LibreOffice is non-functional for headless conversion**
(`soffice --headless --convert-to …` fails to load *any* source file, including a
one-line text file — a broken filter component, not a workbook problem). The
recalculation engine is therefore [`formulas`](https://pypi.org/project/formulas/),
a pure-Python Excel calculator that parses the `.xlsx`, builds the dependency
graph and computes every cell. The independent shadow model remains the
authoritative numeric check; `formulas` provides the in-Excel recalculation and
error-string scan the workflow calls for.

Run the full pipeline:
```bash
cd generator
python3 build.py          # writes ../model.xlsx (+ --snap for per-phase files)
python3 validate.py       # recalc + error scan + shadow reconciliation
python3 scenario_test.py  # 10-scenario stress test
```

---

## 2. Workbook structure (14 tabs)

| Tab | Purpose |
|---|---|
| **Dashboard** | Auto-updating summary: returns (levered/unlevered IRR & MOIC, gross/net, static), development metrics, sources & uses, timeline, sizing tests, refi-or-sell status, error count. |
| **Inputs** | Every scalar assumption as a blue input; each is a workbook named range. |
| **UnitMatrix** | Residential (market-rate + low-income) and retail matrices; derived unit counts, NRSF, rent roll. |
| **Budget** | Development budget line items (incl. 10+ user-definable lines), each with basis / phase / cost curve / modifier and a modeled total. |
| **Timeline** | Milestone table → monthly period grid, key timing (construction maturity, delivery, stabilization, refi, disposition) and phase flags. |
| **DevSpend** | Cost-curve engine — each line spread monthly by its curve; sums reconcile exactly to line totals. |
| **ConLoan** | Construction loan sizing (min LTC / debt yield / DSCR) + monthly draw schedule + sequential interest reserve + loan totals. |
| **LeaseUp** | Monthly unit delivery, absorption, occupancy. |
| **Operating** | Monthly revenue & expense proforma, lease-up expense ramp/elevation, NOI before/after tax, property-tax engine, below-NOI, operating reserve, static (no-growth) NOI. |
| **PermLoan** | Refinance sizing (min LTV / debt yield / DSCR), No-Cash-Out solve, monthly amortization. |
| **Disposition** | Forward-12-month NOI exit value, reassess-on-sale closed form, sale costs, net proceeds. |
| **CashFlow** | Levered & unlevered monthly cash flows + reconciliation aggregates. |
| **Returns** | IRR & MOIC (gross/net), development spread, ROC, TDC metrics, static reference. |
| **Diagnostics** | 22 plain-English integrity checks + single "total errors found" count. |

Named ranges are listed in `build/named_ranges.txt` (245 of them).

---

## 3. Key mechanics (how circularity is avoided)

- **Interest reserve** is resolved by a **sequential monthly draw schedule** — each
  month's interest is computed on the prior month's ending balance
  (`ConLoan` rows), so the reserve is derived row-by-row, never circular.
  Construction sizing is on a cost basis *excluding* capitalized interest, so the
  reserve never feeds back into the commitment.
- **Refi-or-sell switch** (`RefiFlag`) is driven purely by hold length vs.
  construction maturity: `Disposition ≤ Maturity → sell` (perm loan zeroed,
  construction retired at sale); `Disposition > Maturity → refinance` at maturity,
  then hold to sale.
- **Reassess-on-sale** exit value is solved in **closed form**
  `Exit = Fwd NOI-before-tax / (blended exit cap + tax-rate-on-value)`, avoiding
  the value↔tax circularity. Iterative calculation stays OFF.
- **Operating reserve** is model-calculated as the peak cumulative lease-up cash
  shortfall × a coverage multiple (floored by a months-of-opex minimum).
- **No-Cash-Out** solves the perm loan in closed form so net cash at refi = 0
  (`PermLoan = ConPayoff / (1 − points%)`, capped at the sized maximum).

---

## 4. Base-case walkthrough

See `logs/validation_log.md` for the reconciliation record and base-case figures.
