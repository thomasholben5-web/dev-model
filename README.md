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
- **Growth clock** is anchored to a `GrowthStartDate` input (default month 2,
  overridable) and compounds **monthly** across all growth series.
- **Year-1 property tax is dynamic** (non-circular): assessed value =
  stabilized NOI-before-tax / (cap + effective rate), tax = value × effective
  rate, with a manual override (`TaxYear1Override`). The disposition
  reassessment uses the same effective rate for consistency.
- **Construction-period property tax is dynamic** — assessed on land +
  cumulative hard-cost put in place × effective rate, monthly from initial
  closing to stabilization, **capitalized** as a development cost (operating
  tax begins at stabilization, so there is no double count).
- **Operating reserve** is funded at closing and **released to equity at
  stabilization** (recycled, not trapped); the actual lease-up shortfall is
  funded through operating cash flow.
- Every monthly grid carries a **calendar-month reference row** (real dates via
  non-volatile `EDATE`) beneath the period numbers.

---

## 4. Base-case walkthrough

See `logs/validation_log.md` for the reconciliation record and base-case figures.

---

## 5. Enhancement pass (`Claude_Dev_Model.xlsx`)

The user edited `build/model_phase15.xlsx` into **`Claude_Dev_Model.xlsx`** and asked
for the NOI proforma to be formatted like a lender stabilized proforma, plus a general
review. That pass is generated by `generator/enhance.py` (idempotent; re-runnable):

```bash
python3 generator/enhance.py Claude_Dev_Model.xlsx build/model_enhanced.xlsx
```

- **Phase A — operations fix.** The `Lease-Up` milestone duration (`Timeline!D11`) was a
  fractional month, so `INDEX(…,StabMonth)` truncated and stabilization landed one month
  before absorption completed — the *"reaches stabilized occupancy"* diagnostic failed
  (`Total Errors Found = 1`). New closed form
  `=MAX(CEILING(StabLeasedTarget/UnitsLeasedPerMo,1),CEILING(StabLeasedTarget/UnitsDeliveredPerMo,1))`
  makes stabilization land on the whole month absorption actually completes →
  `StabMonth = 26`, all diagnostics pass.
- **Phase B — new `Proforma` tab** (after Dashboard). Stabilized Year-One Operating
  Proforma with `% of / amt` · `$ / SF /mo` · `$ / Unit /mo` · `$ / Month` ·
  `$ / Unit /yr` · `$ / Year` bases. `$ / Year` is the forward-12-month stabilized sum;
  every other basis derives from it. The driver column links to `Inputs` (single source of
  truth). A reconciliation block ties OpCF / below-line split / NOI back to the `Operating`
  engine and feeds the master `Diagnostics` error count.
- **Phase C — formatting/navigation.** Tab colours (output/input/engine/validation groups),
  frozen label panes on the wide monthly grids, gridlines off on the Proforma.

Validation record: `logs/validation.log` (5,829 formulas compile clean; error-string scan
0; shadow reconciliation deltas 0; 251 named ranges preserved).

---

## 6. Remediation pass (`generator/remediate.py`)

A nine-item defect remediation applied directly to `Claude_Dev_Model.xlsx` (the user's
"The Verde" deal — a different deal and column layout than the Meridian generator, so the
workbook is patched, not regenerated). Validation record: `logs/remediation_validation.log`.

| # | Fix | Effect |
|---|---|---|
| 1 | **Untrended proforma drives dev metrics.** New point-in-time untrended stabilized proforma (growth=1, occ=StabOcc); `StabNOI_Untrended`/`StabNOIbt_Untrended` drive YOC / ROC-stab / dev-spread / debt-yield. Trended NOI still drives CF/IRR/refi/disposition. Proforma tab rebuilt: untrended headline + trended reference column + real driver column (FIX 9). | YOC 5.055%(trended)→untrended basis |
| 2 | **Refinance engine live.** Construction term 480→36 mo (+ 12–60 data validation); refi branch activates. | Refi $13.35M @mo 45; TDC $21.8M→$16.9M (stops capitalizing a decade of interest); spread −20→+69 bps |
| 3 | **Property-tax transparency + lease-up toggle.** Derived effective rate + tax/unit surfaced on Inputs; `TaxCapitalizeLeaseUp` toggle (default = capitalize pre-stab tax). | base case unchanged; sanity warning flags low tax/unit |
| 4 | **Static block growth-independent.** `StaticTaxYear1` decoupled from the grown engine. | static IRR now truly no-growth |
| 5 | **Diagnostics prove claims.** Per-sheet workbook-wide `ISERROR` scan; non-trivial two-branch refi checks; economic-sanity warnings; dead-input audit; `WarningsFound` count. | 21 recon PASS + warnings |
| 6 | **Dead inputs wired by name.** Cap adjustments→exit cap, ModelUnits→vacancy (both engines), InvestmentType/Location→Dashboard, NetAcres/FARRatio→site metrics, RetailIncomeMo, DevFeeThirdPartySplit, LeaseUpIncomeOffset→reserve. | 0 dead inputs |
| 7 | **Hardcoded constants→Inputs.** 11 milestone durations + the 13-month delivery lag moved to a new Inputs section. | single source of truth |
| 8 | **Debt yield on active loan.** `DebtYieldStab` = untrended NOI ÷ perm balance when refinanced, else construction. | 7.53% (perm) |
| 9 | **Proforma driver column.** Real per-line drivers (avg rent, vacancy %, per-unit $, eff tax rate). | (folded into #1) |

Both refi branches, the No-Cash-Out toggle, all 21 reconciliation checks, and the native
50,028-cell error scan reconcile with **zero errors**.
