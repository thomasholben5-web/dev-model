# Claude Code Build Prompt — Institutional Multifamily Ground-Up Development Model (Excel)

## Role, objective & governing standard
You are Claude Code, working in an agentic environment with a filesystem, shell, and Python. Build a single, self-contained **Excel workbook (.xlsx)** that underwrites a ground-up multifamily **development**: fully monthly, hold period toggleable from **1 to 360 months**, with construction financing that either refinances into permanent debt or is retired by a sale, depending on hold length.

Build this workbook **as if it will be audited by a Big Four accounting firm and maintained for the next 10 years by analysts who did not build it.** Prioritize transparency, modularity, auditability, maintainability, and calculation performance over minimizing worksheet count or formula count. Use clean, standard-industry labels — no firm-specific terminology, proprietary account codes, or house-style naming.

The model must be **fully dynamic**: every output flows from input cells, no hardcoded calculated results exist anywhere, and the workbook recomputes correctly when any single assumption changes.

## How to build it (Claude Code workflow)
Generate the workbook programmatically and validate it continuously. Do not hand me a workbook you have not proven correct.

- **Generation:** Use Python with `openpyxl` to author formulas, Excel Tables, named ranges, number formats, and data validation. Structure the generator modularly — one module per build phase (`phase_01_...py` … `phase_15_...py`) orchestrated by a `build.py`, with a single source-of-truth assumptions module so no input is defined twice.
- **Recalculation:** `openpyxl` writes formulas but does not compute them. After each phase, force evaluation by recalculating the workbook headlessly with LibreOffice (`soffice --headless --convert-to xlsx --calc`), then re-open with `openpyxl(data_only=True)` and scan **every cell** for Excel error strings (`#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, `#NAME?`, `#NUM!`, `#NULL!`). Use LibreOffice ≥ 24.x for best support of modern functions.
- **Independent shadow reconciliation:** For each phase, also compute the same math independently in pure Python (a "shadow model") and reconcile the Excel-recalculated values against the Python values to a tight tolerance (e.g., 1e-6 relative). The shadow model is the authoritative numeric check; LibreOffice recalc is for error detection and cross-check. Where the recalc engine doesn't fully support a modern function, the shadow model must still confirm the intended result.
- **Phase gating:** Each phase must be **fully functional and internally reconciled before the next phase begins.** No placeholder formulas, temporary hardcoded values, or incomplete mechanics may remain when advancing. Save a versioned workbook per phase (`build/model_phase07.xlsx`) and log that phase's validation results.
- **Audit trail:** Produce a short `README`/audit log documenting workbook structure, tab purposes, every named range, key formulas, and each phase's pass/fail validation record.

## Excel engineering standards
**Do not use:** VBA · macros · Power Query · external workbook links · circular references · iterative calculation mode · hidden calculations · hardcoded calculated outputs · merged cells within calculation areas · volatile functions unless absolutely necessary (`OFFSET`, `INDIRECT`, `TODAY`, `RAND`, `NOW`, `CELL`, `INFO`, etc.).

**Prefer:** Excel Tables · structured references · named ranges · `LET()` where it aids clarity · `XLOOKUP()` · `INDEX/MATCH` where it improves performance · dynamic arrays only where compatible with professional modeling practice.

Optimize workbook calculation speed. Avoid unnecessarily long formulas by using helper calculations where appropriate.

### Formula design standards
Every calculation must:
- Reference assumptions (never a buried input).
- Contain **no hardcoded numeric constants** unless mathematically required (e.g., 12 months/year, 365 days).
- Remain readable and auditable.
- Avoid deeply nested `IF` statements where helper rows improve clarity.
- Be modular rather than monolithic.

## Core mechanics (apply across phases)
- **Monthly periodicity** for every mechanic: development spend, construction draws, interest accrual, interest reserve, lease-up, operations, refinance, and disposition.
- **Hold-period toggle (1–360 months)** drives the disposition month; the cash-flow grid, valuation, and all returns extend and recompute automatically.
- **Milestone-driven timeline** is the model's clock. Each milestone has Start Date, Duration (months), End Date; every event is positioned relative to milestones, never hardcoded dates. Support: Open Escrow, Due Diligence, Entitlement, Pre-Construction, Initial Closing, Construction, 1st Unit Delivery, Lease-Up, Stabilization, Recapitalization/Refinance, Disposition.
- **No circularity by design.** Resolve the interest reserve through a **sequential monthly draw schedule** — each month's interest is computed on the prior month's ending balance, so the reserve is derived row-by-row, never circular. Solve tax-reassessed exit value **algebraically** (closed form) rather than iteratively (see Disposition). Iterative calculation must remain OFF.

### Refinance-or-sell logic (critical)
The construction loan has a term in months → a maturity month = construction start + term. Compare the disposition month (from the hold toggle) to construction maturity:
- **Disposition month ≤ construction maturity:** the asset is sold before the construction loan expires → **no refinance.** The construction loan is retired directly from sale proceeds at disposition; the permanent loan is inactive (zeroed).
- **Disposition month > construction maturity:** the asset is held past expiration → **refinance is assumed.** The permanent loan takes out the construction loan at the recapitalization/refi month (defaulting to construction maturity, capped at maturity), the asset is held to disposition, and the permanent loan is retired from sale proceeds at disposition.

This switch must be automatic and driven entirely by hold length vs. construction term. Both branches (sell-during-construction and refi-then-hold) must reconcile cleanly.

### Operating expense behavior during lease-up
Make both dynamic and user-controllable:
- **Occupancy-driven variable expenses** (contract services, turnover, repairs & maintenance, utilities): during lease-up these run at a **user-input % of the stabilized baseline**, ramping to 100% at stabilization. Expose the ramp % as an input.
- **Elevated lease-up expenses** (advertising, payroll): during lease-up these run **above** the stabilized baseline by a **user-input elevation %**, normalizing to baseline at stabilization. Expose the elevation % as an input.

### Operating reserve (model-calculated)
Derive the operating reserve from the model — size it to the **peak cumulative operating cash shortfall during lease-up** (the maximum funding need before the property covers its own costs), with an optional coverage multiple / months-of-buffer input. Do not treat it as a raw hardcoded input.

### User-definable development budget line items
Within the development budget, provide **at least 10 user-definable "Other Soft Cost / Other Hard Cost" line items** the user can populate with a label and cost that flow through to the cost-curve engine, draw schedule, budget totals, and diagnostics exactly like named line items. Each user line must carry the same attributes as a standard line (cost basis toggle: $ / $-per-unit / $-per-SF; Phase; Timing; Cost Curve). Design the count to be easily extensible.

## Input layer (organize into clearly labeled sections)
Keep every field driveable, at roughly this granularity:

- **Property Details** — Investment Type, Property Name, City, State, Market, Product Type, Location, Inception Date.
- **Unit Delivery & Lease-Up** — lease-up income treatment, unit delivery start, units delivered/month, units leased/month, lease-up expense ramp % and elevation % inputs.
- **Investment Details** — Hold Period (1–360 mo toggle), Cap Rate & Adjustments (market/entry cap, location/size/aging adjustments, exit cap – residential, exit cap – retail), default growth rates (rent, low-income rent, expense, utility, tax).
- **Property Taxes** — operating-period rates (levy %, assessment %, value-adjustment factor, combined rate), post-stabilization value method (income vs. value approach, Year-1 tax), **reassess-on-sale toggle**, disposition-period rates.
- **Timing Milestone Assumptions** — the milestone table (Milestone | Start | Duration (mos) | End).
- **Site Details** — total land cost, FAR ratio, gross/net acres, buildable SF, additional FAR, land SF, $/acre, $/land SF, $/buildable SF, efficiency ratio, max avg unit size, common area, density/net acre, buildable units, buildable NRSF, unit mix (market-rate / low-income / total), existing structure.
- **Initial Capitalization** — construction debt (% of cap, model-driven), equity, sponsor fees, total capitalization; sizing basis (LTC); development cash-needs build-up; cash reserves (interest reserve, operating reserve, lease-up income offset). Include a funding-order toggle (equity-first / pari-passu).
- **Construction Loan Details** — loan amount, start period, LTC, term (mos), interest rate, I/O period, interest method, interest calc type (30/360 or Actual/360), interest reserve, points & other loan costs.
- **Permanent Loan (Construction Loan Replacement)** — loan amount, refinance month, LTV at refi, term (mos), interest rate, I/O period, interest method, interest calc type, property valuation at refi, construction loan payoff, LTV %, total loan uses, points, resulting cash distribution at refi.
- **Fees to Sponsor** — development fee build-up (direct costs, debt service, operating deficit, lease-up income), third-party split, loan recourse/indemnity fee.
- **Loan Sizing Tests** — min debt yield, min DSCR, max LTC factor, max LTV factor, Year-1 NOI, refinance month, NOI at refi; construction-loan tests (debt yield & DSCR at target LTC) and permanent-loan tests (debt yield & DSCR at target LTV), each with a pass/fail check.
- **Unit Matrix** — residential matrix (market-rate & low-income) by floorplan: # units, % of total, avg SF, total SF, market rent, net-effective rent, annual rent, rent/SF; summary-by-unit-type; retail matrix.
- **Stabilized Operating Proforma** — revenue (total rent, low-income rent, gain/loss to lease, vacancy, model unit, bad debt, concessions, retail, parking, utility, other, total) and operating expenses (taxes, insurance, utilities, payroll, mgmt fee, R&M, turnover, contract services, reserves, admin & legal, advertising); NOI; below-the-line (capital reserves, tax prep, asset mgmt fees, debt service); operating CF. Per-unit/month, $/month, per-unit/year, $/year columns, each with its own growth rate.
- **Parking & Storage Revenue Matrix** and **Staffing & Payroll Expense Matrix**.
- **Development Cost Assumptions** — line-item budget grouped Land / Soft Costs (due diligence, entitlements, consultants, other soft costs, financing) / Hard Costs (GC contract, off-sites, utilities, contingency, builder's risk, GC fee), plus the 10+ user-definable lines. Each line: untrended amount, modeled amount, per unit, per SF, Phase, Timing rule, day Modifier, Cost Curve. Summary by grouping and by type.

## Phased build plan (each phase fully functional & reconciled before the next)
1. **Workbook architecture & naming conventions** — tab layout, Tables, named-range scheme, formatting/number-format standards, color conventions (blue input / black formula / green link), input vs. calc vs. output separation.
2. **Timeline engine** — milestone table → monthly period index, phase flags per month, milestone-relative positioning.
3. **Development budget & cost-curve engine** — line items (incl. user-definable) spread monthly by curve (`Straight_Line`, `Bell_Curve`, `Development_Fee`) using Phase/Timing/Modifier; curves must sum exactly to line totals.
4. **Construction loan engine** — monthly draws against cost-curve spend per funding order; balance rolls forward monthly; sizing to min(LTC, debt yield, DSCR).
5. **Interest reserve calculations** — sequential monthly interest on prior-month balance, capitalized; reserve derived, non-circular; lender-required-reserve override.
6. **Unit delivery & lease-up engine** — delivery and absorption schedules to stabilized occupancy.
7. **Stabilized operating model** — revenue & expense proforma with growth; lease-up expense ramp/elevation logic applied monthly.
8. **Property tax engine** — operating-period taxes; post-stabilization reassessment; disposition-period treatment.
9. **Permanent loan / refinance engine** — refi-or-sell switch; takeout sizing to min(LTV, debt yield, DSCR); construction payoff; cash-out and **No-Cash-Out toggle** (solve perm LTV so net cash at refi = 0); amortization.
10. **Long-term operating cash flow** — operations from stabilization through the full hold (up to 360 months), debt service on active loan.
11. **Disposition engine** — exit valuation, sale costs, debt payoff, gain calc, net proceeds.
12. **Levered & unlevered cash flows** — full monthly series, both perspectives.
13. **Return calculations** — levered/unlevered IRR & MOIC (gross & net), development spread/margin, ROC (current & stabilized), TDC ($, $/unit, $/SF), static (no-growth) reference case.
14. **Dashboard** — single summary view that updates automatically from inputs (key returns, sources & uses, timeline, sizing tests, sensitivities to the hold/refi switch).
15. **Diagnostics & validation** — full integrity layer with plain-English checks and a single "errors found" count.

## Disposition specifics
- **Exit value = forward-looking 12-month NOI ÷ exit cap rate** (the 12 months immediately following the disposition date). Ensure the operating model computes 12 months of forward NOI even at a 360-month hold.
- **Reassess-on-sale toggle:** when on, the buyer's forward taxes are based on sale value. Solve this **in closed form** to avoid circularity: `Exit Value = (Forward NOI before property tax) ÷ (Exit Cap Rate + Effective Tax Rate on Value)`. When off, use in-place taxes in forward NOI and `Exit Value = Forward NOI ÷ Exit Cap Rate`.
- Net proceeds: gross sale − broker/sale costs − sales tax − sponsor fee − property tax proration − payoff of the active loan (construction or permanent per the refi-or-sell switch).

## Outputs / returns
Levered & unlevered IRR and MOIC (gross and net of fees), development spread/margin (yield-on-cost vs. exit cap, in bps and %), return on cost (current & stabilized), total development cost ($, $/unit, $/SF), static IRR/MOIC reference, LTC, debt yield, current/exit basis per unit — all surfaced on the Dashboard and fully dynamic.

## Diagnostics / integrity layer
Plain-English checks for input and mechanics errors (e.g., loan not retired by proceeds, permanent term expiring before disposition, units delivered ≠ leased, budget not fully allocated, reserves not drawn to zero, cost curves not summing to totals). Single "total errors found" count with per-check status. Deliver with zero formula errors.

## Final validation checklist (verify before declaring complete)
- ✓ Zero Excel formula errors
- ✓ Zero diagnostic failures
- ✓ Construction draw schedule fully reconciles
- ✓ Development budget fully allocated
- ✓ Cost curves sum exactly to budget totals
- ✓ Interest reserve fully reconciles
- ✓ Construction loan balance reconciles monthly
- ✓ Permanent loan sizing reconciles
- ✓ Permanent loan retires construction loan
- ✓ No-Cash-Out toggle functions correctly
- ✓ Refi-or-sell switch behaves correctly on both sides of construction maturity
- ✓ Lease-up reaches stabilized occupancy
- ✓ Operating statements reconcile
- ✓ Lease-up expense ramp and elevation apply correctly
- ✓ Operating reserve reconciles to peak lease-up shortfall
- ✓ Property valuation reconciles to NOI and cap rate
- ✓ Reassess-on-sale closed form reconciles
- ✓ Disposition proceeds reconcile
- ✓ Levered cash flow reconciles
- ✓ Unlevered cash flow reconciles
- ✓ Levered IRR reconciles · ✓ Unlevered IRR reconciles
- ✓ Levered MOIC reconciles · ✓ Unlevered MOIC reconciles
- ✓ Hold-period toggle functions from 1 to 360 months
- ✓ Dashboard updates automatically

## Deliverable
A single, self-contained, institutional-quality `.xlsx` — fully monthly, hold toggleable 1–360 months, automatic refi-or-sell logic, No-Cash-Out toggle, reassess-on-sale toggle, model-calculated operating reserve, dynamic lease-up expense behavior, and user-definable budget lines — plus the Python generator, the shadow/validation scripts, and the audit log/README. After building, walk me through a base-case result and confirm every checklist item passes.
