# Validation Log & Base-Case Walkthrough

Every figure is produced by the independent Python shadow model and reconciles
to the recalculated workbook (via the `formulas` engine) to <1e-6 relative.

## Base-case deal (Meridian Apartments, Austin TX — 200 units)

| Metric | Value |
|---|---|
| Total units | 200 |
| Net rentable SF | 158,740 |
| Hold period | 84 months |
| Refi or sell | Refinance then hold |
| Growth start month | 2 |
| Direct costs (budget lines) | $57,142,090 |
| Construction-period tax (dynamic) | $781,578 |
| Development fee | $2,285,684 |
| Interest reserve | $7,574,736 |
| Operating reserve (recycled) | $487,441 |
| Total development cost | $69,016,722 |
|   TDC / unit | $345,084 |
|   TDC / NRSF | $435 |
| Stabilized assessed value | $85,613,419 |
| Year-1 stabilized tax (dynamic) | $791,924 |
| Construction loan | $47,190,152 |
| Total equity | $21,826,570 |
| Stabilized NOI (after tax) | $4,059,404 |
| Value at refi | $87,435,378 |
| Permanent loan | $48,878,788 |
| Refi cash distribution | $1,322,046 |
| Exit value (yr-7, reassess) | $92,069,131 |
| Blended exit cap | 5.05% |
| Net sale proceeds | $41,617,848 |
| Yield on cost | 5.88% |
| Development spread | 88 bps |
| Levered IRR | 17.12% |
| Levered MOIC | 2.40x |
| Unlevered IRR | 11.93% |
| Unlevered MOIC | 1.84x |
| Static (no-growth) unlev IRR | 7.95% |
| Static (no-growth) unlev MOIC | 1.51x |

## Return drivers (why an 89bps going-in spread supports a mid-teens IRR)

The development spread measures **going-in** stabilized yield-on-cost vs exit cap
only. On a ~5% cap, 89bps = ~18% value margin on cost, realized by stabilization.
The IRR additionally captures ~7 years of rent growth valued at the exit cap, a
cash-out refinance, interim operating cash flow and leverage. The **static
(no-growth) unlevered IRR of 7.9%** isolates the merchant-development
return with growth stripped out.

## Scenario matrix (workbook recalculated, reconciled to shadow)

All 10 scenarios PASS with 0 error cells: base; sell during construction (hold 30);
sell at maturity (42); refi just past maturity (43); long hold (180); max hold (360);
No-Cash-Out on; reassess-on-sale off; pari-passu funding; lender reserve override.
