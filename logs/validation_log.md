# Validation Log & Base-Case Walkthrough

Every figure below is produced by the independent Python shadow model and
reconciles to the LibreOffice-equivalent recalculated workbook (via the
`formulas` engine) to <1e-6 relative. See validate.py / scenario_test.py.

## Base-case deal (Meridian Apartments, Austin TX — 200 units)

| Metric | Value |
|---|---|
| Total units | 200 |
| Net rentable SF | 158,740 |
| Hold period | 84 months |
| Refi or sell | Refinance then hold |
| Construction maturity | month 42 |
| Refi month | month 42 |
| Direct costs | $57,842,090 |
| Development fee | $2,313,684 |
| Interest reserve | $7,511,212 |
| Operating reserve | $446,890 |
| Total development cost | $68,858,303 |
|   TDC / unit | $344,292 |
|   TDC / NRSF | $434 |
| Construction loan | $47,065,443 |
| Total equity | $21,792,860 |
| Stabilized NOI (fwd 12mo) | $4,148,156 |
| Year-1 NOI | $1,838,529 |
| Value at refi | $89,325,145 |
| Permanent loan | $49,935,220 |
| Refi cash distribution | $2,495,262 |
| Exit value (reassess-on-sale) | $78,872,082 |
| Blended exit cap | 5.05% |
| Net sale proceeds | $27,739,471 |
| Yield on cost | 6.02% |
| Development spread | 102 bps |
| Return on cost (stab) | 6.02% |
| Levered IRR | 11.27% |
| Levered MOIC | 1.77x |
| Unlevered IRR | 9.25% |
| Unlevered MOIC | 1.61x |
| Static unlevered IRR | 5.50% |
| Static unlevered MOIC | 1.33x |

## Phase validation record

| Phase | Module | Status |
|---|---|---|
| 01 | Architecture, inputs, unit matrix, budget | PASS |
| 02 | Timeline engine | PASS |
| 03 | Cost-curve engine (curves sum to line totals) | PASS |
| 04 | Construction loan sizing & draws | PASS |
| 05 | Interest reserve (sequential, non-circular) | PASS |
| 06 | Unit delivery & lease-up | PASS |
| 07 | Stabilized operating model + ramp/elevation | PASS |
| 08 | Property tax engine | PASS |
| 09 | Permanent loan / refinance + No-Cash-Out | PASS |
| 10 | Long-term operating cash flow | PASS |
| 11 | Disposition (reassess closed form) | PASS |
| 12 | Levered & unlevered cash flows | PASS |
| 13 | Returns + static reference | PASS |
| 14 | Dashboard | PASS |
| 15 | Diagnostics (22 checks) | PASS |

## Scenario matrix (workbook recalculated, reconciled to shadow)

| Scenario | Result |
|---|---|
| Base (hold 84, refi) | PASS — 0 err cells |
| Sell during construction (hold 30) | PASS — no refi, con retired at sale |
| Sell at maturity (hold 42) | PASS — no refi |
| Refi just past maturity (hold 43) | PASS — refi engaged |
| Long hold (hold 180) | PASS |
| Max hold (hold 360) | PASS |
| No-Cash-Out on | PASS — net cash at refi = 0 |
| Reassess-on-sale off | PASS |
| Pari-passu funding | PASS |
| Lender reserve override | PASS |

## Final checklist

- [x] Zero Excel formula errors
- [x] Zero diagnostic failures
- [x] Construction draw schedule reconciles
- [x] Development budget fully allocated
- [x] Cost curves sum exactly to totals
- [x] Interest reserve reconciles
- [x] Construction loan balance reconciles monthly
- [x] Permanent loan sizing reconciles
- [x] Permanent loan retires construction loan
- [x] No-Cash-Out toggle functions
- [x] Refi-or-sell switch correct both sides of maturity
- [x] Lease-up reaches stabilized occupancy
- [x] Operating statements reconcile
- [x] Lease-up ramp & elevation apply
- [x] Operating reserve = peak shortfall
- [x] Property valuation reconciles to NOI & cap
- [x] Reassess-on-sale closed form reconciles
- [x] Disposition proceeds reconcile
- [x] Levered cash flow reconciles
- [x] Unlevered cash flow reconciles
- [x] Levered & Unlevered IRR reconcile
- [x] Levered & Unlevered MOIC reconcile
- [x] Hold-period toggle functions 1..360
- [x] Dashboard updates automatically
