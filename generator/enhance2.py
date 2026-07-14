"""
enhance2.py — usability / functionality / formatting enhancement passes applied
to Claude_Dev_Model.xlsx (post-remediation).  Phased; run one phase at a time.

Usage: python3 enhance2.py <src.xlsx> <dst.xlsx> phaseA[,phaseB,phaseC]
"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

NAVY = "FF1F3864"; INK = "FF000000"; BLUE_IN = "FF0000CC"; GREEN_LK = "FF006100"
AMBER_TXT = "FF9C6500"; AMBER_FILL = "FFFFF2CC"; RED_TXT = "FF9C0006"; RED_FILL = "FFFFC7CE"
GREEN_TXT = "FF006100"; GREEN_FILL = "FFC6EFCE"
FILL_IN = "FFFDF3D0"
F_SECT = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
F_LBL  = Font(name="Calibri", size=10, color=INK)
F_HINT = Font(name="Calibri", size=9, italic=True, color="FF595959")
F_IN   = Font(name="Calibri", size=10, color=BLUE_IN)
F_LINK = Font(name="Calibri", size=10, color=GREEN_LK)
F_RESV = Font(name="Calibri", size=9, color=GREEN_LK)
F_AMBER = Font(name="Calibri", size=10, bold=True, color=AMBER_TXT)
FILL_SECT = PatternFill("solid", fgColor=NAVY)
FILL_INPUT = PatternFill("solid", fgColor=FILL_IN)
FILL_AMBER = PatternFill("solid", fgColor=AMBER_FILL)
RIGHT = Alignment(horizontal="right"); LEFT = Alignment(horizontal="left", wrap_text=False); CENTER = Alignment(horizontal="center")
FMT_INT = '#,##0'


def S(ws, coord, value=None, font=None, fill=None, num=None, align=None):
    c = ws[coord]
    if value is not None: c.value = value
    if font: c.font = font
    if fill: c.fill = fill
    if num: c.number_format = num
    if align: c.alignment = align
    return c


def repoint(wb, name, ref):
    if name in wb.defined_names:
        del wb.defined_names[name]
    wb.defined_names.add(DefinedName(name, attr_text=ref))


# ---- toggle definitions: name -> (row, true_label, false_label) 1=true --------
TOGGLES = {
    'LeaseUpIncomeOffset': (31, 'Offset vs equity', 'Do not offset'),
    'TaxMethodValue':      (50, 'Value Approach', 'Income Approach'),
    'ReassessOnSale':      (52, 'Yes — reassess at sale', 'No — in-place taxes'),
    'FundingOrderPari':    (56, 'Pari-Passu', 'Equity First'),
    'ConActual360':        (64, 'Actual/360', '30/360'),
    'PermActual360':       (77, 'Actual/360', '30/360'),
    'NoCashOut':           (81, 'No Cash-Out (solve LTV to zero)', 'Cash-Out Permitted'),
    'TaxCapitalizeLeaseUp':(135, 'Capitalize in budget', 'Expense from delivery'),
}
TEXT_LISTS = {
    'State': (8, ["AZ", "TX", "CA", "CO", "NV", "UT", "NM", "WA", "OR", "ID", "FL", "GA", "NC", "TN"]),
    'ProductType': (10, ["Garden", "Mid-Rise", "High-Rise", "Wrap", "Podium", "Townhome"]),
    'Location': (11, ["Urban", "Suburban", "Suburban Infill", "Rural"]),
    'InvestmentType': (5, ["Ground-Up Development", "Value-Add", "Core", "Core-Plus"]),
}

# ---- numeric validation: name -> (row, kind, lo, hi, whole, unit) -------------
R, G, PCT, MON, CNT = 'rate', 'growth', 'pct', 'money', 'count'
VALID = {
    'HoldPeriodMonths': (15, 'hold', 1, 360, True, 'months'),
    'EntryCapRate': (16, R, 0, 0.25, False, 'rate %'),
    'CapAdjLocation': (17, 'adj', -0.05, 0.05, False, '± rate'),
    'CapAdjSize': (18, 'adj', -0.05, 0.05, False, '± rate'),
    'CapAdjAging': (19, 'adj', -0.05, 0.05, False, '± rate'),
    'ExitCapRetail': (21, R, 0, 0.25, False, 'rate %'),
    'GrowthRent': (22, G, -0.10, 0.20, False, 'annual %'),
    'GrowthLowIncRent': (23, G, -0.10, 0.20, False, 'annual %'),
    'GrowthExpense': (24, G, -0.10, 0.20, False, 'annual %'),
    'GrowthUtility': (25, G, -0.10, 0.20, False, 'annual %'),
    'GrowthTax': (26, G, -0.10, 0.20, False, 'annual %'),
    'GrowthOther': (27, G, -0.10, 0.20, False, 'annual %'),
    'UnitsDeliveredPerMo': (32, CNT, 1, 1000, True, 'units/mo'),
    'UnitsLeasedPerMo': (33, CNT, 1, 1000, True, 'units/mo'),
    'LeaseUpRampPct': (34, PCT, 0, 1, False, '% of stab'),
    'LeaseUpElevationPct': (35, PCT, 0, 1, False, 'uplift %'),
    'StabilizedOccupancy': (36, 'occ', 0.50, 1.00, False, 'occupancy %'),
    'LandCost': (39, MON, 0, 1e12, False, '$'),
    'GrossAcres': (40, 'acre', 0, 10000, False, 'acres'),
    'FARRatio': (42, 'far', 0, 20, False, 'ratio'),
    'EfficiencyRatio': (43, 'occ', 0.50, 1.00, False, 'NRSF/gross %'),
    'ExistingStructureCost': (44, MON, 0, 1e12, False, '$'),
    'TaxLevyPct': (47, R, 0, 0.25, False, 'rate %'),
    'TaxAssessmentPct': (48, PCT, 0, 1, False, 'assessment %'),
    'TaxValueAdjFactor': (49, PCT, 0, 1, False, 'factor'),
    'TaxYear1Override': (51, MON, 0, 1e9, False, '$ (0=auto)'),
    'DispoTaxRatePct': (53, R, 0, 0.25, False, 'rate % (0=auto)'),
    'TargetLTC': (57, 'ltv', 0, 0.90, False, 'LTC %'),
    'OpReserveCoverage': (58, 'mult', 1.0, 3.0, False, 'x multiple'),
    'OpReserveMinMonths': (59, CNT, 0, 24, True, 'months'),
    'ConLoanTermMonths': (62, 'conterm', 12, 60, True, 'months'),
    'ConRate': (63, R, 0, 0.25, False, 'rate %'),
    'ConPointsPct': (65, R, 0, 0.10, False, 'points %'),
    'ConOtherCosts': (66, MON, 0, 1e9, False, '$'),
    'ConMinDebtYield': (67, R, 0, 0.25, False, 'rate %'),
    'ConMinDSCR': (68, 'dscr', 1.0, 2.0, False, 'x'),
    'LenderReqReserveOverride': (69, MON, 0, 1e9, False, '$ (0=auto)'),
    'PermMaxLTV': (72, 'ltv', 0, 0.90, False, 'LTV %'),
    'PermRate': (73, R, 0, 0.25, False, 'rate %'),
    'PermTermMonths': (74, 'permterm', 60, 480, True, 'months'),
    'PermIOMonths': (75, CNT, 0, 120, True, 'months'),
    'PermAmortMonths': (76, 'permterm', 60, 480, True, 'months'),
    'PermPointsPct': (78, R, 0, 0.10, False, 'points %'),
    'PermMinDebtYield': (79, R, 0, 0.25, False, 'rate %'),
    'PermMinDSCR': (80, 'dscr', 1.0, 2.0, False, 'x'),
    'RefiMonthOverride': (82, CNT, 0, 360, True, 'month (0=auto)'),
    'DevFeePctDirect': (85, R, 0, 0.10, False, '% of direct'),
    'DevFeeThirdPartySplit': (86, PCT, 0, 1, False, 'split %'),
    'SponsorDispoFeePct': (87, R, 0, 0.10, False, '% of sale'),
    'RecourseFeePct': (88, R, 0, 0.10, False, '% of loan'),
    'BrokerCostPct': (91, R, 0, 0.10, False, '% of sale'),
    'SalesTaxPct': (92, R, 0, 0.10, False, '% of sale'),
    'TaxProrationMonths': (93, CNT, 0, 12, True, 'months'),
    'BadDebtPct': (96, R, 0, 0.10, False, '% of GPR'),
    'ConcessionsPct': (97, R, 0, 0.10, False, '% of GPR'),
    'LossToLeasePct': (98, R, 0, 0.10, False, '% of GPR'),
    'ModelUnits': (99, CNT, 0, 1000, True, 'units'),
    'ParkingPerUnitMo': (100, MON, 0, 100000, False, '$/unit/mo'),
    'RetailIncomeMo': (101, MON, 0, 1e7, False, '$/mo'),
    'UtilReimbPerUnitMo': (102, MON, 0, 100000, False, '$/unit/mo'),
    'OtherIncPerUnitMo': (103, MON, 0, 100000, False, '$/unit/mo'),
    'ExpInsurancePerUnit': (106, MON, 0, 100000, False, '$/unit/yr'),
    'ExpUtilitiesPerUnit': (107, MON, 0, 100000, False, '$/unit/yr'),
    'ExpPayrollPerUnit': (108, MON, 0, 100000, False, '$/unit/yr'),
    'ExpRMPerUnit': (109, MON, 0, 100000, False, '$/unit/yr'),
    'ExpTurnoverPerUnit': (110, MON, 0, 100000, False, '$/unit/yr'),
    'ExpContractPerUnit': (111, MON, 0, 100000, False, '$/unit/yr'),
    'ExpAdminPerUnit': (112, MON, 0, 100000, False, '$/unit/yr'),
    'ExpAdvertisingPerUnit': (113, MON, 0, 100000, False, '$/unit/yr'),
    'MgmtFeePct': (114, R, 0, 0.10, False, '% of EGI'),
    'CapReservePerUnit': (115, MON, 0, 100000, False, '$/unit/yr'),
    'TaxPrepPerUnit': (116, MON, 0, 100000, False, '$/unit/yr'),
    'AssetMgmtFeePct': (117, R, 0, 0.10, False, '% of EGI'),
}
# milestone durations + lag: all positive whole
for _nm, _row in [('DeliveryLagMonths', 120), ('Dur_OpenEscrow', 121), ('Dur_DueDiligence', 122),
                  ('Dur_Entitlement', 123), ('Dur_PreConstruction', 124), ('Dur_InitialClosing', 125),
                  ('Dur_Construction', 126), ('Dur_FirstUnitDelivery', 127), ('Dur_Stabilization', 128),
                  ('Dur_Recap', 129), ('Dur_Disposition', 130)]:
    VALID[_nm] = (_row, 'dur', 1, 120, True, 'months')


def _fmt_range(kind, lo, hi):
    if kind in ('rate', 'growth', 'pct', 'adj', 'occ', 'ltv'):
        return f"{lo:.0%} to {hi:.0%}"
    if kind in ('dscr', 'mult'):
        return f"{lo:.2f}x to {hi:.2f}x"
    if kind == 'money':
        return f"$0+"
    return f"{lo:,.0f} to {hi:,.0f}"


def phase_a(wb):
    ip = wb['Inputs']
    ip.column_dimensions['E'].width = 22   # hint / resolved column
    ip.column_dimensions['D'].width = 30
    ip.column_dimensions['F'].width = 10

    # ---- A1: toggles -> dropdown text in C, resolved 0/1 in F (name -> F) -----
    for name, (row, tl, fl) in TOGGLES.items():
        cur = ip.cell(row, 3).value
        # existing value is 0/1 (or formula-resolved); pick label from int
        try:
            cur_is_true = int(float(cur)) == 1
        except (TypeError, ValueError):
            cur_is_true = True
        S(ip, f'C{row}', tl if cur_is_true else fl, F_IN, FILL_INPUT, align=LEFT)
        dv = DataValidation(type='list', formula1=f'"{tl},{fl}"', allow_blank=False,
                            showErrorMessage=True)
        dv.error = f'Choose "{tl}" or "{fl}".'; dv.errorTitle = 'Pick from the list'
        dv.prompt = f'{tl} = 1 · {fl} = 0'; dv.promptTitle = 'Toggle'
        ip.add_data_validation(dv); dv.add(ip[f'C{row}'])
        S(ip, f'F{row}', f'=IF(C{row}="{tl}",1,0)', F_RESV, num=FMT_INT, align=CENTER)
        repoint(wb, name, f"'Inputs'!$F${row}")

    # ---- A1: text dropdowns (no helper) --------------------------------------
    for name, (row, opts) in TEXT_LISTS.items():
        dv = DataValidation(type='list', formula1='"' + ",".join(opts) + '"',
                            allow_blank=True, showErrorMessage=False)
        dv.prompt = 'Pick from list or type your own'; dv.promptTitle = name
        ip.add_data_validation(dv); dv.add(ip[f'C{row}'])

    # ---- A2 + A3: numeric data validation + input-hint column D --------------
    for name, (row, kind, lo, hi, whole, unit) in VALID.items():
        dvtype = 'whole' if whole else 'decimal'
        dv = DataValidation(type=dvtype, operator='between', formula1=repr(lo), formula2=repr(hi),
                            allow_blank=False, showErrorMessage=True, showInputMessage=True)
        rng = _fmt_range(kind, lo, hi)
        dv.errorTitle = 'Out of range'
        dv.error = f'{ip.cell(row,2).value}: enter {unit} in {rng}.'
        dv.promptTitle = ip.cell(row, 2).value
        dv.prompt = f'Unit: {unit}.  Expected: {rng}.'
        ip.add_data_validation(dv); dv.add(ip[f'C{row}'])
        # A3 hint in D: keep existing "what it drives" note, append unit + range
        note = ip.cell(row, 4).value
        drives = str(note) if note else ''
        hint = f'{unit}, {rng}' + (f' · {drives}' if drives else '')
        S(ip, f'D{row}', hint, F_HINT, align=LEFT)

    # ---- A4: Input Sanity block (amber soft warnings) ------------------------
    start = ip.max_row + 2
    S(ip, f'B{start}', 'Input Sanity (soft warnings — review, not errors)', F_SECT, FILL_SECT, align=LEFT)
    for col in 'CDEF': ip[f'{col}{start}'].fill = FILL_SECT
    r = start + 1
    checks = [
        ('Construction term ≤ 48 months', '=ConLoanTermMonths', FMT_INT,
         '=IF(ConLoanTermMonths<=48,"OK","CHECK")'),
        ('Stabilized tax / unit in $500–$4,000', '=IF(TotalUnits=0,0,StabTaxAnnual/TotalUnits)', '$#,##0',
         '=IF(AND(StabTaxAnnual/TotalUnits>=500,StabTaxAnnual/TotalUnits<=4000),"OK","CHECK")'),
        ('Absorption pace ≤ delivery pace', '=UnitsLeasedPerMo', FMT_INT,
         '=IF(UnitsLeasedPerMo<=UnitsDeliveredPerMo,"OK","CHECK")'),
        ('Untrended YOC ≥ exit cap (positive spread)', '=YieldOnCost-ExitCapResidential', '0.00%',
         '=IF(YieldOnCost>=ExitCapResidential,"OK","CHECK")'),
        ('Stabilized occupancy 85%–98%', '=StabilizedOccupancy', '0.0%',
         '=IF(AND(StabilizedOccupancy>=0.85,StabilizedOccupancy<=0.98),"OK","CHECK")'),
        ('Rent growth ≤ 5%', '=GrowthRent', '0.00%',
         '=IF(GrowthRent<=0.05,"OK","CHECK")'),
    ]
    sanity_first = r
    for label, cval, fmt, dform in checks:
        S(ip, f'B{r}', label, F_LBL, align=LEFT)
        S(ip, f'C{r}', cval, F_LINK, num=fmt, align=RIGHT)
        S(ip, f'E{r}', dform, F_AMBER, align=CENTER)
        r += 1
    sanity_last = r - 1
    repoint(wb, 'InputSanityWarnings', f"'Inputs'!$E${sanity_first}")  # anchor for CF/summary
    # count of sanity checks flagged (referenced by Dashboard later)
    S(ip, f'B{r+0}', 'Sanity checks flagged', F_LBL, align=LEFT)
    S(ip, f'C{r+0}', f'=COUNTIF(E{sanity_first}:E{sanity_last},"CHECK")', F_LINK, num=FMT_INT, align=RIGHT)
    repoint(wb, 'SanityFlags', f"'Inputs'!$C${r}")

    # amber conditional formatting on the sanity status cells
    from openpyxl.formatting.rule import CellIsRule
    amber = CellIsRule(operator='equal', formula=['"CHECK"'],
                       font=Font(bold=True, color=AMBER_TXT), fill=FILL_AMBER)
    ip.conditional_formatting.add(f'E{sanity_first}:E{sanity_last}', amber)
    return {'sanity': (sanity_first, sanity_last), 'toggles': len(TOGGLES), 'validated': len(VALID)}


PHASES = {'phaseA': phase_a}


def main():
    src, dst, which = sys.argv[1], sys.argv[2], sys.argv[3]
    wb = openpyxl.load_workbook(src)
    for n in which.split(','):
        info = PHASES[n](wb)
        print(f"applied {n}: {info}")
    wb.save(dst)
    print(f"saved {dst}")


# ======================================================================  PHASE B
F_TITLE = Font(name="Calibri", size=16, bold=True, color=NAVY)
F_HDR = Font(name="Calibri", size=9, bold=True, color="FFFFFFFF")
F_TOT = Font(name="Calibri", size=10, bold=True, color=INK)
F_FORM = Font(name="Calibri", size=10, color=INK)
FILL_TOTAL = PatternFill("solid", fgColor="FFE2EFDA")
FILL_HDR = PatternFill("solid", fgColor="FF2E5496")
FMT_MONEY = '$#,##0;($#,##0);"–"'
TOPB = Border(top=Side(style="thin", color="FF808080"))

# scenario-driven inputs: name -> (inputs_row, base, upside, downside, fmt)
SCEN = [
    ('GrowthRent', 22, 0.03, 0.04, 0.015, '0.00%'),
    ('GrowthExpense', 24, 0.03, 0.025, 0.04, '0.00%'),
    ('EntryCapRate', 16, 0.0525, 0.0500, 0.0575, '0.00%'),
    ('ExitCapRetail', 21, 0.065, 0.0625, 0.070, '0.00%'),
    ('StabilizedOccupancy', 36, 0.94, 0.95, 0.90, '0.0%'),
    ('UnitsLeasedPerMo', 33, 18, 22, 12, '#,##0'),
    ('ConRate', 63, 0.057, 0.052, 0.065, '0.00%'),
    ('PermRate', 73, 0.0625, 0.0575, 0.070, '0.00%'),
]


def _group_years(ws):
    ws.sheet_properties.outlinePr.summaryRight = True
    for y in range(1, 32):            # 31 years -> months 1..372 (cols E..NL)
        first_m = (y - 1) * 12 + 1
        first_col = 5 + first_m - 1   # month 1 -> col E(5)
        g0 = get_column_letter(first_col + 1)
        g1 = get_column_letter(min(first_col + 11, 376))
        if first_col + 1 <= 376:
            ws.column_dimensions.group(g0, g1, outline_level=1, hidden=False)


def phase_b(wb):
    # ---- B2: year column grouping on every 360-col tab -----------------------
    for name in ['Timeline', 'DevSpend', 'ConLoan', 'LeaseUp', 'Operating',
                 'PermLoan', 'CashFlow', 'Returns']:
        _group_years(wb[name])

    # ---- B5: binding sizing constraint (con + perm) --------------------------
    cl = wb['ConLoan']
    S(cl, 'B21', 'Binding Constraint', F_LBL, align=LEFT)
    S(cl, 'C21', '=IF(ConCostCommit=MaxByLTC,"LTC",IF(ConCostCommit=MaxByDebtYield,"Debt Yield","DSCR"))',
      F_LINK, align=RIGHT)
    repoint(wb, 'ConBinding', "'ConLoan'!$C$21")
    pl = wb['PermLoan']
    S(pl, 'B21', 'Binding Constraint', F_LBL, align=LEFT)
    S(pl, 'C21', '=IF(RefiFlag=0,"n/a (sell)",IF(PermSized=MaxLTVperm,"LTV",IF(PermSized=MaxDYperm,"Debt Yield","DSCR")))',
      F_LINK, align=RIGHT)
    repoint(wb, 'PermBinding', "'PermLoan'!$C$21")

    # ---- B4: scenario switcher (Scenarios tab + repoint inputs to INDEX) ------
    if 'Scenarios' in wb.sheetnames:
        del wb['Scenarios']
    sc = wb.create_sheet('Scenarios')
    wb.move_sheet('Scenarios', -(len(wb.sheetnames) - 2))  # after Dashboard
    for col, w in {'A': 3.7, 'B': 30, 'C': 14, 'D': 14, 'E': 14, 'F': 14}.items():
        sc.column_dimensions[col].width = w
    sc.sheet_view.showGridLines = False
    sc.sheet_properties.tabColor = '7030A0'
    S(sc, 'B1', 'Scenario Switcher', F_TITLE)
    S(sc, 'B3', 'Active Scenario', F_LBL, align=LEFT)
    S(sc, 'C3', 'Base', F_IN, FILL_INPUT, align=CENTER)
    dv = DataValidation(type='list', formula1='"Base,Upside,Downside"', allow_blank=False,
                        showErrorMessage=True)
    dv.promptTitle = 'Scenario'; dv.prompt = 'Base / Upside / Downside drives all key assumptions'
    sc.add_data_validation(dv); dv.add(sc['C3'])
    S(sc, 'B4', 'Scenario Number', F_LBL, align=LEFT)
    S(sc, 'C4', '=IF(C3="Base",1,IF(C3="Upside",2,3))', F_LINK, num=FMT_INT, align=CENTER)
    repoint(wb, 'ScenarioSelect', "'Scenarios'!$C$3")
    repoint(wb, 'ScenarioNum', "'Scenarios'!$C$4")
    # header
    hr = 6
    for j, h in enumerate(['Driver', 'Base', 'Upside', 'Downside', 'Active'], start=2):
        S(sc, f'{get_column_letter(j)}{hr}', h, F_HDR, FILL_HDR, align=(LEFT if j == 2 else CENTER))
    r = hr + 1
    ip = wb['Inputs']
    for name, irow, base, up, down, fmt in SCEN:
        S(sc, f'B{r}', ip.cell(irow, 2).value, F_LBL, align=LEFT)
        S(sc, f'C{r}', base, F_IN, FILL_INPUT, num=fmt, align=CENTER)
        S(sc, f'D{r}', up, F_IN, FILL_INPUT, num=fmt, align=CENTER)
        S(sc, f'E{r}', down, F_IN, FILL_INPUT, num=fmt, align=CENTER)
        S(sc, f'F{r}', f'=INDEX(C{r}:E{r},ScenarioNum)', F_LINK, num=fmt, align=CENTER)
        # repoint the Inputs cell to read the active scenario value
        ip.cell(irow, 3).value = f"=Scenarios!$F${r}"
        ip.cell(irow, 3).font = F_LINK
        ip.cell(irow, 3).fill = PatternFill(fill_type=None)
        r += 1

    # ---- B1 + B3: Annual Summary tab (Sources & Uses + annual rollup) --------
    if 'Annual' in wb.sheetnames:
        del wb['Annual']
    an = wb.create_sheet('Annual')
    wb.move_sheet('Annual', -(len(wb.sheetnames) - 3))  # after Scenarios
    an.sheet_view.showGridLines = False
    an.sheet_properties.tabColor = '375623'
    an.column_dimensions['A'].width = 3.7
    an.column_dimensions['B'].width = 34
    S(an, 'B1', 'Annual Summary', F_TITLE)
    S(an, 'B2', '="Scenario: "&ScenarioSelect&"   •   Hold: "&HoldPeriodMonths&" months"', F_HINT)

    # Sources & Uses block (B3)
    S(an, 'B4', 'Sources & Uses', F_SECT, FILL_SECT, align=LEFT)
    for col in 'CD': an[f'{col}4'].fill = FILL_SECT
    su = [
        ('Uses', None), ('  Land', '=BudgetLand'), ('  Hard Costs', '=BudgetHard'),
        ('  Soft Costs', '=BudgetSoft'), ('  Development Fee', '=DevFee'),
        ('  Loan Costs (points, recourse, other)', '=ConPoints+RecourseFee+ConOtherCosts'),
        ('  Interest Reserve', '=InterestReserveFunded'), ('  Operating Reserve', '=OpReserve'),
        ('  Total Uses (TDC)', '=TDC_Total'),
        ('Sources', None), ('  Construction Loan', '=ConLoanTotal'), ('  Equity', '=EquityTotal'),
        ('  Refi Cash-Out (memo)', '=RefiCashOut'), ('  Total Sources', '=ConLoanTotal+EquityTotal'),
        ('  Sources − Uses (check = 0)', '=ConLoanTotal+EquityTotal-TDC_Total'),
    ]
    r = 5
    for label, form in su:
        bold = label in ('Uses', 'Sources') or 'Total' in label or 'check' in label
        S(an, f'B{r}', label, F_TOT if bold else F_LBL, align=LEFT)
        if form:
            S(an, f'C{r}', form, F_TOT if bold else F_FORM, num=FMT_MONEY, align=RIGHT)
        if 'Total Uses' in label or 'Total Sources' in label or 'check' in label:
            for col in 'BC': an[f'{col}{r}'].border = TOPB
        r += 1
    repoint(wb, 'SU_Check', f"'Annual'!$C${r-1}")

    # Annual rollup grid
    gr = r + 1
    S(an, f'B{gr}', 'Annual Cash Flow (Year 1 … 31)', F_SECT, FILL_SECT, align=LEFT)
    hdr = gr + 1
    n_years = 31
    S(an, f'B{hdr}', 'Year →', F_HDR, FILL_HDR, align=LEFT)
    for y in range(1, n_years + 1):
        col = get_column_letter(2 + y)
        S(an, f'{col}{hdr}', y, F_HDR, FILL_HDR, num=FMT_INT, align=RIGHT)
        an.column_dimensions[col].width = 12
    # metric rows: (label, sheet, row, mode)  mode: 'flow' sum12 / 'eoy' year-end
    METRICS = [
        ('Development Spend', 'DevSpend', 34, 'flow'),
        ('Construction Debt Draw', 'ConLoan', 29, 'flow'),
        ('Equity Draw', 'ConLoan', 27, 'flow'),
        ('Effective Gross Income', 'Operating', 23, 'flow'),
        ('Net Operating Income', 'Operating', 38, 'flow'),
        ('Construction Interest (capitalized)', 'ConLoan', 31, 'flow'),
        ('Permanent Debt Service', 'PermLoan', 29, 'flow'),
        ('Unlevered Cash Flow', 'CashFlow', 7, 'flow'),
        ('Levered Cash Flow', 'CashFlow', 6, 'flow'),
        ('Construction Loan Balance (EOY)', 'ConLoan', 32, 'eoy'),
        ('Permanent Loan Balance (EOY)', 'PermLoan', 30, 'eoy'),
        ('Physical Occupancy (EOY)', 'LeaseUp', 10, 'eoy'),
    ]
    rr = hdr + 1
    metric_rows = {}
    for label, sheet, srow, mode in METRICS:
        S(an, f'B{rr}', label, F_LBL, align=LEFT)
        occ = (label.startswith('Physical'))
        for y in range(1, n_years + 1):
            col = get_column_letter(2 + y)
            rng = f"{sheet}!$E${srow}:$NL${srow}"
            m0 = f"(({col}${hdr}-1)*12+1)"
            if mode == 'flow':
                f = f"=SUM(INDEX({rng},{m0}):INDEX({rng},{m0}+11))"
            else:
                f = f"=INDEX({rng},{m0}+11)"
            S(an, f'{col}{rr}', f, F_FORM, num=('0.0%' if occ else FMT_MONEY), align=RIGHT)
        metric_rows[label] = rr
        rr += 1

    # reconciliation: annual NOI total ties to monthly engine total
    rr += 1
    S(an, f'B{rr}', 'Check: ΣAnnual NOI − monthly total', F_TOT, align=LEFT)
    noi_row = metric_rows['Net Operating Income']
    last_col = get_column_letter(2 + n_years)
    S(an, f'C{rr}', f"=SUM(C{noi_row}:{last_col}{noi_row})-Operating!$D$38", F_TOT, num=FMT_MONEY, align=RIGHT)
    repoint(wb, 'AnnualNOICheck', f"'Annual'!$C${rr}")

    an.freeze_panes = 'C' + str(hdr + 1)
    return {'scenarios': len(SCEN), 'annual_metrics': len(METRICS), 'su_rows': len(su)}


PHASES['phaseB'] = phase_b


if __name__ == "__main__":
    main()
