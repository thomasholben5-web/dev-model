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


if __name__ == '__main__':
    main()
