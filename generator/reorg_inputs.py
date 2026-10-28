"""
reorg_inputs.py — rebuild the Inputs tab into a dense, multi-column, dashboard-style
layout (modeled on the reference underwriting workbook): a KPI summary strip across
the top, then grouped input sections laid out side-by-side in parallel column blocks.

Also:
  - relabels the exit-tax toggle as "Dispo Tax Driver" with Reassessed / Imputed
  - adds "Static Value (Untrended NOI / Entry Cap)" (named StaticValue), shown in the
    KPI strip and surfaced on the Dashboard

Every named input is repointed to its new cell; nothing downstream (which all keys off
named ranges) changes, so model outputs are identical.
"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import CellIsRule

NAVY = "FF1F3864"; BLUE = "FF0000CC"; GREEN = "FF006100"; INK = "FF000000"
FILL_IN = PatternFill("solid", fgColor="FFFDF3D0")
FILL_SECT = PatternFill("solid", fgColor=NAVY)
FILL_KPI = PatternFill("solid", fgColor="FFEAF0F8")
FILL_TITLE = PatternFill("solid", fgColor="FF1F3864")
AMBER_T = "FF9C6500"; AMBER_F = PatternFill("solid", fgColor="FFFFF2CC")
F_TITLE = Font(name="Calibri", size=15, bold=True, color="FFFFFFFF")
F_SECT = Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
F_LBL = Font(name="Calibri", size=10, color=INK)
F_IN = Font(name="Calibri", size=10, color=BLUE)
F_LINK = Font(name="Calibri", size=10, color=GREEN)
F_FORM = Font(name="Calibri", size=10, color=INK)
F_RESV = Font(name="Calibri", size=8, color=GREEN)
F_KPIL = Font(name="Calibri", size=9, color="FF595959")
F_KPIV = Font(name="Calibri", size=12, bold=True, color=NAVY)
F_HINT = Font(name="Calibri", size=8, italic=True, color="FF808080")
F_AMBER = Font(name="Calibri", size=9, bold=True, color=AMBER_T)
RIGHT = Alignment(horizontal="right"); LEFT = Alignment(horizontal="left"); CENTER = Alignment(horizontal="center")
FMT_INT = '#,##0'; FMT_PCT = '0.00%'; FMT_PCT1 = '0.0%'; FMT_MONEY = '$#,##0;($#,##0)'
FMT_MULT = '0.00"x"'; FMT_BPS = '#,##0" bps"'; FMT_NUM2 = '#,##0.00'

# ---- toggles: name -> (true_label, false_label)  (1 = true_label) -------------
TOGGLES = {
    'LeaseUpIncomeOffset': ('Offset vs equity', 'Do not offset'),
    'TaxMethodValue': ('Value Approach', 'Income Approach'),
    'ReassessOnSale': ('Reassessed', 'Imputed'),   # relabeled -> "Dispo Tax Driver"
    'FundingOrderPari': ('Pari-Passu', 'Equity First'),
    'ConActual360': ('Actual/360', '30/360'),
    'PermActual360': ('Actual/360', '30/360'),
    'NoCashOut': ('No Cash-Out (solve LTV to zero)', 'Cash-Out Permitted'),
    'TaxCapitalizeLeaseUp': ('Capitalize in budget', 'Expense from delivery'),
    'InterestMethodCurrentPay': ('Current Pay', 'Capitalized'),
}
TEXT_LISTS = {
    'InvestmentType': ["Ground-Up Development", "Value-Add", "Core", "Core-Plus"],
    'State': ["AZ", "TX", "CA", "CO", "NV", "UT", "NM", "WA", "OR", "ID", "FL", "GA", "NC", "TN"],
    'ProductType': ["Garden", "Mid-Rise", "High-Rise", "Wrap", "Podium", "Townhome"],
    'Location': ["Urban", "Suburban", "Suburban Infill", "Rural"],
}
FORMULA_INPUTS = {'ExitCapResidential', 'GrowthStartDate', 'NetAcres'}
FREE_TEXT = {'PropertyName', 'City', 'Market'}
DATE_INPUTS = {'InceptionDate'}

# ---- numeric validation: name -> (lo, hi, whole, unit) ------------------------
R = (0, 0.25, False, 'rate %'); G = (-0.10, 0.20, False, 'annual %')
VALID = {
    'HoldPeriodMonths': (1, 360, True, 'months'), 'EntryCapRate': R, 'CapAdjLocation': (-0.05, 0.05, False, '± rate'),
    'CapAdjSize': (-0.05, 0.05, False, '± rate'), 'CapAdjAging': (-0.05, 0.05, False, '± rate'), 'ExitCapRetail': R,
    'GrowthRent': G, 'GrowthLowIncRent': G, 'GrowthExpense': G, 'GrowthUtility': G, 'GrowthTax': G, 'GrowthOther': G,
    'UnitsDeliveredPerMo': (1, 1000, True, 'units/mo'), 'UnitsLeasedPerMo': (1, 1000, True, 'units/mo'),
    'LeaseUpRampPct': (0, 1, False, '% of stab'), 'LeaseUpElevationPct': (0, 1, False, 'uplift %'),
    'StabilizedOccupancy': (0.50, 1.0, False, 'occupancy %'), 'LandCost': (0, 1e12, False, '$'),
    'GrossAcres': (0, 10000, False, 'acres'), 'FARRatio': (0, 20, False, 'ratio'),
    'EfficiencyRatio': (0.5, 1.0, False, 'NRSF/gross %'), 'ExistingStructureCost': (0, 1e12, False, '$'),
    'TaxLevyPct': R, 'TaxAssessmentPct': (0, 1, False, 'assessment %'), 'TaxValueAdjFactor': (0, 1, False, 'factor'),
    'TaxYear1Override': (0, 1e9, False, '$ (0=auto)'), 'DispoTaxRatePct': R, 'TargetLTC': (0, 0.90, False, 'LTC %'),
    'OpReserveCoverage': (1.0, 3.0, False, 'x'), 'OpReserveMinMonths': (0, 24, True, 'months'),
    'ConLoanTermMonths': (12, 60, True, 'months'), 'ConRate': R, 'ConPointsPct': (0, 0.10, False, 'points %'),
    'ConOtherCosts': (0, 1e9, False, '$'), 'ConMinDebtYield': R, 'ConMinDSCR': (1.0, 2.0, False, 'x'),
    'LenderReqReserveOverride': (0, 1e9, False, '$ (0=auto)'), 'PermMaxLTV': (0, 0.90, False, 'LTV %'), 'PermRate': R,
    'PermTermMonths': (60, 480, True, 'months'), 'PermIOMonths': (0, 120, True, 'months'),
    'PermAmortMonths': (60, 480, True, 'months'), 'PermPointsPct': (0, 0.10, False, 'points %'),
    'PermMinDebtYield': R, 'PermMinDSCR': (1.0, 2.0, False, 'x'), 'RefiMonthOverride': (0, 360, True, 'mo (0=auto)'),
    'DevFeePctDirect': (0, 0.10, False, '% direct'), 'DevFeeThirdPartySplit': (0, 1, False, 'split %'),
    'SponsorDispoFeePct': (0, 0.10, False, '% sale'), 'RecourseFeePct': (0, 0.10, False, '% loan'),
    'BrokerCostPct': (0, 0.10, False, '% sale'), 'SalesTaxPct': (0, 0.10, False, '% sale'),
    'TaxProrationMonths': (0, 12, True, 'months'), 'BadDebtPct': (0, 0.10, False, '% GPR'),
    'ConcessionsPct': (0, 0.10, False, '% GPR'), 'LossToLeasePct': (0, 0.10, False, '% GPR'),
    'ModelUnits': (0, 1000, True, 'units'), 'ParkingPerUnitMo': (0, 1e5, False, '$/unit/mo'),
    'RetailIncomeMo': (0, 1e7, False, '$/mo'), 'UtilReimbPerUnitMo': (0, 1e5, False, '$/unit/mo'),
    'OtherIncPerUnitMo': (0, 1e5, False, '$/unit/mo'), 'ExpInsurancePerUnit': (0, 1e5, False, '$/unit/yr'),
    'ExpUtilitiesPerUnit': (0, 1e5, False, '$/unit/yr'), 'ExpPayrollPerUnit': (0, 1e5, False, '$/unit/yr'),
    'ExpRMPerUnit': (0, 1e5, False, '$/unit/yr'), 'ExpTurnoverPerUnit': (0, 1e5, False, '$/unit/yr'),
    'ExpContractPerUnit': (0, 1e5, False, '$/unit/yr'), 'ExpAdminPerUnit': (0, 1e5, False, '$/unit/yr'),
    'ExpAdvertisingPerUnit': (0, 1e5, False, '$/unit/yr'), 'MgmtFeePct': (0, 0.10, False, '% EGI'),
    'CapReservePerUnit': (0, 1e5, False, '$/unit/yr'), 'TaxPrepPerUnit': (0, 1e5, False, '$/unit/yr'),
    'AssetMgmtFeePct': (0, 0.10, False, '% EGI'), 'RefiLagMonths': (0, 24, True, 'months'),
    'IntConvMonthOverride': (0, 360, True, 'mo (0=auto)'),
}
for _d in ['DeliveryLagMonths', 'Dur_OpenEscrow', 'Dur_DueDiligence', 'Dur_Entitlement', 'Dur_PreConstruction',
           'Dur_InitialClosing', 'Dur_Construction', 'Dur_FirstUnitDelivery', 'Dur_Stabilization', 'Dur_Recap',
           'Dur_Disposition']:
    VALID[_d] = (0, 120, True, 'months')

# nicer display labels
LABELS = {
    'InvestmentType': 'Investment Type', 'PropertyName': 'Property Name', 'ProductType': 'Product Type',
    'InceptionDate': 'Open Escrow Date', 'HoldPeriodMonths': 'Hold Period (mo)', 'EntryCapRate': 'Entry Cap Rate',
    'CapAdjLocation': 'Cap Adj — Location', 'CapAdjSize': 'Cap Adj — Size', 'CapAdjAging': 'Cap Adj — Aging',
    'ExitCapResidential': 'Exit Cap — Residential', 'ExitCapRetail': 'Exit Cap — Retail', 'GrowthRent': 'Rent Growth',
    'GrowthLowIncRent': 'Low-Income Rent Growth', 'GrowthExpense': 'Expense Growth', 'GrowthUtility': 'Utility Growth',
    'GrowthTax': 'Tax Growth', 'GrowthOther': 'Other-Income Growth', 'GrowthStartDate': 'Growth Start Date',
    'LeaseUpIncomeOffset': 'Lease-Up Income', 'UnitsDeliveredPerMo': 'Units Delivered / Mo',
    'UnitsLeasedPerMo': 'Units Leased / Mo', 'LeaseUpRampPct': 'Lease-Up Opex Ramp', 'LeaseUpElevationPct': 'Lease-Up Opex Uplift',
    'StabilizedOccupancy': 'Stabilized Occupancy', 'DeliveryLagMonths': 'Delivery Lag (mo)',
    'Dur_OpenEscrow': 'Open Escrow (mo)', 'Dur_DueDiligence': 'Due Diligence (mo)', 'Dur_Entitlement': 'Entitlement (mo)',
    'Dur_PreConstruction': 'Pre-Construction (mo)', 'Dur_InitialClosing': 'Initial Closing (mo)',
    'Dur_Construction': 'Construction (mo)', 'Dur_FirstUnitDelivery': '1st Unit Delivery (mo)',
    'Dur_Stabilization': 'Stabilization (mo)', 'Dur_Recap': 'Recap / Refi (mo)', 'Dur_Disposition': 'Disposition (mo)',
    'FundingOrderPari': 'Funding Order', 'TargetLTC': 'Target LTC', 'OpReserveCoverage': 'Op Reserve Coverage',
    'OpReserveMinMonths': 'Op Reserve Min Months', 'ConLoanTermMonths': 'Con Loan Term (mo)', 'ConRate': 'Con Rate',
    'ConActual360': 'Con Interest Calc', 'ConPointsPct': 'Con Points', 'ConOtherCosts': 'Con Other Costs',
    'ConMinDebtYield': 'Con Min Debt Yield', 'ConMinDSCR': 'Con Min DSCR', 'LenderReqReserveOverride': 'Lender Reserve Override',
    'RefiLagMonths': 'Refi Lag (mo)', 'InterestMethodCurrentPay': 'Interest Method', 'IntConvMonthOverride': 'Interest Conversion Mo',
    'PermMaxLTV': 'Perm Max LTV', 'PermRate': 'Perm Rate', 'PermTermMonths': 'Perm Term (mo)', 'PermIOMonths': 'Perm I/O (mo)',
    'PermAmortMonths': 'Perm Amort (mo)', 'PermActual360': 'Perm Interest Calc', 'PermPointsPct': 'Perm Points',
    'PermMinDebtYield': 'Perm Min Debt Yield', 'PermMinDSCR': 'Perm Min DSCR', 'NoCashOut': 'Cash-Out at Refi',
    'RefiMonthOverride': 'Refi Month Override', 'DevFeePctDirect': 'Dev Fee % of Direct',
    'DevFeeThirdPartySplit': '3rd-Party Fee Split', 'SponsorDispoFeePct': 'Sponsor Dispo Fee', 'RecourseFeePct': 'Recourse Fee',
    'BrokerCostPct': 'Broker / Sale Cost', 'SalesTaxPct': 'Transfer / Sales Tax', 'TaxProrationMonths': 'Tax Proration (mo)',
    'TaxLevyPct': 'Tax Levy %', 'TaxAssessmentPct': 'Assessment Ratio %', 'TaxValueAdjFactor': 'Value-Adj Factor',
    'TaxMethodValue': 'Post-Stab Tax Method', 'TaxYear1Override': 'Year-1 Tax Override', 'ReassessOnSale': 'Dispo Tax Driver',
    'DispoTaxRatePct': 'Buyer Tax Rate', 'TaxCapitalizeLeaseUp': 'Lease-Up Tax', 'LandCost': 'Land Cost',
    'GrossAcres': 'Gross Acres', 'NetAcres': 'Net Acres', 'FARRatio': 'FAR Ratio', 'EfficiencyRatio': 'Efficiency Ratio',
    'ExistingStructureCost': 'Existing Structure Cost', 'BadDebtPct': 'Bad Debt %', 'ConcessionsPct': 'Concessions %',
    'LossToLeasePct': 'Loss to Lease %', 'ModelUnits': 'Model / Down Units', 'ParkingPerUnitMo': 'Parking $/unit/mo',
    'RetailIncomeMo': 'Retail Income $/mo', 'UtilReimbPerUnitMo': 'Utility Reimb $/unit/mo',
    'OtherIncPerUnitMo': 'Other Income $/unit/mo', 'ExpInsurancePerUnit': 'Insurance $/unit/yr',
    'ExpUtilitiesPerUnit': 'Utilities $/unit/yr', 'ExpPayrollPerUnit': 'Payroll $/unit/yr', 'ExpRMPerUnit': 'R&M $/unit/yr',
    'ExpTurnoverPerUnit': 'Turnover $/unit/yr', 'ExpContractPerUnit': 'Contract Svc $/unit/yr',
    'ExpAdminPerUnit': 'Admin & Legal $/unit/yr', 'ExpAdvertisingPerUnit': 'Advertising $/unit/yr',
    'MgmtFeePct': 'Management Fee %', 'CapReservePerUnit': 'Capital Reserve $/unit/yr', 'TaxPrepPerUnit': 'Tax Prep $/unit/yr',
    'AssetMgmtFeePct': 'Asset Mgmt Fee %',
}

# ---- column-block layout: list of (section_title, [input names]) --------------
BLOCK_A = [
    ('Property Details', ['InvestmentType', 'PropertyName', 'City', 'State', 'Market', 'ProductType', 'Location', 'InceptionDate']),
    ('Investment Details', ['HoldPeriodMonths', 'EntryCapRate']),
    ('Cap Rate & Adjustments', ['CapAdjLocation', 'CapAdjSize', 'CapAdjAging', 'ExitCapResidential', 'ExitCapRetail']),
    ('Default Growth Rates', ['GrowthRent', 'GrowthLowIncRent', 'GrowthExpense', 'GrowthUtility', 'GrowthTax', 'GrowthOther', 'GrowthStartDate']),
    ('Unit Delivery & Lease-Up', ['LeaseUpIncomeOffset', 'UnitsDeliveredPerMo', 'UnitsLeasedPerMo', 'LeaseUpRampPct', 'LeaseUpElevationPct', 'StabilizedOccupancy']),
]
BLOCK_B = [
    ('Timing Milestone Durations', ['DeliveryLagMonths', 'Dur_OpenEscrow', 'Dur_DueDiligence', 'Dur_Entitlement', 'Dur_PreConstruction', 'Dur_InitialClosing', 'Dur_Construction', 'Dur_FirstUnitDelivery', 'Dur_Stabilization', 'Dur_Recap', 'Dur_Disposition']),
    ('Capitalization & Funding', ['FundingOrderPari', 'TargetLTC', 'OpReserveCoverage', 'OpReserveMinMonths']),
    ('Construction Loan', ['ConLoanTermMonths', 'ConRate', 'ConActual360', 'ConPointsPct', 'ConOtherCosts', 'ConMinDebtYield', 'ConMinDSCR', 'LenderReqReserveOverride']),
    ('Construction Interest & Refi', ['RefiLagMonths', 'InterestMethodCurrentPay', 'IntConvMonthOverride']),
]
BLOCK_C = [
    ('Permanent Loan (Refinance)', ['PermMaxLTV', 'PermRate', 'PermTermMonths', 'PermIOMonths', 'PermAmortMonths', 'PermActual360', 'PermPointsPct', 'PermMinDebtYield', 'PermMinDSCR', 'NoCashOut', 'RefiMonthOverride']),
    ('Sponsor Fees', ['DevFeePctDirect', 'DevFeeThirdPartySplit', 'SponsorDispoFeePct', 'RecourseFeePct']),
    ('Disposition & Sale', ['BrokerCostPct', 'SalesTaxPct', 'TaxProrationMonths']),
    ('Property Taxes', ['TaxLevyPct', 'TaxAssessmentPct', 'TaxValueAdjFactor', 'TaxMethodValue', 'TaxYear1Override', 'ReassessOnSale', 'DispoTaxRatePct', 'TaxCapitalizeLeaseUp']),
]
BLOCK_D = [
    ('Site Details', ['LandCost', 'GrossAcres', 'NetAcres', 'FARRatio', 'EfficiencyRatio', 'ExistingStructureCost']),
    ('Stabilized Revenue (Other Income)', ['BadDebtPct', 'ConcessionsPct', 'LossToLeasePct', 'ModelUnits', 'ParkingPerUnitMo', 'RetailIncomeMo', 'UtilReimbPerUnitMo', 'OtherIncPerUnitMo']),
    ('Stabilized Operating Expenses (Annual)', ['ExpInsurancePerUnit', 'ExpUtilitiesPerUnit', 'ExpPayrollPerUnit', 'ExpRMPerUnit', 'ExpTurnoverPerUnit', 'ExpContractPerUnit', 'ExpAdminPerUnit', 'ExpAdvertisingPerUnit', 'MgmtFeePct', 'CapReservePerUnit', 'TaxPrepPerUnit', 'AssetMgmtFeePct']),
]
BLOCKS = [('B', BLOCK_A), ('F', BLOCK_B), ('J', BLOCK_C), ('N', BLOCK_D)]
HDR_ROW = 8  # sections start here


def col_off(base, n):
    return get_column_letter(openpyxl.utils.column_index_from_string(base) + n)


def build(wb):
    ip = wb['Inputs']
    # capture current value + number_format for every named input
    cur = {}
    for nm, dn in list(wb.defined_names.items()):
        for t, c in dn.destinations:
            if t == 'Inputs':
                cell = ip[c.replace('$', '')]
                cur[nm] = (cell.value, cell.number_format)
    # resolve each toggle's CURRENT 1/0 state from the display text (column C of its
    # row) mapped via the old true-label, then pick the new display label.
    OLD_TRUE = {
        'LeaseUpIncomeOffset': 'Offset vs equity', 'TaxMethodValue': 'Value Approach',
        'ReassessOnSale': 'Yes — reassess at sale', 'FundingOrderPari': 'Pari-Passu',
        'ConActual360': 'Actual/360', 'PermActual360': 'Actual/360',
        'NoCashOut': 'No Cash-Out (solve LTV to zero)', 'TaxCapitalizeLeaseUp': 'Capitalize in budget',
        'InterestMethodCurrentPay': 'Current Pay',
    }
    tog_display = {}
    for nm, (tl, fl) in TOGGLES.items():
        dn = wb.defined_names.get(nm)
        state = 0
        if dn:
            for t, c in dn.destinations:
                row = int(''.join(ch for ch in c if ch.isdigit()))
                disp = ip.cell(row, 3).value  # display text in column C
                state = 1 if disp == OLD_TRUE.get(nm) else 0
                break
        tog_display[nm] = tl if state == 1 else fl

    # ---- wipe the sheet ----
    ip.delete_rows(1, ip.max_row + 5)
    ip.data_validations.dataValidation = []
    ip.conditional_formatting._cf_rules.clear()
    ip.sheet_view.showGridLines = False

    # column widths
    widths = {'A': 2}
    for base in ['B', 'F', 'J', 'N', 'R']:
        widths[base] = 26; widths[col_off(base, 1)] = 13; widths[col_off(base, 2)] = 8
        widths[col_off(base, 3)] = 2
    for c, w in widths.items():
        ip.column_dimensions[c].width = w

    def put(coord, v, font=None, fill=None, num=None, align=None):
        cc = ip[coord]
        cc.value = v
        if font: cc.font = font
        if fill: cc.fill = fill
        if num: cc.number_format = num
        if align: cc.alignment = align
        return cc

    def repoint(name, ref):
        if name in wb.defined_names: del wb.defined_names[name]
        wb.defined_names.add(DefinedName(name, attr_text=ref))

    # ---- title + KPI strip ----
    put('B1', '="Multifamily Development Model  —  Master Input   |   "&PropertyName', F_TITLE)
    for col in range(2, 18): ip.cell(1, col).fill = FILL_TITLE
    kpis = [  # (row, col_label, label, formula, fmt)
        (3, 'B', 'Levered IRR (net)', '=LevIRR', FMT_PCT), (3, 'E', 'Levered MOIC', '=LevMOIC', FMT_MULT),
        (3, 'H', 'Untrended YOC', '=YieldOnCost', FMT_PCT), (3, 'K', 'Dev Spread', '=DevSpreadBps', FMT_BPS),
        (3, 'N', 'TDC', '=TDC_Total', FMT_MONEY), (3, 'Q', 'TDC / Unit', '=TDCperUnit', FMT_MONEY),
        (6, 'B', 'Unlevered IRR (net)', '=UnlevIRR', FMT_PCT), (6, 'E', 'Unlevered MOIC', '=UnlevMOIC', FMT_MULT),
        (6, 'H', 'Static Value', '=StaticValue', FMT_MONEY), (6, 'K', 'Realized LTC', '=LTC_Realized', FMT_PCT),
        (6, 'N', 'Debt Yield (stab)', '=DebtYieldStab', FMT_PCT), (6, 'Q', 'Model Status', '=IF(ErrorsFound=0,"PASS","REVIEW")', '@'),
    ]
    for r, c, lab, f, fmt in kpis:
        put(f'{c}{r}', lab, F_KPIL, FILL_KPI, align=LEFT)
        vc = put(f'{c}{r+1}', f, F_KPIV, FILL_KPI, num=fmt, align=LEFT)
        for cc in (c, col_off(c, 1)):
            ip[f'{cc}{r}'].fill = FILL_KPI; ip[f'{cc}{r+1}'].fill = FILL_KPI

    # ---- section blocks ----
    sanity_ref = []
    for base, block in BLOCKS:
        lc = base; vc = col_off(base, 1); rc = col_off(base, 2)
        row = HDR_ROW
        for title, keys in block:
            put(f'{lc}{row}', title, F_SECT, FILL_SECT, align=LEFT)
            for cc in (vc, rc): ip[f'{cc}{row}'].fill = FILL_SECT
            row += 1
            for key in keys:
                label = LABELS.get(key, key)
                put(f'{lc}{row}', label, F_LBL, align=LEFT)
                value, numfmt = cur.get(key, (0, 'General'))
                if key in TOGGLES:
                    tl, fl = TOGGLES[key]
                    put(f'{vc}{row}', tog_display[key], F_IN, FILL_IN, align=LEFT)
                    dv = DataValidation(type='list', formula1=f'"{tl},{fl}"', allow_blank=False, showErrorMessage=True)
                    dv.promptTitle = label; dv.prompt = f'{tl} = 1 · {fl} = 0'
                    ip.add_data_validation(dv); dv.add(ip[f'{vc}{row}'])
                    put(f'{rc}{row}', f'=IF({vc}{row}="{tl}",1,0)', F_RESV, num=FMT_INT, align=CENTER)
                    repoint(key, f"'Inputs'!${rc}${row}")
                elif key in FORMULA_INPUTS:
                    put(f'{vc}{row}', value, F_LINK, num=numfmt, align=RIGHT)
                    repoint(key, f"'Inputs'!${vc}${row}")
                else:
                    isdate = key in DATE_INPUTS
                    put(f'{vc}{row}', value, F_IN, FILL_IN, num=numfmt,
                        align=(LEFT if (isinstance(value, str) or isdate) else RIGHT))
                    repoint(key, f"'Inputs'!${vc}${row}")
                    if key in TEXT_LISTS:
                        dv = DataValidation(type='list', formula1='"' + ",".join(TEXT_LISTS[key]) + '"', allow_blank=True)
                        dv.promptTitle = label; ip.add_data_validation(dv); dv.add(ip[f'{vc}{row}'])
                    elif key in VALID:
                        lo, hi, whole, unit = VALID[key]
                        dv = DataValidation(type=('whole' if whole else 'decimal'), operator='between',
                                            formula1=repr(lo), formula2=repr(hi), allow_blank=False,
                                            showErrorMessage=True, showInputMessage=True)
                        dv.errorTitle = 'Out of range'; dv.error = f'{label}: enter {unit}.'
                        dv.promptTitle = label; dv.prompt = f'Unit: {unit}.  Range: {lo}–{hi}.'
                        ip.add_data_validation(dv); dv.add(ip[f'{vc}{row}'])
                row += 1
            row += 1  # gap between sections

    # ---- derived / static-value / sanity block (column R) ----
    base = 'R'; lc = base; vc = col_off(base, 1); rc = col_off(base, 2)
    row = HDR_ROW
    put(f'{lc}{row}', 'Derived & Static Value', F_SECT, FILL_SECT, align=LEFT)
    for cc in (vc, rc): ip[f'{cc}{row}'].fill = FILL_SECT
    row += 1
    derived = [
        ('Effective Tax Rate on Value', '=TaxLevyPct*TaxAssessmentPct*TaxValueAdjFactor', '0.000%', None),
        ('Stabilized Tax / Unit / Yr', '=IF(TotalUnits=0,0,StabTaxAnnual/TotalUnits)', '$#,##0', None),
        ('Static Value (Untrended NOI ÷ Entry Cap)', '=IF(EntryCapRate=0,0,StabNOI_Untrended/EntryCapRate)', FMT_MONEY, 'StaticValue'),
        ('Static Value / Unit', '=IF(TotalUnits=0,0,StaticValue/TotalUnits)', FMT_MONEY, None),
        ('Density (units / net acre)', '=IF(NetAcres=0,0,TotalUnits/NetAcres)', '#,##0.0', None),
        ('Max Buildable SF by FAR', '=FARRatio*NetAcres*43560', FMT_INT, None),
    ]
    for lab, f, fmt, nm in derived:
        put(f'{lc}{row}', lab, F_LBL, align=LEFT)
        put(f'{vc}{row}', f, F_LINK, num=fmt, align=RIGHT)
        if nm: repoint(nm, f"'Inputs'!${vc}${row}")
        row += 1
    row += 1
    put(f'{lc}{row}', 'Input Sanity (soft warnings)', F_SECT, FILL_SECT, align=LEFT)
    for cc in (vc, rc): ip[f'{cc}{row}'].fill = FILL_SECT
    row += 1
    sanity = [
        ('Construction term ≤ 48 mo', '=ConLoanTermMonths', FMT_INT, '=IF(ConLoanTermMonths<=48,"OK","CHECK")'),
        ('Stab tax/unit $500–$4,000', '=IF(TotalUnits=0,0,StabTaxAnnual/TotalUnits)', '$#,##0', '=IF(AND(StabTaxAnnual/TotalUnits>=500,StabTaxAnnual/TotalUnits<=4000),"OK","CHECK")'),
        ('Absorption ≤ delivery pace', '=UnitsLeasedPerMo', FMT_INT, '=IF(UnitsLeasedPerMo<=UnitsDeliveredPerMo,"OK","CHECK")'),
        ('Untrended YOC ≥ exit cap', '=YieldOnCost-ExitCapResidential', FMT_PCT, '=IF(YieldOnCost>=ExitCapResidential,"OK","CHECK")'),
        ('Occupancy 85%–98%', '=StabilizedOccupancy', FMT_PCT1, '=IF(AND(StabilizedOccupancy>=0.85,StabilizedOccupancy<=0.98),"OK","CHECK")'),
        ('Rent growth ≤ 5%', '=GrowthRent', FMT_PCT, '=IF(GrowthRent<=0.05,"OK","CHECK")'),
    ]
    s_first = row
    for lab, cval, fmt, dform in sanity:
        put(f'{lc}{row}', lab, F_LBL, align=LEFT)
        put(f'{vc}{row}', cval, F_LINK, num=fmt, align=RIGHT)
        put(f'{rc}{row}', dform, F_AMBER, align=CENTER)
        row += 1
    s_last = row - 1
    ip.conditional_formatting.add(f'{rc}{s_first}:{rc}{s_last}',
                                  CellIsRule(operator='equal', formula=['"CHECK"'], font=F_AMBER, fill=AMBER_F))
    put(f'{lc}{row}', 'Sanity checks flagged', F_LBL, align=LEFT)
    put(f'{vc}{row}', f'=COUNTIF({rc}{s_first}:{rc}{s_last},"CHECK")', F_LINK, num=FMT_INT, align=RIGHT)
    repoint('SanityFlags', f"'Inputs'!${vc}${row}")
    repoint('InputSanityWarnings', f"'Inputs'!${rc}${s_first}")

    # ---- update the one external address ref: Diagnostics error scan of Inputs ----
    dg = wb['Diagnostics']
    # scan all input value columns (C, G, K, O) + derived (S) — numeric where present
    dg['F17'].value = ('=SUMPRODUCT(--ISERROR(Inputs!C9:C40))+SUMPRODUCT(--ISERROR(Inputs!G9:G40))'
                       '+SUMPRODUCT(--ISERROR(Inputs!K9:K40))+SUMPRODUCT(--ISERROR(Inputs!O9:O40))'
                       '+SUMPRODUCT(--ISERROR(Inputs!S9:S30))')

    # ---- surface Static Value on the Dashboard (under Cost & Basis, right col) ----
    db = wb['Dashboard']
    # find a free row after Exit Basis / Unit (row 11 in the Phase C layout)
    db['E12'].value = 'Static Value (untrended)'; db['E12'].font = F_LBL; db['E12'].alignment = LEFT
    db['F12'].value = '=StaticValue'; db['F12'].font = F_LINK; db['F12'].number_format = FMT_MONEY; db['F12'].alignment = RIGHT

    ip.freeze_panes = 'A8'
    return {'kpis': len(kpis), 'sections': sum(len(b) for _, b in BLOCKS) + 2}


def main():
    src, dst = sys.argv[1], sys.argv[2]
    wb = openpyxl.load_workbook(src)
    print('applied:', build(wb))
    wb.save(dst)
    print('saved', dst)


if __name__ == '__main__':
    main()
