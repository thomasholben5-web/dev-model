"""
remediate.py — applies the numbered remediation fixes to the user's edited
workbook (Claude_Dev_Model.xlsx = "The Verde" deal).  The original generator
describes a different deal (Meridian) and writes inputs in a different column
layout, so the user's workbook is patched directly rather than regenerated.

Each fix is a function; run.py-style CLI applies a subset so validation can be
checkpointed. Every added value is a formula referencing existing named ranges
or cells — no hard-keyed calculated outputs.

Usage:  python3 remediate.py <src.xlsx> <dst.xlsx> fix1[,fix2,...|all]
"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

# ---- house style ------------------------------------------------------------
NAVY = "FF1F3864"; INK = "FF000000"; BLUE_IN = "FF0000CC"; GREEN_LK = "FF006100"
FILL_IN = "FFFDF3D0"; FILL_TOT = "FFE2EFDA"; FILL_SUB = "FFF2F2F2"
F_TITLE = Font(name="Calibri", size=16, bold=True, color=NAVY)
F_SUB   = Font(name="Calibri", size=10, italic=True, color="FF595959")
F_SECT  = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
F_LBL   = Font(name="Calibri", size=10, color=INK)
F_LBLB  = Font(name="Calibri", size=10, bold=True, color=INK)
F_LINK  = Font(name="Calibri", size=10, color=GREEN_LK)
F_FORM  = Font(name="Calibri", size=10, color=INK)
F_TOT   = Font(name="Calibri", size=10, bold=True, color=INK)
F_HDR   = Font(name="Calibri", size=9, bold=True, color="FFFFFFFF")
F_IN    = Font(name="Calibri", size=10, color=BLUE_IN)
FILL_SECT  = PatternFill("solid", fgColor=NAVY)
FILL_TOTAL = PatternFill("solid", fgColor=FILL_TOT)
FILL_BAND  = PatternFill("solid", fgColor=FILL_SUB)
FILL_HDR   = PatternFill("solid", fgColor="FF2E5496")
FILL_INPUT = PatternFill("solid", fgColor=FILL_IN)
FILL_REF   = PatternFill("solid", fgColor="FFF7F7F7")
RIGHT = Alignment(horizontal="right"); LEFT = Alignment(horizontal="left"); CENTER = Alignment(horizontal="center")
_thin = Side(style="thin", color="FFBFBFBF")
TOPB = Border(top=Side(style="thin", color="FF808080"))
FMT_MONEY = '#,##0;(#,##0)'; FMT_PSF = '$#,##0.00;($#,##0.00)'; FMT_UMO = '#,##0;(#,##0)'
FMT_PCT = '0.00%'; FMT_AMT = '$#,##0'; FMT_INT = '#,##0'


def S(ws, coord, value=None, font=None, fill=None, num=None, align=None, border=None):
    c = ws[coord]
    if value is not None: c.value = value
    if font: c.font = font
    if fill: c.fill = fill
    if num: c.number_format = num
    if align: c.alignment = align
    if border: c.border = border
    return c


# =====================================================================  FIX 1
def fix_1(wb):
    """Untrended stabilized proforma drives development metrics; rebuild Proforma."""
    if 'Proforma' in wb.sheetnames:
        del wb['Proforma']
    ws = wb.create_sheet('Proforma')
    wb.move_sheet('Proforma', -(len(wb.sheetnames) - 2))  # after Dashboard
    for col, w in {'A': 3.7, 'B': 32, 'C': 14, 'D': 11, 'E': 12, 'F': 13, 'G': 14, 'H': 15}.items():
        ws.column_dimensions[col].width = w
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = '375623'

    S(ws, 'B1', 'Stabilized Operating Proforma — Untrended', F_TITLE)
    S(ws, 'B2', '=""&PropertyName&"  •  "&City&", "&State&"  •  point-in-time at stabilized occupancy (no growth trending)"', F_SUB)

    # context band
    def band(coord, text):
        S(ws, coord, text, F_SECT, FILL_SECT, align=LEFT)
    band('B4', 'Basis')
    for col in 'CDEFGH': ws[f'{col}4'].fill = FILL_SECT
    ctx = [(5,'Total Units','=TotalUnits',FMT_INT,'Occupied Units (stab)','=StabLeasedTarget',FMT_INT),
           (6,'Net Rentable SF','=NRSF',FMT_INT,'Stabilized Occupancy','=StabilizedOccupancy',FMT_PCT),
           (7,'Avg Unit SF','=AvgUnitSF',FMT_INT,'Effective Tax Rate on Value','=EffTaxRateOnValue','0.000%')]
    for rr,l1,f1,fmt1,l2,f2,fmt2 in ctx:
        S(ws,f'B{rr}',l1,F_LBL,align=LEFT); S(ws,f'C{rr}',f1,F_LINK,num=fmt1,align=RIGHT)
        S(ws,f'E{rr}',l2,F_LBL,align=LEFT); S(ws,f'F{rr}',f2,F_LINK,num=fmt2,align=RIGHT)

    HDR = 10
    heads = ['Untrended Stabilized Proforma','% of / amt','$ / SF /mo','$ / Unit /mo','$ / Month','$ / Year','Trended $/yr (ref)']
    for i,h in enumerate(heads):
        S(ws, f'{get_column_letter(2+i)}{HDR}', h, F_HDR, FILL_HDR, align=(LEFT if i==0 else RIGHT))

    def opsum(r):  # trended fwd-12mo (reference column)
        return (f'=SUM(INDEX(Operating!$E${r}:$NL${r},StabMonth):INDEX(Operating!$E${r}:$NL${r},StabMonth+11))')

    # rows: (kind, label, untrended G-formula, driver, driver_fmt, trended_op_row)
    # G-formula computes annual untrended $; {EGI},{NTOP} placeholders resolved later.
    rows = [
        ('SECT','Revenue',None,None,None,None),
        ('line','Gross Potential Rent','=(BaseMarketRentMo+BaseLowRentMo)*12','=(BaseMarketRentMo+BaseLowRentMo)/TotalUnits',FMT_AMT,14),
        ('line','Vacancy Loss','=-{GPR}*(1-StabilizedOccupancy)','=-(1-StabilizedOccupancy)',FMT_PCT,15),
        ('line','Gain / (Loss) to Lease','=-{GPR}*LossToLeasePct','=LossToLeasePct',FMT_PCT,16),
        ('line','Bad Debt','=-{GPR}*BadDebtPct','=BadDebtPct',FMT_PCT,17),
        ('line','Concessions','=-{GPR}*ConcessionsPct','=ConcessionsPct',FMT_PCT,18),
        ('line','Model / Down Units','=-{GPR}*ModelUnits/TotalUnits','=ModelUnits',FMT_INT,None),
        ('line','Parking & Storage','=ParkingPerUnitMo*StabLeasedTarget*12','=ParkingPerUnitMo',FMT_AMT,19),
        ('line','Utility Reimbursement','=UtilReimbPerUnitMo*StabLeasedTarget*12','=UtilReimbPerUnitMo',FMT_AMT,20),
        ('line','Other Income','=OtherIncPerUnitMo*StabLeasedTarget*12','=OtherIncPerUnitMo',FMT_AMT,21),
        ('line','Retail Income','=RetailRentMo*StabilizedOccupancy*12','=RetailRentMo',FMT_AMT,22),
        ('egi','Effective Gross Income',None,None,None,23),
        ('SECT','Operating Expenses',None,None,None,None),
        ('line','Insurance','=ExpInsurancePerUnit*TotalUnits','=ExpInsurancePerUnit',FMT_AMT,25),
        ('line','Utilities','=ExpUtilitiesPerUnit*TotalUnits','=ExpUtilitiesPerUnit',FMT_AMT,27),
        ('line','Payroll','=ExpPayrollPerUnit*TotalUnits','=ExpPayrollPerUnit',FMT_AMT,31),
        ('line','Management Fee','=MgmtFeePct*{EGI}','=MgmtFeePct',FMT_PCT,33),
        ('line','Repairs & Maintenance','=ExpRMPerUnit*TotalUnits','=ExpRMPerUnit',FMT_AMT,28),
        ('line','Turnover','=ExpTurnoverPerUnit*TotalUnits','=ExpTurnoverPerUnit',FMT_AMT,29),
        ('line','Contract Services','=ExpContractPerUnit*TotalUnits','=ExpContractPerUnit',FMT_AMT,30),
        ('line','Admin & Legal','=ExpAdminPerUnit*TotalUnits','=ExpAdminPerUnit',FMT_AMT,26),
        ('line','Advertising','=ExpAdvertisingPerUnit*TotalUnits','=ExpAdvertisingPerUnit',FMT_AMT,32),
        ('tax','Property Taxes',None,'=EffTaxRateOnValue','0.000%',36),
        ('opex','Total Operating Expenses',None,None,None,37),
        ('noi','Net Operating Income',None,None,None,38),
        ('SECT','Below-the-Line Expenses',None,None,None,None),
        ('line','Capital Reserves','=CapReservePerUnit*TotalUnits','=CapReservePerUnit',FMT_AMT,None),
        ('line','Tax Prep & Appraisal','=TaxPrepPerUnit*TotalUnits','=TaxPrepPerUnit',FMT_AMT,None),
        ('line','Asset Management Fees','=AssetMgmtFeePct*{EGI}','=AssetMgmtFeePct',FMT_PCT,None),
        ('belowtot','Total Below-the-Line',None,None,None,None),
        ('cf','Proforma Operating Cash Flow',None,None,None,42),
    ]
    r = HDR + 1
    ref = {}  # label -> row
    exp_first = exp_last = None
    below_first = below_last = None
    for kind,label,gform,drv,dfmt,oprow in rows:
        if kind == 'SECT':
            band(f'B{r}', label)
            for col in 'CDEFGH': ws[f'{col}{r}'].fill = FILL_SECT
            r += 1; continue
        bold = kind in ('egi','opex','noi','cf','belowtot')
        S(ws, f'B{r}', label, F_TOT if bold else F_LBL, align=LEFT)
        if drv: S(ws, f'C{r}', drv, F_LINK, num=dfmt, align=RIGHT)
        ref[label] = r
        # track expense block for tax + total
        if kind == 'line' and label in ('Insurance','Utilities','Payroll','Management Fee','Repairs & Maintenance',
                                         'Turnover','Contract Services','Admin & Legal','Advertising'):
            exp_first = exp_first or r; exp_last = r
        if kind == 'line' and label in ('Capital Reserves','Tax Prep & Appraisal','Asset Management Fees'):
            below_first = below_first or r; below_last = r
        # resolve G formula
        gpr = ref.get('Gross Potential Rent'); egi = ref.get('Effective Gross Income')
        if kind == 'egi':
            g = f'=SUM(G{ref["Gross Potential Rent"]}:G{r-1})'
        elif kind == 'tax':
            g = (f'=IF(TaxYear1Override>0,TaxYear1Override,'
                 f'(G{egi}-SUM(G{exp_first}:G{exp_last}))/(EntryCapRate+EffTaxRateOnValue)*EffTaxRateOnValue)')
        elif kind == 'opex':
            g = f'=SUM(G{exp_first}:G{ref["Property Taxes"]})'
        elif kind == 'noi':
            g = f'=G{egi}-G{ref["Total Operating Expenses"]}'
        elif kind == 'belowtot':
            g = f'=SUM(G{below_first}:G{below_last})'
        elif kind == 'cf':
            g = f'=G{ref["Net Operating Income"]}-G{ref["Total Below-the-Line"]}'
        else:
            g = gform.replace('{GPR}', f'G{gpr}').replace('{EGI}', f'G{egi}')
        hfont = F_TOT if bold else F_FORM
        S(ws, f'G{r}', g, hfont, num=FMT_MONEY, align=RIGHT)
        S(ws, f'D{r}', f'=IF(NRSF=0,0,G{r}/12/NRSF)', hfont, num=FMT_PSF, align=RIGHT)
        S(ws, f'E{r}', f'=IF(TotalUnits=0,0,G{r}/12/TotalUnits)', hfont, num=FMT_UMO, align=RIGHT)
        S(ws, f'F{r}', f'=G{r}/12', hfont, num=FMT_MONEY, align=RIGHT)
        # trended reference column
        if oprow:
            S(ws, f'H{r}', opsum(oprow), F_FORM, num=FMT_MONEY, align=RIGHT)
            ws[f'H{r}'].fill = FILL_REF
        if bold:
            for col in 'BCDEFG':
                ws[f'{col}{r}'].fill = FILL_TOTAL if kind != 'noi' else FILL_BAND
                ws[f'{col}{r}'].border = TOPB
            if kind == 'noi':
                for col in 'BCDEFG': ws[f'{col}{r}'].font = F_TOT
        r += 1

    # named ranges for untrended NOI (drives dev metrics)
    noi_row = ref['Net Operating Income']; egi_row = ref['Effective Gross Income']
    # NOI-before-tax = EGI - non-tax opex
    nbt = f"'Proforma'!$G${egi_row}-SUM('Proforma'!$G${exp_first}:$G${exp_last})"
    wb.defined_names.add(openpyxl.workbook.defined_name.DefinedName(
        'StabNOI_Untrended', attr_text=f"'Proforma'!$G${noi_row}"))
    # StabNOIbt_Untrended -> place a helper cell
    hb = r + 1
    S(ws, f'B{hb}', 'Untrended Stabilized NOI before Property Tax', F_LBL, align=LEFT)
    S(ws, f'G{hb}', f'=G{egi_row}-SUM(G{exp_first}:G{exp_last})', F_FORM, num=FMT_MONEY, align=RIGHT)
    wb.defined_names.add(openpyxl.workbook.defined_name.DefinedName(
        'StabNOIbt_Untrended', attr_text=f"'Proforma'!$G${hb}"))

    # reconciliation / diagnostic rows
    r = hb + 2
    band(f'B{r}', 'Reconciliation');
    for col in 'CDEFGH': ws[f'{col}{r}'].fill = FILL_SECT
    rec0 = r + 1; r = rec0
    S(ws, f'B{r}', 'Untrended NOI ≤ Trended NOI (growth ≥ 0)', F_LBL, align=LEFT)
    S(ws, f'C{r}', f'=StabNOI_Untrended-StabNOI', F_FORM, num=FMT_MONEY, align=RIGHT)
    S(ws, f'D{r}', f'=IF(OR(GrowthRent<0,GrowthExpense<0),"n/a",IF(StabNOI_Untrended<=StabNOI+1,"PASS","FAIL"))', F_LBLB, align=CENTER)
    r += 1
    S(ws, f'B{r}', 'Untrended YOC (StabNOI_Untrended / TDC)', F_LBL, align=LEFT)
    S(ws, f'C{r}', f'=IF(TDC_Total=0,0,StabNOI_Untrended/TDC_Total)', F_FORM, num='0.000%', align=RIGHT)
    S(ws, f'D{r}', f'=IF(YieldOnCost-ExitCapResidential<0,"SPREAD<0","OK")', F_LBLB, align=CENTER)
    rec_last = r
    ws.freeze_panes = 'C11'

    # ---- repoint development metrics to untrended -----------------------
    R = wb['Returns']
    R['C20'].value = '=IF(TDC_Total=0,0,StabNOI_Untrended/TDC_Total)'   # YieldOnCost
    R['C24'].value = '=IF(TDC_Total=0,0,StabNOI_Untrended/TDC_Total)'   # ROCstab
    # C21/C22 DevSpread already = YieldOnCost-ExitCapResidential (auto-follows)
    # DebtYieldStab -> untrended NOI over ACTIVE loan (FIX 8 also)
    R['C25'].value = '=IF(RefiFlag=1,IF(PermLoan=0,0,StabNOI_Untrended/PermLoan),IF(ConLoanTotal=0,0,StabNOI_Untrended/ConLoanTotal))'
    return {'proforma_rows': ref, 'rec_range': (rec0, rec_last)}


# =====================================================================  FIX 2
def fix_2(wb):
    """Refinance engine live: realistic construction term + input validation."""
    ip = wb['Inputs']
    # ConLoanTermMonths is at Inputs!C62 (label B62). Was =40*12 (=480) -> 36.
    cell = None
    for r in range(1, ip.max_row + 1):
        if ip.cell(r, 2).value == 'Con Loan Term Months':
            cell = f'C{r}'; break
    assert cell, 'ConLoanTermMonths label not found'
    ip[cell].value = 36
    ip[cell].font = F_IN
    ip[cell].fill = FILL_INPUT
    ip[cell].number_format = FMT_INT
    # input validation: whole number 12..60, warning if exceeded
    dv = DataValidation(type='whole', operator='between', formula1='12', formula2='60',
                        allow_blank=False, showErrorMessage=True)
    dv.errorTitle = 'Construction term out of range'
    dv.error = ('Construction-loan term is typically 12-60 months. Values above 60 keep the '
                'loan interest-only for the full hold and suppress the refinance. Enter 12-60.')
    dv.promptTitle = 'Construction Loan Term (months)'
    dv.prompt = 'Realistic range 12-60. Drives construction maturity and the refi-or-sell switch.'
    ip.add_data_validation(dv)
    dv.add(ip[cell])
    return {'con_term_cell': cell}


FIRST_MONTH_COL = 5   # column E
LAST_MONTH_COL = 376  # column NL

def _monthcols():
    for m in range(1, LAST_MONTH_COL - FIRST_MONTH_COL + 2):
        yield m, get_column_letter(FIRST_MONTH_COL + m - 1)


def _add_input_row(ip, row, label, value, fmt, note, name, reg_names):
    S(ip, f'B{row}', label, F_LBL, align=LEFT)
    S(ip, f'C{row}', value, F_IN, FILL_INPUT, num=fmt, align=RIGHT)
    if note:
        c = ip.cell(row=row, column=4, value=note); c.font = F_SUB; c.alignment = LEFT
    reg_names.append((name, f'C{row}'))


def _add_named(wb, name, ref):
    if name in wb.defined_names:
        del wb.defined_names[name]
    wb.defined_names.add(openpyxl.workbook.defined_name.DefinedName(name, attr_text=ref))


# =====================================================================  FIX 7
def fix_7(wb):
    """Move hardcoded milestone durations + delivery lag to Inputs."""
    ip = wb['Inputs']; tl = wb['Timeline']
    start = ip.max_row + 2
    S(ip, f'B{start}', 'Timing Milestone Durations (months)', F_SECT, FILL_SECT, align=LEFT)
    reg = []
    rows = [
        ('DeliveryLagMonths', 13, 'Months from construction start to 1st unit delivery'),
        ('Dur_OpenEscrow', 1, 'Open Escrow duration'),
        ('Dur_DueDiligence', 4, 'Due Diligence duration'),
        ('Dur_Entitlement', 5, 'Entitlement duration'),
        ('Dur_PreConstruction', 3, 'Pre-Construction duration'),
        ('Dur_InitialClosing', 1, 'Initial Closing duration'),
        ('Dur_Construction', 22, 'Construction duration'),
        ('Dur_FirstUnitDelivery', 1, '1st Unit Delivery duration'),
        ('Dur_Stabilization', 1, 'Stabilization duration'),
        ('Dur_Recap', 1, 'Recapitalization / Refinance duration'),
        ('Dur_Disposition', 1, 'Disposition duration'),
    ]
    r = start + 1
    for name, val, note in rows:
        _add_input_row(ip, r, name, val, FMT_INT, note, name, reg); r += 1
    for name, ref in reg:
        _add_named(wb, name, f"'Inputs'!${ref[0]}${ref[1:]}")
    # repoint Timeline duration cells + delivery lag to the named inputs
    tl['D4'].value = '=Dur_OpenEscrow'
    tl['D5'].value = '=Dur_DueDiligence'
    tl['D6'].value = '=Dur_Entitlement'
    tl['D7'].value = '=Dur_PreConstruction'
    tl['D8'].value = '=Dur_InitialClosing'
    tl['D9'].value = '=Dur_Construction'
    tl['D10'].value = '=Dur_FirstUnitDelivery'
    tl['D12'].value = '=Dur_Stabilization'
    tl['D13'].value = '=Dur_Recap'
    tl['D14'].value = '=Dur_Disposition'
    tl['C10'].value = '=+MS_Construction_Start+DeliveryLagMonths'  # was +13 hardcoded
    for c in ('D4','D5','D6','D7','D8','D9','D10','D12','D13','D14','C10'):
        tl[c].font = F_FORM
    return {'inputs_added': len(rows), 'start_row': start}


# =====================================================================  FIX 4
def fix_4(wb):
    """Static (no-growth) block independent of all growth factors."""
    op = wb['Operating']
    # add growth-free static-tax helpers below the static block
    base = op.max_row + 2
    S(op, f'B{base}', 'Static (No-Growth) Property-Tax Basis', F_SECT, FILL_SECT, align=LEFT)
    r1, r2, r3 = base + 1, base + 2, base + 3
    S(op, f'B{r1}', 'Static Stabilized NOI before Tax (fwd 12mo)', F_LBL, align=LEFT)
    S(op, f'C{r1}', '=SUM(INDEX(E$68:NL$68,StabMonth):INDEX(E$68:NL$68,StabMonth+11))', F_FORM, num=FMT_MONEY, align=RIGHT)
    S(op, f'B{r2}', 'Static Assessed Value (growth-free)', F_LBL, align=LEFT)
    S(op, f'C{r2}', f'=IF((EntryCapRate+EffTaxRateOnValue)=0,0,C{r1}/(EntryCapRate+EffTaxRateOnValue))', F_FORM, num=FMT_MONEY, align=RIGHT)
    S(op, f'B{r3}', 'Static Year-1 Tax (growth-free)', F_LBL, align=LEFT)
    S(op, f'C{r3}', f'=IF(TaxYear1Override>0,TaxYear1Override,C{r2}*EffTaxRateOnValue)', F_FORM, num=FMT_MONEY, align=RIGHT)
    _add_named(wb, 'StaticTaxYear1', f"'Operating'!$C${r3}")
    # rewrite static property tax row 69 to use StaticTaxYear1 (no growth), flat
    for m, col in _monthcols():
        op[f'{col}69'].value = (
            f'=IF({m}<StabMonth,0,IF(TaxMethodValue>=1,StaticTaxYear1/12,'
            f'MAX(0,{col}$68*12)/EntryCapRate*EffTaxRateOnValue/12))')
    return {'static_tax_helpers': (r1, r3)}


# =====================================================================  FIX 3
def fix_3(wb):
    """Property-tax transparency + optional lease-up accrual toggle."""
    ip = wb['Inputs']; op = wb['Operating']; ds = wb['DevSpend']
    # 3.1 transparency: derived effective rate + per-unit tax on Inputs
    start = ip.max_row + 2
    S(ip, f'B{start}', 'Property Tax — Derived (sanity check)', F_SECT, FILL_SECT, align=LEFT)
    r = start + 1
    S(ip, f'B{r}', 'Effective Tax Rate on Value', F_LBL, align=LEFT)
    S(ip, f'C{r}', '=TaxLevyPct*TaxAssessmentPct*TaxValueAdjFactor', F_LINK, num='0.000%', align=RIGHT)
    c = ip.cell(r, 4, value='Levy × assessment ratio × value-adj factor (no double count)'); c.font = F_SUB; c.alignment = LEFT
    r += 1
    S(ip, f'B{r}', 'Stabilized Tax per Unit / Yr', F_LBL, align=LEFT)
    S(ip, f'C{r}', '=IF(TotalUnits=0,0,StabTaxAnnual/TotalUnits)', F_LINK, num=FMT_AMT, align=RIGHT)
    c = ip.cell(r, 4, value='Sanity-check band ≈ $800–$3,500 for many US multifamily markets'); c.font = F_SUB; c.alignment = LEFT
    r += 1
    # 3.2 toggle: capitalize pre-stabilization (lease-up) property tax?
    S(ip, f'B{r}', 'Capitalize Lease-Up Tax', F_LBL, align=LEFT)
    S(ip, f'C{r}', 1, F_IN, FILL_INPUT, num=FMT_INT, align=RIGHT)
    c = ip.cell(r, 4, value='1=capitalize pre-stab tax in budget (default); 0=accrue as operating from delivery'); c.font = F_SUB; c.alignment = LEFT
    _add_named(wb, 'TaxCapitalizeLeaseUp', f"'Inputs'!$C${r}")
    # derived operating-tax start month (boundary between capitalized & operating tax)
    tl = wb['Timeline']
    tstart = tl.max_row + 2
    S(tl, f'B{tstart}', 'Operating Tax Start Month (capitalization boundary)', F_LBL, align=LEFT)
    S(tl, f'C{tstart}', '=IF(TaxCapitalizeLeaseUp>=1,StabMonth,DeliverStart)', F_FORM, num=FMT_INT, align=RIGHT)
    _add_named(wb, 'TaxOpStartMonth', f"'Timeline'!$C${tstart}")
    # rewrite operating tax row 36: 0 before boundary; income-approach during lease-up; stabilized formula after
    for m, col in _monthcols():
        op[f'{col}36'].value = (
            f'=IF({m}<TaxOpStartMonth,0,'
            f'IF({m}<StabMonth,MAX(0,{col}$35*12)/EntryCapRate*EffTaxRateOnValue/12,'
            f'IF(TaxMethodValue>=1,TaxYear1*(1+GrowthTax)^(MAX(0,({m}-StabMonth)/12))/12,'
            f'MAX(0,{col}$35*12)/EntryCapRate*EffTaxRateOnValue/12)))')
    # DevSpend construction-tax gate stops at the operating-tax boundary (no double count)
    for m, col in _monthcols():
        ds[f'{col}38'].value = (
            f'=IF(AND({m}>=MS_Initial_Closing_Start,{m}<TaxOpStartMonth),'
            f'(LandCost+{col}$37)*EffTaxRateOnValue/12,0)')
    return {'tax_toggle_added': True}


# =====================================================================  FIX 6
def fix_6(wb):
    """Wire previously dead inputs into the model; keep single-source-of-truth by name."""
    ip = wb['Inputs']; op = wb['Operating']; db = wb['Dashboard']
    # (a) cap-rate adjustments -> exit cap BY NAME (was SUM(C16:C19) cell-range)
    ip['C20'].value = '=EntryCapRate+CapAdjLocation+CapAdjSize+CapAdjAging'
    # (b) InvestmentType + Location -> Dashboard subtitle
    db['B2'].value = ('=""&PropertyName&"  •  "&City&", "&State&"  •  "&ProductType'
                      '&"  •  "&InvestmentType&"  ("&Location&")"')
    # (c) NetAcres + FARRatio -> derived site metrics on Inputs
    start = ip.max_row + 2
    S(ip, f'B{start}', 'Site Density (derived)', F_SECT, FILL_SECT, align=LEFT)
    r = start + 1
    S(ip, f'B{r}', 'Density (units / net acre)', F_LBL, align=LEFT)
    S(ip, f'C{r}', '=IF(NetAcres=0,0,TotalUnits/NetAcres)', F_LINK, num='#,##0.0', align=RIGHT); r += 1
    S(ip, f'B{r}', 'Max Buildable SF by FAR', F_LBL, align=LEFT)
    S(ip, f'C{r}', '=FARRatio*NetAcres*43560', F_LINK, num=FMT_INT, align=RIGHT)
    c = ip.cell(r, 4, value='FAR × net acres × 43,560 sf/acre'); c.font = F_SUB; c.alignment = LEFT
    # (d) ModelUnits -> trended vacancy (row 15): add model-unit non-revenue loss
    for m, col in _monthcols():
        op[f'{col}15'].value = f'=-{col}$14*(1-LeaseUp!{col}$10)-{col}$14*ModelUnits/TotalUnits'
    # (e) RetailIncomeMo -> add to retail income line (row 22) on top of matrix retail
    for m, col in _monthcols():
        op[f'{col}22'].value = (
            f'=IF({m}>=DeliverStart,(RetailRentMo+RetailIncomeMo)*{col}$10*'
            f'MIN(1,LeaseUp!{col}$10/StabilizedOccupancy),0)')
    # (f) DevFeeThirdPartySplit -> sponsor keeps (1-split) of dev fee in gross-fee addback
    R = wb['Returns']
    for row in (11, 12):
        for m, col in _monthcols():
            cur = R[f'{col}{row}'].value
            if isinstance(cur, str) and 'DevFee+' in cur:
                R[f'{col}{row}'].value = cur.replace('DevFee+', 'DevFee*(1-DevFeeThirdPartySplit)+')
    # (g) LeaseUpIncomeOffset -> operating-reserve credit for recycled lease-up income
    #     when 0, reserve is grossed up by lease-up EGI (income not credited).
    op['C51'].value = (
        '=MAX(MAX(0,-PeakShort+(1-LeaseUpIncomeOffset)*'
        'SUM(INDEX(E$23:NL$23,DeliverStart):INDEX(E$23:NL$23,MAX(DeliverStart,StabMonth-1))))*OpReserveCoverage,'
        'OpReserveMinMonths*INDEX(E$37:NL$37,StabMonth))')
    return {'wired': ['CapAdj*', 'InvestmentType', 'Location', 'NetAcres', 'FARRatio',
                      'ModelUnits', 'RetailIncomeMo', 'DevFeeThirdPartySplit', 'LeaseUpIncomeOffset']}


# =====================================================================  FIX 5
def _dead_input_names(wb):
    import re
    inp = {}
    for nm, dn in wb.defined_names.items():
        for title, coord in dn.destinations:
            if title == 'Inputs':
                inp[nm] = coord
    blob = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('='):
                    blob.append(c.value)
    blob = "\n".join(blob)
    dead = [nm for nm in inp
            if not re.search(r'(?<![A-Za-z0-9_])' + re.escape(nm) + r'(?![A-Za-z0-9_])', blob)]
    return sorted(dead)


def fix_5(wb):
    """Diagnostics prove their claims: workbook-wide error scan, non-trivial refi
    checks, economic sanity warnings, dead-input audit."""
    dg = wb['Diagnostics']

    # ---- (1) workbook-wide formula-error scan (per-sheet helper cells) -------
    # Each sheet's calc block is scanned in its own helper cell (column F), and the
    # headline error count sums them.  Per-sheet helpers keep every formula small,
    # auditable, and fast — and pinpoint which sheet holds an error if one appears.
    # 1-D output vectors per sheet (rows / columns) — the computed cells where an
    # error would actually surface — kept 1-D so the recalc engine verifies them too.
    scan_ranges = [
        ("Timeline", ["Timeline!E41:NL41", "Timeline!C17:C25"]),
        ("DevSpend", ["DevSpend!E34:NL34", "DevSpend!E38:NL38", "DevSpend!D34:D43"]),
        ("ConLoan", ["ConLoan!E31:NL31", "ConLoan!E32:NL32", "ConLoan!C35:C48"]),
        ("LeaseUp", ["LeaseUp!E8:NL8", "LeaseUp!E10:NL10", "LeaseUp!C13:C14"]),
        ("Operating", ["Operating!E23:NL23", "Operating!E36:NL36", "Operating!E38:NL38",
                        "Operating!E42:NL42", "Operating!E69:NL69", "Operating!C47:C77"]),
        ("PermLoan", ["PermLoan!E30:NL30", "PermLoan!C4:C33"]),
        ("Disposition", ["Disposition!C4:C25"]),
        ("CashFlow", ["CashFlow!E6:NL6", "CashFlow!E7:NL7", "CashFlow!C10:C15"]),
        ("Returns", ["Returns!E11:NL11", "Returns!E12:NL12", "Returns!E39:NL39", "Returns!C4:C42"]),
        ("Proforma", ["Proforma!G12:G47", "Proforma!C5:C8"]),
        ("Dashboard", ["Dashboard!F5:F36"]),  # F-col numeric; C-col mixes status text
        ("Budget", ["Budget!J5:J39", "Budget!D5:D32"]),
        ("UnitMatrix", ["UnitMatrix!C14:C27", "UnitMatrix!F4:F11"]),
        # Inputs are hard-keyed; scan only the numeric derived cells (skip text/date
        # cells, which the recalc engine cannot ISERROR-scan).
        ("Inputs", ["Inputs!C16:C27", "Inputs!C119:C145"]),
    ]
    S(dg, 'F3', 'Errors by sheet', F_HDR, FILL_HDR, align=RIGHT)
    hrow = 4
    for name, rngs in scan_ranges:
        expr = "+".join(f"SUMPRODUCT(--ISERROR({rng}))" for rng in rngs)
        c = dg.cell(hrow, 6, value=f"={expr}")
        c.font = F_FORM; c.number_format = FMT_INT; c.alignment = RIGHT
        cc = dg.cell(hrow, 7, value=name); cc.font = F_SUB; cc.alignment = LEFT
        hrow += 1
    scan_last = hrow - 1
    dg['B4'].value = 'Total formula errors (workbook-wide scan = 0)'
    dg['C4'].value = f"=SUM(F4:F{scan_last})"
    dg['D4'].value = f'=IF(SUM(F4:F{scan_last})=0,"PASS","FAIL")'

    # ---- (2) non-trivial refi-or-sell + conditional perm-retires-con ---------
    dg['C16'].value = '=RefiFlag'
    dg['D16'].value = ('=IF(RefiFlag=1,IF(AND(PermLoan>0,ABS(Chk_RefiBalance)<1,RefiMonth<=ConMaturity),'
                       '"PASS","FAIL"),IF(AND(PermLoan=0,ActivePayoff<0),"PASS","FAIL"))')
    dg['B17'].value = 'Permanent loan retires construction (refi only)'
    dg['C17'].value = '=IF(RefiFlag=1,PermLoan-ConPayoffBal,0)'
    dg['D17'].value = '=IF(RefiFlag=1,IF(PermLoan>=ConPayoffBal-1,"PASS","FAIL"),"N/A")'

    # ---- (3) rebuild summary tail with sanity warnings & dead-input audit ----
    for rr in range(26, dg.max_row + 1):
        for cc in range(2, 5):
            dg.cell(rr, cc).value = None

    def sect(r, text):
        S(dg, f'B{r}', text, F_SECT, FILL_SECT, align=LEFT)
        for col in 'CD': dg[f'{col}{r}'].fill = FILL_SECT

    def chk(r, label, cval, dform, fmt=FMT_MONEY):
        S(dg, f'B{r}', label, F_LBL, align=LEFT)
        S(dg, f'C{r}', cval, F_FORM, num=fmt, align=RIGHT)
        S(dg, f'D{r}', dform, F_LBLB, align=CENTER)

    r = 27
    sect(r, 'Economic Sanity (warnings — do not count as errors)'); r += 1
    warn_first = r
    chk(r, 'Development spread (untrended YOC − exit cap)', '=DevSpreadBps',
        '=IF(YieldOnCost-ExitCapResidential<0,"WARN","OK")', '#,##0" bps"'); r += 1
    chk(r, 'Construction term within 12–60 months', '=ConLoanTermMonths',
        '=IF(AND(ConLoanTermMonths>=12,ConLoanTermMonths<=60),"OK","WARN")', FMT_INT); r += 1
    chk(r, 'Stabilized tax per unit within $500–$4,000', '=IF(TotalUnits=0,0,StabTaxAnnual/TotalUnits)',
        '=IF(AND(StabTaxAnnual/TotalUnits>=500,StabTaxAnnual/TotalUnits<=4000),"OK","WARN")', FMT_AMT); r += 1
    chk(r, 'Absorption pace ≤ delivery pace', '=UnitsLeasedPerMo',
        '=IF(UnitsLeasedPerMo>UnitsDeliveredPerMo,"WARN","OK")', FMT_INT); r += 1
    chk(r, 'Untrended NOI ≤ trended NOI (growth ≥ 0)', '=StabNOI-StabNOI_Untrended',
        '=IF(OR(GrowthRent<0,GrowthExpense<0),"n/a",IF(StabNOI_Untrended<=StabNOI+1,"OK","WARN"))'); r += 1
    warn_last = r - 1

    r += 1
    sect(r, 'Input Hygiene'); r += 1
    dead = _dead_input_names(wb)
    S(dg, f'B{r}', 'Dead input named ranges (0 formula refs)', F_LBL, align=LEFT)
    S(dg, f'C{r}', len(dead), F_FORM, num=FMT_INT, align=RIGHT)
    S(dg, f'D{r}', 'OK' if not dead else 'WARN', F_LBLB, align=CENTER)
    note = ', '.join(dead) if dead else 'none — every input drives ≥1 formula'
    c = dg.cell(r, 5, value='build-verified: ' + note); c.font = F_SUB; c.alignment = LEFT
    dead_row = r
    r += 2

    sect(r, 'Summary'); r += 1
    err_row = r
    S(dg, f'B{r}', 'TOTAL ERRORS FOUND', F_LBLB, align=LEFT)
    S(dg, f'C{r}', '=COUNTIF(D4:D25,"FAIL")+COUNTIF(Proforma!$D$46:$D$46,"FAIL")',
      F_TOT, num=FMT_INT, align=RIGHT)
    _add_named(wb, 'ErrorsFound', f"'Diagnostics'!$C${err_row}")
    r += 1
    S(dg, f'B{r}', 'TOTAL WARNINGS', F_LBLB, align=LEFT)
    S(dg, f'C{r}', f'=COUNTIF(D{warn_first}:D{warn_last},"WARN")+COUNTIF(D{dead_row}:D{dead_row},"WARN")',
      F_TOT, num=FMT_INT, align=RIGHT)
    _add_named(wb, 'WarningsFound', f"'Diagnostics'!$C${r}")
    r += 1
    S(dg, f'B{r}', 'MODEL STATUS', F_LBLB, align=LEFT)
    S(dg, f'C{r}', f'=IF(C{err_row}=0,"ALL CHECKS PASS","REVIEW FAILURES")', F_TOT, align=RIGHT)
    return {'dead_inputs': dead, 'error_row': err_row}


FIXES = {'fix1': fix_1, 'fix2': fix_2, 'fix3': fix_3, 'fix4': fix_4,
         'fix5': fix_5, 'fix6': fix_6, 'fix7': fix_7}


def main():
    src, dst, which = sys.argv[1], sys.argv[2], sys.argv[3]
    wb = openpyxl.load_workbook(src)
    names = list(FIXES.keys()) if which == 'all' else which.split(',')
    for n in names:
        info = FIXES[n](wb)
        print(f"applied {n}: {info if info else ''}")
    wb.save(dst)
    print(f"saved {dst}")


if __name__ == '__main__':
    main()
