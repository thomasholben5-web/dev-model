"""
batch2 — Timeline overhaul + related review fixes (all in one validated change):

  #4  Consolidate milestones to 6 user-facing phases (Open Escrow -> Due Diligence,
      Initial Closing -> Pre-Construction, Recap+Disposition grouped as Exit).
      1st Unit Delivery and Stabilization become point markers (their durations drove
      nothing — delivery timing is DeliveryLagMonths, stabilization end was cosmetic).
      Durations are now entered DIRECTLY on the Timeline (Duration column = blue inputs);
      the Inputs-tab duration section is retired and the Dur_* names repoint to the Timeline.
  Annual: "Development Spend" now captures all cash uses ex-interest (ConLoan fundable uses)
      so Development Cost + Construction Interest = TDC, with a check cell (AnnualTDCCheck).
  Growth: remove the off-by-one so growth begins IN the chosen start month (not the month
      after) — GrowthStartMonth drops its trailing +1.
  UX: remove the automatic data-validation input-message pop-ups on every input (keep the
      dropdowns and the invalid-entry alerts).

Value impact: the milestone/duration moves are neutral (kept phases keep identical start
formulas; merged phases share start months; the two retired durations fed only unused ends).
The growth off-by-one fix intentionally shifts growth onset one month earlier.
"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string as ci
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

BLUE = "FF0000CC"; GREEN = "FF006100"; INK = "FF000000"; GREY = "FF808080"
FILL_IN = PatternFill("solid", fgColor="FFFDF3D0")
F_IN = Font(name="Calibri", size=10, color=BLUE)
F_LINK = Font(name="Calibri", size=10, color=GREEN)
F_LBL = Font(name="Calibri", size=10, color=INK)
F_MARK = Font(name="Calibri", size=10, color=GREY)
F_HINT = Font(name="Calibri", size=9, italic=True, color=GREY)
F_EXIT = Font(name="Calibri", size=9, bold=True, color=GREY)
LEFT = Alignment(horizontal="left"); RIGHT = Alignment(horizontal="right")

# (row, name, start_formula, duration, kind)   kind: 'input' (blue) / 'derived' (green) / 'marker' (grey)
LEASEUP_DUR = '=MAX(CEILING(StabLeasedTarget/UnitsLeasedPerMo,1),CEILING(StabLeasedTarget/UnitsDeliveredPerMo,1))'
MILESTONES = [
    (4,  'Due Diligence',              1,                                             4,           'input'),
    (5,  'Entitlement',                '=+MS_Due_Diligence_End',                       5,           'input'),
    (6,  'Pre-Construction',           '=+MS_Entitlement_End',                         3,           'input'),
    (7,  'Construction',               '=+MS_Pre_Construction_End',                    22,          'input'),
    (8,  '1st Unit Delivery',          '=+MS_Construction_Start+DeliveryLagMonths',    1,           'marker'),
    (9,  'Lease-Up',                   '=+MS_1st_Unit_Delivery_Start',                 LEASEUP_DUR, 'derived'),
    (10, 'Stabilization',              '=+MS_Lease_Up_End',                            1,           'marker'),
    (11, 'Recapitalization/Refinance', '=IF(RefiFlag=1,RefiMonth,MS_Stabilization_End)', 1,         'marker'),
    (12, 'Disposition',                '=HoldPeriodMonths',                            1,           'marker'),
]
# MS_<base> -> new row; Start col C, End col E. Merged bases point at their host row.
MS_ROW = {
    'MS_Due_Diligence': 4, 'MS_Open_Escrow': 4,
    'MS_Entitlement': 5,
    'MS_Pre_Construction': 6, 'MS_Initial_Closing': 6,
    'MS_Construction': 7,
    'MS_1st_Unit_Delivery': 8,
    'MS_Lease_Up': 9,
    'MS_Stabilization': 10,
    'MS_Recapitalization_Refinance': 11,
    'MS_Disposition': 12,
}
# Dur_<name> -> Timeline duration cell it now reads (the 4 real inputs + markers for the rest)
DUR_ROW = {
    'Dur_OpenEscrow': 4, 'Dur_DueDiligence': 4, 'Dur_Entitlement': 5, 'Dur_PreConstruction': 6,
    'Dur_InitialClosing': 6, 'Dur_Construction': 7, 'Dur_FirstUnitDelivery': 8,
    'Dur_Stabilization': 10, 'Dur_Recap': 11, 'Dur_Disposition': 12,
}


def build(wb):
    tl = wb['Timeline']; bg = wb['Budget']; ip = wb['Inputs']; an = wb['Annual']

    def repoint(name, ref):
        if name in wb.defined_names: del wb.defined_names[name]
        wb.defined_names.add(DefinedName(name, attr_text=ref))

    # ---- #4 consolidated milestone block with durations entered on the Timeline ----
    for r in range(4, 15):
        for col in ('A', 'B', 'C', 'D', 'E'):
            tl[f'{col}{r}'].value = None
    for r, name, start, dur, kind in MILESTONES:
        tl[f'B{r}'].value = name; tl[f'B{r}'].font = F_LBL; tl[f'B{r}'].alignment = LEFT
        tl[f'C{r}'].value = start; tl[f'C{r}'].number_format = '#,##0'; tl[f'C{r}'].alignment = RIGHT
        dcell = tl[f'D{r}']; dcell.value = dur; dcell.number_format = '#,##0'; dcell.alignment = RIGHT
        if kind == 'input':
            dcell.font = F_IN; dcell.fill = FILL_IN
        elif kind == 'derived':
            dcell.font = F_LINK
        else:
            dcell.font = F_MARK
        tl[f'E{r}'].value = f'=SUM(C{r}:D{r})-1'; tl[f'E{r}'].number_format = '#,##0'; tl[f'E{r}'].alignment = RIGHT
    tl['A11'].value = 'Exit'; tl['A11'].font = F_EXIT
    tl['A4'].value = None

    # repoint MS_ names + Milestone ranges + Dur_ names
    for base, row in MS_ROW.items():
        repoint(f'{base}_Start', f"Timeline!$C${row}")
        repoint(f'{base}_End',   f"Timeline!$E${row}")
    repoint('MilestoneNames', "Timeline!$B$4:$B$12")
    repoint('MilestoneStart', "Timeline!$C$4:$C$12")
    repoint('MilestoneEnd',   "Timeline!$E$4:$E$12")
    for name, row in DUR_ROW.items():
        repoint(name, f"Timeline!$D${row}")

    # retire the Inputs duration section (durations now live on the Timeline)
    for r in range(10, 20):
        ip[f'G{r}'].value = None; ip[f'G{r}'].fill = PatternFill()
    ip['F10'].value = '→ Durations now entered on the Timeline tab'; ip['F10'].font = F_HINT
    for r in range(11, 20):
        ip[f'F{r}'].value = None

    # ---- #4 (budget) remap the Initial Closing / Open Escrow phased lines ----
    remapped = 0
    for r in range(5, 98):
        if bg[f'G{r}'].value == 'Initial Closing':
            bg[f'G{r}'].value = 'Pre-Construction'; remapped += 1
        elif bg[f'G{r}'].value == 'Open Escrow':
            bg[f'G{r}'].value = 'Due Diligence'; remapped += 1
    for dv in list(bg.data_validations.dataValidation):
        if 'Initial Closing' in (dv.formula1 or '') or 'Open Escrow' in (dv.formula1 or ''):
            bg.data_validations.dataValidation.remove(dv)
    dvg = DataValidation(type='list',
                         formula1='"Due Diligence, Entitlement, Pre-Construction, Construction, Lease-Up, Stabilization"',
                         allow_blank=True, showErrorMessage=True, showInputMessage=False)
    bg.add_data_validation(dvg); dvg.add('G5:G97')

    # ---- Annual reconciliation: Dev Cost (ex-interest) + Construction Interest = TDC ----
    an['B23'].value = 'Development Cost (cash, ex-interest)'
    for col in range(3, 34):
        L = get_column_letter(col); v = an[f'{L}23'].value
        if isinstance(v, str):
            an[f'{L}23'].value = v.replace('DevSpend!$E$100:$NL$100', 'ConLoan!$E$26:$NL$26')
    an['B28'].value = 'Construction Interest (capitalized, financed)'
    an['B37'].value = 'Check: Σ Dev Cost + Σ Cap. Interest − TDC'
    an['C37'].value = '=SUM(C23:AG23)+SUM(C28:AG28)-TDC_Total'
    an['C37'].number_format = '#,##0;(#,##0);"–"'
    an['D37'].value = '=IF(ABS(C37)<1,"PASS","FAIL")'
    repoint('AnnualTDCCheck', "Annual!$C$37")

    # ---- growth off-by-one: growth begins IN the chosen start month ----
    tl['C22'].value = ('=(YEAR(GrowthStartDate)-YEAR(InceptionDate))*12'
                       '+MONTH(GrowthStartDate)-MONTH(InceptionDate)')

    # ---- remove automatic input-message pop-ups on every data validation ----
    popups = 0
    for ws in wb.worksheets:
        for dv in ws.data_validations.dataValidation:
            if dv.showInputMessage:
                popups += 1
            dv.showInputMessage = False
            dv.promptTitle = None; dv.prompt = None

    return {'milestone_rows': len(MILESTONES), 'budget_phase_remapped': remapped,
            'input_popups_removed': popups}


def main():
    src, dst = sys.argv[1], sys.argv[2]
    wb = openpyxl.load_workbook(src)
    print('applied:', build(wb)); wb.save(dst); print('saved', dst)


if __name__ == '__main__':
    main()
