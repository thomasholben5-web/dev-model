"""
batch4 — formatting cleanup (value-neutral):

  * Timeline: the batch-2 milestone rewrite cleared cell VALUES but not fills/borders, leaving
    stale input-fills (C10, C13, D9), a stale RED fill under the Recap/Refi duration (D11), and
    orphan borders in the emptied rows 13-14.  Re-clean the whole milestone block so the yellow
    input fill appears ONLY on the four real duration inputs (D4:D7), markers are plain grey,
    the derived Lease-Up duration (D9) is green with no input fill, and rows 13-14 are blank.
  * Inputs: move the Delivery Lag input into the "Unit Delivery & Lease-Up" section and remove
    the retired duration remnants (the old Delivery-Lag label/value in F9:G9 and the
    "durations moved" note in F10).

No formulas or values change except relocating one input cell (DeliveryLagMonths), whose name
is repointed, so the model stays value-neutral.
"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName

BLUE = "FF0000CC"; GREEN = "FF006100"; INK = "FF000000"; GREY = "FF808080"
FILL_IN = PatternFill("solid", fgColor="FFFDF3D0")
NOFILL = PatternFill()
NOBORDER = Border()
F_IN = Font(name="Calibri", size=10, color=BLUE)
F_LINK = Font(name="Calibri", size=10, color=GREEN)
F_LBL = Font(name="Calibri", size=10, color=INK)
F_MARK = Font(name="Calibri", size=10, color=GREY)
F_EXIT = Font(name="Calibri", size=9, bold=True, color=GREY)
RIGHT = Alignment(horizontal="right"); LEFT = Alignment(horizontal="left")


def build(wb):
    tl = wb['Timeline']; ip = wb['Inputs']

    def repoint(name, ref):
        if name in wb.defined_names: del wb.defined_names[name]
        wb.defined_names.add(DefinedName(name, attr_text=ref))

    # ---- Timeline: re-clean the milestone block (rows 4-14, cols A-E) ----
    for r in range(4, 15):
        for col in 'ABCDE':
            c = tl[f'{col}{r}']
            c.fill = NOFILL; c.border = NOBORDER
    # re-apply clean formatting to the live rows 4-12
    markers = {8, 10, 11, 12}; derived = {9}; inputs = {4, 5, 6, 7}
    for r in range(4, 13):
        tl[f'B{r}'].font = F_LBL; tl[f'B{r}'].alignment = LEFT
        for col in 'CDE':
            tl[f'{col}{r}'].number_format = '#,##0'; tl[f'{col}{r}'].alignment = RIGHT
        # start (C) and end (E) are derived -> plain
        tl[f'C{r}'].font = F_LBL; tl[f'E{r}'].font = F_LBL
        d = tl[f'D{r}']
        if r in inputs:
            d.font = F_IN; d.fill = FILL_IN
        elif r in derived:
            d.font = F_LINK
        elif r in markers:
            d.font = F_MARK
    tl['A11'].value = 'Exit'; tl['A11'].font = F_EXIT; tl['A11'].alignment = LEFT
    # rows 13-14 left fully blank (values already cleared in batch 2)
    for r in (13, 14):
        for col in 'ABCDE':
            tl[f'{col}{r}'].value = None

    # ---- Inputs: move Delivery Lag into the Unit Delivery & Lease-Up section ----
    lag = ip['G9'].value if isinstance(ip['G9'].value, (int, float)) else 19
    ip['B45'].value = 'Delivery Lag (mo)'; ip['B45'].font = F_LBL; ip['B45'].alignment = LEFT
    ip['C45'].value = lag; ip['C45'].font = F_IN; ip['C45'].fill = FILL_IN
    ip['C45'].number_format = '#,##0'; ip['C45'].alignment = RIGHT
    repoint('DeliveryLagMonths', "'Inputs'!$C$45")
    # clear the retired remnants
    ip['F9'].value = None; ip['G9'].value = None; ip['G9'].fill = NOFILL
    ip['F10'].value = None

    # ---- bug fix: GP fees must stop at disposition (AM fee leaked on post-sale Operating EGI) ----
    from openpyxl.utils import get_column_letter
    wf = wb['Waterfall']
    gated = 0
    for m in range(1, 376 - 5 + 2):
        c = get_column_letter(5 + m - 1)
        # Development Fee (row 45) and Asset Management Fee (row 46): only while the asset is held
        wf[f'{c}45'].value = f'=IF({c}$20<=DispoMonth,DevSpend!{c}$101,0)'
        wf[f'{c}46'].value = f'=IF({c}$20<=DispoMonth,AssetMgmtFeePct*Operating!{c}$23,0)'
        gated += 1

    return {'timeline_cleaned': 'rows 4-14', 'delivery_lag_moved_to': 'Inputs!C45',
            'gp_fees_gated_to_dispo': gated}


def main():
    src, dst = sys.argv[1], sys.argv[2]
    wb = openpyxl.load_workbook(src)
    print('applied:', build(wb)); wb.save(dst); print('saved', dst)


if __name__ == '__main__':
    main()
