"""
waterfall.py — item 6: add an institutional monthly equity Waterfall tab modeled on the
reference underwriting model.

Structure (non-circular, monthly, no macro/solver):
  * The deal's Levered Cash Flow to Equity (CashFlow!row6) is the equity cash stream;
    negative months are contributions, positive months are distributable cash.
  * GP co-invests a share of every contribution; the LP funds the rest.
  * Distributions cascade through IRR-hurdle tiers:
      Tier 1  return of capital + preferred return  -> pro-rata (GP = co-invest %)
      Tier 2  pref .. Hurdle 2                       -> GP promote 2
      Tier 3  Hurdle 2 .. Hurdle 3                   -> GP promote 3
      Tier 4  residual above Hurdle 3                -> GP promote 4
    Each tier accrues its hurdle on cumulative contributions; a tier is satisfied once
    cumulative distributions reach that accrued threshold.
  * GP fee income (development fee, asset-management fee, disposition fee) is shown as a
    separate GP-with-fees cash flow.
  * LP / GP / deal IRR and MOIC, plus a Waterfall Summary block.
"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

NAVY = "FF1F3864"; INK = "FF000000"; BLUE = "FF0000CC"; GREEN = "FF006100"
FILL_SECT = PatternFill("solid", fgColor=NAVY)
FILL_IN = PatternFill("solid", fgColor="FFFDF3D0")
FILL_HDR = PatternFill("solid", fgColor="FF2E5496")
FILL_TOT = PatternFill("solid", fgColor="FFE2EFDA")
F_TITLE = Font(name="Calibri", size=15, bold=True, color=NAVY)
F_SECT = Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
F_LBL = Font(name="Calibri", size=10, color=INK)
F_IN = Font(name="Calibri", size=10, color=BLUE)
F_TOT = Font(name="Calibri", size=10, bold=True, color=INK)
F_FORM = Font(name="Calibri", size=10, color=INK)
F_HDR = Font(name="Calibri", size=9, bold=True, color="FFFFFFFF")
F_LINK = Font(name="Calibri", size=10, color=GREEN)
F_KPI = Font(name="Calibri", size=12, bold=True, color=NAVY)
RIGHT = Alignment(horizontal="right"); LEFT = Alignment(horizontal="left"); CENTER = Alignment(horizontal="center")
TOPB = Border(top=Side(style="thin", color="FF808080"))
MONEY = '#,##0;(#,##0);"–"'; PCT = '0.00%'; MULT = '0.00"x"'
FIRST, LAST = 5, 376


def build(wb):
    if 'Waterfall' in wb.sheetnames:
        del wb['Waterfall']
    ws = wb.create_sheet('Waterfall')
    # place after Returns
    wb.move_sheet('Waterfall', -(len(wb.sheetnames) - 1 - wb.sheetnames.index('Returns')))
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = '375623'
    ws.column_dimensions['A'].width = 2.5
    ws.column_dimensions['B'].width = 38
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 15
    for m in range(1, LAST - FIRST + 2):
        ws.column_dimensions[get_column_letter(FIRST + m - 1)].width = 12

    def S(coord, v, font=None, fill=None, num=None, align=None):
        c = ws[coord]; c.value = v
        if font: c.font = font
        if fill: c.fill = fill
        if num: c.number_format = num
        if align: c.alignment = align
        return c

    def repoint(name, ref):
        if name in wb.defined_names: del wb.defined_names[name]
        wb.defined_names.add(DefinedName(name, attr_text=ref))

    S('B1', 'Equity Waterfall — LP / GP Promote', F_TITLE)
    S('B2', '="IRR-hurdle waterfall  •  "&PropertyName', Font(name="Calibri", size=10, italic=True, color="FF595959"))

    # ---- inputs (blue) ----
    S('B4', 'Waterfall Assumptions', F_SECT, FILL_SECT, align=LEFT)
    for col in ('C', 'D'): ws[f'{col}4'].fill = FILL_SECT
    inp = [
        ('GP Co-Invest %', 0.10, PCT, 'GPCoInvest', (0, 1)),
        ('Preferred Return (Tier 1)', 0.08, PCT, 'PrefReturn', (0, 0.30)),
        ('Hurdle 2 IRR', 0.12, PCT, 'Hurdle2IRR', (0, 0.50)),
        ('Hurdle 3 IRR', 0.15, PCT, 'Hurdle3IRR', (0, 0.60)),
        ('GP Promote — Tier 2', 0.20, PCT, 'Promote2', (0, 1)),
        ('GP Promote — Tier 3', 0.30, PCT, 'Promote3', (0, 1)),
        ('GP Promote — Tier 4 (residual)', 0.40, PCT, 'Promote4', (0, 1)),
    ]
    r = 5
    for lab, val, fmt, nm, (lo, hi) in inp:
        S(f'B{r}', lab, F_LBL, align=LEFT)
        S(f'C{r}', val, F_IN, FILL_IN, num=fmt, align=RIGHT)
        dv = DataValidation(type='decimal', operator='between', formula1=repr(lo), formula2=repr(hi),
                            allow_blank=False, showErrorMessage=True)
        dv.promptTitle = lab; ws.add_data_validation(dv); dv.add(ws[f'C{r}'])
        repoint(nm, f"'Waterfall'!$C${r}")
        r += 1

    # ---- KPI summary ----
    S('B14', 'Returns by Partner', F_SECT, FILL_SECT, align=LEFT)
    for col in ('C', 'D'): ws[f'{col}14'].fill = FILL_SECT
    S('C15', 'IRR', F_HDR, FILL_HDR, align=RIGHT); S('D15', 'MOIC', F_HDR, FILL_HDR, align=RIGHT)
    # (filled after engine rows known)

    # ====================== monthly engine ======================
    HDR = 20
    S(f'B{HDR}', 'Period →', F_HDR, FILL_HDR, align=LEFT); S(f'D{HDR}', 'Total', F_HDR, FILL_HDR, align=RIGHT)
    S(f'B{HDR+1}', 'Month Ending', F_LBL, align=LEFT)
    for m in range(1, LAST - FIRST + 2):
        col = get_column_letter(FIRST + m - 1)
        S(f'{col}{HDR}', m, F_HDR, FILL_HDR, num='#,##0', align=RIGHT)
        S(f'{col}{HDR+1}', f'=Operating!{col}$4', F_FORM, num='mmm-yy', align=RIGHT)

    PR = HDR  # period-number row = header row
    # row map
    base = HDR + 2
    rows = {}
    order = [
        ('SECT', 'Equity Cash Flow'),
        ('LevCF', 'Levered Cash Flow to Equity'),
        ('Contrib', 'Total Equity Contribution'),
        ('Avail', 'Cash Available to Distribute'),
        ('SECT', 'Hurdle Accruals (cumulative capital at hurdle)'),
        ('Acc1', 'Accrued @ Preferred'),
        ('Acc2', 'Accrued @ Hurdle 2'),
        ('Acc3', 'Accrued @ Hurdle 3'),
        ('CumBoM', 'Cumulative Distributions (BoM)'),
        ('SECT', 'Tier Distributions'),
        ('d1', 'Tier 1 — Return of Capital + Pref'),
        ('d2', 'Tier 2 — Pref to Hurdle 2'),
        ('d3', 'Tier 3 — Hurdle 2 to Hurdle 3'),
        ('d4', 'Tier 4 — Residual'),
        ('CumEoM', 'Cumulative Distributions (EoM)'),
        ('SECT', 'LP / GP Split'),
        ('GPd', 'GP Distributions'),
        ('LPd', 'LP Distributions'),
        ('GPc', 'GP Contributions'),
        ('LPc', 'LP Contributions'),
        ('GPnet', 'GP Net Cash Flow'),
        ('LPnet', 'LP Net Cash Flow'),
        ('SECT', 'GP Fee Income'),
        ('fDev', 'Development Fee'),
        ('fAM', 'Asset Management Fee'),
        ('fDispo', 'Disposition Fee'),
        ('GPfee', 'GP Net Cash Flow incl. Fees'),
    ]
    r = base
    for key, lab in order:
        if key == 'SECT':
            S(f'B{r}', lab, F_SECT, FILL_SECT, align=LEFT)
            S(f'D{r}', '', F_SECT, FILL_SECT)
            for m in range(1, LAST - FIRST + 2):
                ws[f'{get_column_letter(FIRST+m-1)}{r}'].fill = FILL_SECT
            r += 1; continue
        rows[key] = r; r += 1
    # write formulas per month
    def prevcol(m):
        return get_column_letter(FIRST + m - 2)
    for m in range(1, LAST - FIRST + 2):
        c = get_column_letter(FIRST + m - 1); p = prevcol(m); first = (m == 1)
        def f(key, formula):
            cell = ws[f'{c}{rows[key]}']; cell.value = formula
            cell.number_format = MONEY; cell.font = F_FORM; cell.alignment = RIGHT
        f('LevCF', f'=CashFlow!{c}$6')
        f('Contrib', f'=MAX(0,-{c}{rows["LevCF"]})')
        f('Avail', f'=MAX(0,{c}{rows["LevCF"]})')
        pref = '(1+PrefReturn)^(1/12)'; h2 = '(1+Hurdle2IRR)^(1/12)'; h3 = '(1+Hurdle3IRR)^(1/12)'
        f('Acc1', f'={("0" if first else p+str(rows["Acc1"]))}*{pref}+{c}{rows["Contrib"]}')
        f('Acc2', f'={("0" if first else p+str(rows["Acc2"]))}*{h2}+{c}{rows["Contrib"]}')
        f('Acc3', f'={("0" if first else p+str(rows["Acc3"]))}*{h3}+{c}{rows["Contrib"]}')
        f('CumBoM', f'={("0" if first else p+str(rows["CumEoM"]))}')
        av = f'{c}{rows["Avail"]}'; cb = f'{c}{rows["CumBoM"]}'
        d1 = f'{c}{rows["d1"]}'; d2 = f'{c}{rows["d2"]}'; d3 = f'{c}{rows["d3"]}'
        f('d1', f'=MIN({av},MAX(0,{c}{rows["Acc1"]}-{cb}))')
        f('d2', f'=MIN({av}-{d1},MAX(0,{c}{rows["Acc2"]}-{cb}-{d1}))')
        f('d3', f'=MIN({av}-{d1}-{d2},MAX(0,{c}{rows["Acc3"]}-{cb}-{d1}-{d2}))')
        f('d4', f'={av}-{d1}-{d2}-{d3}')
        f('CumEoM', f'={cb}+{av}')
        f('GPd', f'={d1}*GPCoInvest+{d2}*Promote2+{d3}*Promote3+{c}{rows["d4"]}*Promote4')
        f('LPd', f'={av}-{c}{rows["GPd"]}')
        f('GPc', f'={c}{rows["Contrib"]}*GPCoInvest')
        f('LPc', f'={c}{rows["Contrib"]}*(1-GPCoInvest)')
        f('GPnet', f'={c}{rows["GPd"]}-{c}{rows["GPc"]}')
        f('LPnet', f'={c}{rows["LPd"]}-{c}{rows["LPc"]}')
        f('fDev', f'=DevSpend!{c}$35')
        f('fAM', f'=AssetMgmtFeePct*Operating!{c}$23')
        f('fDispo', f'=IF({c}${PR}=DispoMonth,SponsorDispoFeePct*ExitValue,0)')
        f('GPfee', f'={c}{rows["GPnet"]}+{c}{rows["fDev"]}+{c}{rows["fAM"]}+{c}{rows["fDispo"]}')
    # labels + totals (col D)
    for key, lab in order:
        if key == 'SECT': continue
        rr = rows[key]
        ws[f'B{rr}'].value = lab
        ws[f'B{rr}'].font = F_TOT if key in ('GPnet', 'LPnet', 'GPfee', 'CumEoM') else F_LBL
        ws[f'B{rr}'].alignment = LEFT
        if key in ('CumBoM', 'CumEoM', 'Acc1', 'Acc2', 'Acc3'):
            S(f'D{rr}', f'=INDEX(E{rr}:NL{rr},HoldPeriodMonths)', F_FORM, num=MONEY, align=RIGHT)
        else:
            S(f'D{rr}', f'=SUM(E{rr}:NL{rr})', (F_TOT if key in ('GPnet', 'LPnet', 'GPfee') else F_FORM), num=MONEY, align=RIGHT)

    # ---- partner IRR/MOIC (named) ----
    def irr(row):
        return f'=IFERROR((1+IRR(E{row}:NL{row},0.01))^12-1,"n/a")'
    repoint_cells = {
        'LP_IRR': (16, 'LP', irr(rows['LPnet']), f'=IF(SUMIF(E{rows["LPc"]}:NL{rows["LPc"]},">0")=0,0,SUM(E{rows["LPd"]}:NL{rows["LPd"]})/SUM(E{rows["LPc"]}:NL{rows["LPc"]}))'),
        'GP_IRR': (17, 'GP (promote)', irr(rows['GPnet']), f'=IF(SUM(E{rows["GPc"]}:NL{rows["GPc"]})=0,0,SUM(E{rows["GPd"]}:NL{rows["GPd"]})/SUM(E{rows["GPc"]}:NL{rows["GPc"]}))'),
        'GPfee_IRR': (18, 'GP incl. fees', irr(rows['GPfee']), f'=IF(SUM(E{rows["GPc"]}:NL{rows["GPc"]})=0,0,(SUM(E{rows["GPd"]}:NL{rows["GPd"]})+SUM(E{rows["fDev"]}:NL{rows["fDev"]})+SUM(E{rows["fAM"]}:NL{rows["fAM"]})+SUM(E{rows["fDispo"]}:NL{rows["fDispo"]}))/SUM(E{rows["GPc"]}:NL{rows["GPc"]}))'),
    }
    for nm, (krow, lab, irrf, moicf) in repoint_cells.items():
        S(f'B{krow}', lab, F_LBL, align=LEFT)
        S(f'C{krow}', irrf, F_KPI, num=PCT, align=RIGHT)
        S(f'D{krow}', moicf, F_KPI, num=MULT, align=RIGHT)
        repoint(nm, f"'Waterfall'!$C${krow}")

    # ---- reconciliation: GP + LP distributions == total available ----
    S(f'B{r+1}', 'Check: GP + LP dist − total distributions', F_LBL, align=LEFT)
    S(f'C{r+1}', f'=SUM(E{rows["GPd"]}:NL{rows["GPd"]})+SUM(E{rows["LPd"]}:NL{rows["LPd"]})-SUM(E{rows["Avail"]}:NL{rows["Avail"]})',
      F_FORM, num=MONEY, align=RIGHT)
    S(f'D{r+1}', f'=IF(ABS(C{r+1})<1,"PASS","FAIL")', F_TOT, align=CENTER)
    repoint('WaterfallCheck', f"'Waterfall'!$C${r+1}")
    S(f'B{r+2}', 'Check: LP + GP contributions − total equity', F_LBL, align=LEFT)
    S(f'C{r+2}', f'=SUM(E{rows["GPc"]}:NL{rows["GPc"]})+SUM(E{rows["LPc"]}:NL{rows["LPc"]})-SUM(E{rows["Contrib"]}:NL{rows["Contrib"]})',
      F_FORM, num=MONEY, align=RIGHT)
    S(f'D{r+2}', f'=IF(ABS(C{r+2})<1,"PASS","FAIL")', F_TOT, align=CENTER)

    # year grouping + freeze
    ws.sheet_properties.outlinePr.summaryRight = True
    for y in range(1, 32):
        fc = FIRST + (y - 1) * 12
        g0 = get_column_letter(fc + 1); g1 = get_column_letter(min(fc + 11, LAST))
        if fc + 1 <= LAST:
            ws.column_dimensions.group(g0, g1, outline_level=1, hidden=False)
    ws.freeze_panes = f'E{HDR+2}'
    return {'engine_rows': len(rows)}


def main():
    src, dst = sys.argv[1], sys.argv[2]
    wb = openpyxl.load_workbook(src)
    print('applied:', build(wb)); wb.save(dst); print('saved', dst)


if __name__ == '__main__':
    main()
