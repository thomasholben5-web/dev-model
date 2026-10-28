"""
fix_monthlycf.py — reconcile the Monthly CF tab's uses section to Total Development Cost.

The original "Development & Uses" section only summed development *spend* (direct costs,
dev fee, construction-period tax), so its "Total Project Uses" came in ~$863K light of the
Inputs-tab TDC.  TDC also carries the financing costs and reserves that are funded through
the capital stack but are not monthly development draws:

    TDC = ConCostExInt + InterestReserveFunded
        = DirectCosts + DevFee + ConTaxTotal
          + ConOtherCosts + ConPoints + RecourseFee + OpReserve   (funded at initial closing)
          + capitalized construction-loan interest                 (accrues monthly)

The model already lays these out monthly:
    ConLoan!row26  "Fundable Uses"  = dev spend + closing-month (other+points+recourse+opres)
    ConLoan!row31  capitalized interest (monthly accrual), sums to InterestReserveFunded
so  -(ConLoan!row26 + ConLoan!row31)  is the true monthly TDC and foots exactly to TDC_Total.

This rebuilds only the Monthly CF tab: the uses section is expanded to break out each item
in the month it is funded, the subtotal becomes "Total Development Cost (TDC)", and a named
reconciliation cell (MonthlyCFCheck) proves Inputs TDC − Monthly uses = 0.  Nothing else on
the tab changes; no other sheet references it.
"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

NAVY = "FF1F3864"; INK = "FF000000"; GREEN = "FF006100"
FILL_SECT = PatternFill("solid", fgColor=NAVY)
FILL_HDR = PatternFill("solid", fgColor="FF2E5496")
FILL_CHK = PatternFill("solid", fgColor="FFE2EFDA")
F_TITLE = Font(name="Calibri", size=15, bold=True, color=NAVY)
F_SECT = Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
F_LBL = Font(name="Calibri", size=10, color=INK)
F_TOT = Font(name="Calibri", size=10, bold=True, color=INK)
F_FORM = Font(name="Calibri", size=10, color=INK)
F_HDR = Font(name="Calibri", size=9, bold=True, color="FFFFFFFF")
RIGHT = Alignment(horizontal="right"); LEFT = Alignment(horizontal="left"); CENTER = Alignment(horizontal="center")
TOPB = Border(top=Side(style="thin", color="FF808080"))
MONEY = '#,##0;(#,##0);"–"'
FIRST, LAST = 5, 376


def build(wb):
    if 'Monthly CF' in wb.sheetnames:
        del wb['Monthly CF']
    ws = wb.create_sheet('Monthly CF')
    wb.move_sheet('Monthly CF', -(len(wb.sheetnames) - wb.sheetnames.index('CashFlow')))
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = '7F7F7F'
    ws.column_dimensions['A'].width = 2.5
    ws.column_dimensions['B'].width = 36
    ws.column_dimensions['C'].width = 2
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

    S('B1', 'Monthly Cash Flow — Consolidated', F_TITLE)
    S('B2', '="All flows by month  •  "&PropertyName&"  •  hold "&HoldPeriodMonths&" months"',
      Font(name="Calibri", size=10, italic=True, color="FF595959"))
    S('B3', 'Period →', F_HDR, FILL_HDR, align=LEFT); S('D3', 'Total', F_HDR, FILL_HDR, align=RIGHT)
    S('B4', 'Month Ending', F_LBL, align=LEFT)
    for m in range(1, LAST - FIRST + 2):
        col = get_column_letter(FIRST + m - 1)
        S(f'{col}3', m, F_HDR, FILL_HDR, num='#,##0', align=RIGHT)
        S(f'{col}4', f'=Operating!{col}$4', F_FORM, num='mmm-yy', align=RIGHT)

    # (kind, label, per-month template with {c}=col)   kinds: SECT/row/tot/bal/cum/chk
    CLOSE = 'IF({c}$3=MS_Initial_Closing_Start,{item},0)'
    ROWS = [
        ('SECT', 'Development & Uses', None),
        ('row', 'Development Spend', '=-DevSpend!{c}$34'),
        ('row', 'Sponsor Development Fee', '=-DevSpend!{c}$35'),
        ('row', 'Construction-Period Property Tax', '=-DevSpend!{c}$38'),
        ('row', 'Other Loan / Closing Costs', '=-' + CLOSE.format(c='{c}', item='ConOtherCosts')),
        ('row', 'Origination Points', '=-' + CLOSE.format(c='{c}', item='ConPoints')),
        ('row', 'Recourse / Indemnity Fee', '=-' + CLOSE.format(c='{c}', item='RecourseFee')),
        ('row', 'Operating Reserve (funded)', '=-' + CLOSE.format(c='{c}', item='OpReserve')),
        ('row', 'Capitalized Construction Interest', '=-ConLoan!{c}$31'),
        ('tot', 'Total Development Cost (TDC)', '=-(ConLoan!{c}$26+ConLoan!{c}$31)'),
        ('chk', 'Reconciliation: Inputs TDC − Monthly Uses (=0)', None),
        ('SECT', 'Debt', None),
        ('row', 'Construction Debt Draw', '=ConLoan!{c}$29'),
        ('row', 'Construction Interest — Capitalized', '=-ConLoan!{c}$31'),
        ('row', 'Construction Interest — Current Pay', '=-ConLoan!{c}$50'),
        ('row', 'Permanent Interest', '=-PermLoan!{c}$27'),
        ('row', 'Permanent Principal', '=-PermLoan!{c}$28'),
        ('tot', 'Total Cash Debt Service', '=-(ConLoan!{c}$50+PermLoan!{c}$29)'),
        ('bal', 'Construction Loan Balance (EOM)', '=ConLoan!{c}$32'),
        ('bal', 'Permanent Loan Balance (EOM)', '=PermLoan!{c}$30'),
        ('bal', 'Interest Reserve Balance (EOM)', '=ConLoan!{c}$51'),
        ('SECT', 'Equity', None),
        ('row', 'Equity Contribution', '=ConLoan!{c}$27'),
        ('row', 'Refi Cash Distribution', '=IF({c}$3=RefiMonth,RefiCashOut,0)'),
        ('SECT', 'Operations', None),
        ('row', 'Effective Gross Income', '=Operating!{c}$23'),
        ('row', 'Total Operating Expenses', '=-Operating!{c}$37'),
        ('tot', 'Net Operating Income', '=Operating!{c}$38'),
        ('row', 'Asset Management Fees', '=-AssetMgmtFeePct*Operating!{c}$23'),
        ('row', 'Capital Reserves & Tax Prep', '=-(Operating!{c}$40-AssetMgmtFeePct*Operating!{c}$23)'),
        ('tot', 'Property Operating Cash Flow', '=Operating!{c}$42'),
        ('SECT', 'Project Cash Flow', None),
        ('row', 'Net Sale Proceeds (at disposition)', '=IF({c}$3=DispoMonth,NetSale,0)'),
        ('tot', 'Unlevered Cash Flow', '=CashFlow!{c}$7'),
        ('tot', 'Levered Cash Flow to Equity', '=CashFlow!{c}$6'),
        ('cum', 'Cumulative Levered Cash Flow', None),
        ('cum', 'Cumulative Equity Invested', None),
    ]
    r = 5
    lev_row = None; tdc_tot_row = None
    for kind, label, tmpl in ROWS:
        if kind == 'SECT':
            S(f'B{r}', label, F_SECT, FILL_SECT, align=LEFT)
            S(f'D{r}', '', F_SECT, FILL_SECT)
            for m in range(1, LAST - FIRST + 2):
                ws[f'{get_column_letter(FIRST+m-1)}{r}'].fill = FILL_SECT
            r += 1; continue
        if kind == 'chk':
            # single reconciliation cell: Inputs TDC + (negative) monthly uses total == 0
            S(f'B{r}', label, F_LBL, align=LEFT)
            S(f'D{r}', f'=TDC_Total+D{tdc_tot_row}', F_TOT, FILL_CHK, MONEY, RIGHT)
            S(f'B{r}', label + '  ', F_LBL, align=LEFT)
            # PASS/FAIL marker just left of the Total column
            if 'MonthlyCFCheck' in wb.defined_names: del wb.defined_names['MonthlyCFCheck']
            wb.defined_names.add(DefinedName('MonthlyCFCheck', attr_text=f"'Monthly CF'!$D${r}"))
            r += 1; continue
        bold = kind in ('tot',)
        S(f'B{r}', label, F_TOT if bold else F_LBL, align=LEFT)
        if kind == 'cum':
            for m in range(1, LAST - FIRST + 2):
                col = get_column_letter(FIRST + m - 1); prev = get_column_letter(FIRST + m - 2)
                if 'Invested' in label:
                    f = f'={("0" if m==1 else prev+str(r))}+MIN(0,{col}{lev_row})'
                else:
                    f = f'={("0" if m==1 else prev+str(r))}+{col}{lev_row}'
                S(f'{col}{r}', f, F_FORM, num=MONEY, align=RIGHT)
            S(f'D{r}', f'=INDEX(E{r}:NL{r},HoldPeriodMonths)', F_TOT, num=MONEY, align=RIGHT)
        else:
            if label == 'Levered Cash Flow to Equity':
                lev_row = r
            if label.startswith('Total Development Cost'):
                tdc_tot_row = r
            for m in range(1, LAST - FIRST + 2):
                col = get_column_letter(FIRST + m - 1)
                S(f'{col}{r}', tmpl.format(c=col), F_TOT if bold else F_FORM, num=MONEY, align=RIGHT)
            if kind == 'bal':
                S(f'D{r}', f'=INDEX(E{r}:NL{r},HoldPeriodMonths)', F_FORM, num=MONEY, align=RIGHT)
            else:
                S(f'D{r}', f'=SUM(E{r}:NL{r})', F_TOT if bold else F_FORM, num=MONEY, align=RIGHT)
        if bold:
            for cc in ('B', 'D'): ws[f'{cc}{r}'].border = TOPB
        r += 1

    ws.sheet_properties.outlinePr.summaryRight = True
    for y in range(1, 32):
        fc = FIRST + (y - 1) * 12
        g0 = get_column_letter(fc + 1); g1 = get_column_letter(min(fc + 11, LAST))
        if fc + 1 <= LAST:
            ws.column_dimensions.group(g0, g1, outline_level=1, hidden=False)
    ws.freeze_panes = 'E5'
    return {'rows': r - 5, 'tdc_total_row': tdc_tot_row}


def main():
    src, dst = sys.argv[1], sys.argv[2]
    wb = openpyxl.load_workbook(src)
    print('applied:', build(wb)); wb.save(dst); print('saved', dst)


if __name__ == '__main__':
    main()
