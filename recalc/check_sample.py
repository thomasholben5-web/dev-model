"""check_sample.py — recalc + dump every Diagnostics check, reconciliations, key outputs, errors."""
import sys, warnings, numpy as np, openpyxl, formulas
warnings.filterwarnings("ignore")
ERRS=("#REF!","#DIV/0!","#VALUE!","#N/A","#NAME?","#NUM!","#NULL!")
def scalar(v):
    try: val=v.value
    except AttributeError: val=v
    if isinstance(val,np.ndarray):
        f=val.ravel(); val=f[0] if f.size else None
    return val
def main():
    src=sys.argv[1]
    xl=formulas.ExcelModel().loads(src).finish(); sol=xl.calculate()
    kmap={}
    for k in sol.keys():
        if "!" not in k: continue
        left,cell=k.rsplit("!",1); cell=cell.replace("$","").upper().strip("'")
        sheet=left.split("]")[-1].strip("'").upper(); kmap[(sheet,cell)]=k
    def get(sheet,cell):
        key=kmap.get((sheet.upper(),cell.upper()))
        return scalar(sol.get(key)) if key else "<none>"
    # error scan
    hits=[(k,scalar(v)) for k,v in sol.items() if "!" in k and isinstance(scalar(v),str) and any(e in scalar(v) for e in ERRS)]
    print(f"CELLS {len(sol)}  ERRORS {len(hits)}")
    for k,v in hits[:30]: print("  ERR",k,v)
    print("\n== Diagnostics checks (label | value | status) ==")
    dg_rows=list(range(4,26))+list(range(43,48))
    wb=openpyxl.load_workbook(src)
    dg=wb['Diagnostics']
    fails=0
    for r in dg_rows:
        lbl=dg[f'B{r}'].value
        if not lbl: continue
        val=get('Diagnostics',f'C{r}'); stat=get('Diagnostics',f'D{r}')
        flag='' if stat in ('PASS','N/A','OK') else '  <<< FAIL'
        if flag: fails+=1
        vs=f'{val:,.4f}' if isinstance(val,(int,float)) else str(val)
        print(f'  r{r:<2} {str(lbl)[:48]:48} {vs:>16}  {stat}{flag}')
    print(f"\nFAILS: {fails}")
    print("\n== warnings ==")
    for r in range(28,33):
        lbl=dg[f'B{r}'].value
        if lbl: print(f'  {str(lbl)[:46]:46} {get("Diagnostics",f"C{r}")}  {get("Diagnostics",f"D{r}")}')
    print("\n== key outputs ==")
    for nm,(sh,cl) in {'TotalUnits':('UnitMatrix','C14'),'TDC_Total':('ConLoan','C39'),'LevIRR':('Returns','C4'),
                       'UnlevIRR':('Returns','C7') if False else ('Returns','C7'),'StaticValue':('Inputs','S11'),
                       'DevFee':('Budget','J113'),'ExitValue':('Disposition','C10'),'LTC_Realized':('ConLoan','C18'),
                       'StabMonth':('Timeline','C20'),'DispoMonth':('Timeline','C21'),'ConStart':('Timeline','C17')}.items():
        print(f'  {nm:14} = {get(sh,cl)}')
if __name__=='__main__': main()
