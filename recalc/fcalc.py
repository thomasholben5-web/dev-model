"""Full-workbook recalc with `formulas`: error scan + named-range dump (robust keys)."""
import sys, os, json, warnings, numpy as np, openpyxl, formulas
warnings.filterwarnings("ignore")
ERRS=("#REF!","#DIV/0!","#VALUE!","#N/A","#NAME?","#NUM!","#NULL!")
def scalar(v):
    try: val=v.value
    except AttributeError: val=v
    if isinstance(val,np.ndarray):
        f=val.ravel(); val=f[0] if f.size else None
    return val
def main():
    src=sys.argv[1]; names=sys.argv[2].split(",") if len(sys.argv)>2 and sys.argv[2] else []
    xl=formulas.ExcelModel().loads(src).finish(); sol=xl.calculate()
    # normalized key map: strip to SHEET!CELL (upper, no $, no book)
    kmap={}
    for k in sol.keys():
        if "!" not in k: continue
        left,cell=k.rsplit("!",1)
        cell=cell.replace("$","").upper().strip("'")
        sheet=left.split("]")[-1].strip("'").upper()
        kmap[(sheet,cell)]=k
    hits=[]
    for k,v in sol.items():
        if "!" not in k: continue
        val=scalar(v)
        if isinstance(val,str) and any(e in val for e in ERRS): hits.append((k,val))
    print(f"CELLS {len(sol)}  ERRORS {len(hits)}")
    for k,v in hits[:60]: print("  ERR",k,v)
    wb=openpyxl.load_workbook(src); out={}
    for nm in names:
        dn=wb.defined_names.get(nm)
        if not dn: out[nm]="<missing>"; continue
        val=None
        for title,coord in dn.destinations:
            key=kmap.get((title.upper(),coord.replace("$","").upper()))
            if key is not None: val=scalar(sol.get(key))
            break
        out[nm]=None if val is None else (val if isinstance(val,str) else float(val))
        print(f"  {nm} = {out[nm]}")
    json.dump({"cells":len(sol),"errors":len(hits),"named":out,"err_cells":[list(h) for h in hits[:60]]},
              open(src+".calc.json","w"))
if __name__=="__main__": main()
