import sys, warnings, numpy as np, formulas
warnings.filterwarnings("ignore")
src=sys.argv[1]; coords=sys.argv[2].split(",")
xl=formulas.ExcelModel().loads(src).finish(); sol=xl.calculate()
kmap={}
for k in sol.keys():
    if "!" not in k: continue
    left,cell=k.rsplit("!",1); cell=cell.replace("$","").upper().strip("'")
    sheet=left.split("]")[-1].strip("'").upper(); kmap[(sheet,cell)]=k
def val(sh,cl):
    k=kmap.get((sh.upper(),cl.upper()))
    if k is None: return "<nokey>"
    v=sol.get(k)
    try:
        vv=v.value
        if isinstance(vv,np.ndarray): vv=vv.ravel()[0] if vv.size else None
        return vv
    except: return v
for co in coords:
    sh,cl=co.split("!"); print(f"  {co} = {val(sh,cl)}")
