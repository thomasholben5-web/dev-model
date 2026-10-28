import sys, openpyxl
out=sys.argv[1]; src=sys.argv[2]
ERRS={'#REF!','#DIV/0!','#VALUE!','#N/A','#NAME?','#NUM!','#NULL!'}
wb=openpyxl.load_workbook(out, data_only=True)
bad=[f"{ws.title}!{c.coordinate}={c.value}" for ws in wb.worksheets for row in ws.iter_rows()
     for c in row if isinstance(c.value,str) and any(e in c.value for e in ERRS)]
print(f"ERRORS {len(bad)}")
for b in bad[:40]: print("  ",b)
wbf=openpyxl.load_workbook(src)
for nm in (sys.argv[3].split(",") if len(sys.argv)>3 else []):
    dn=wbf.defined_names.get(nm)
    if not dn: print(f"  {nm}=<missing>"); continue
    for t,co in dn.destinations:
        print(f"  {nm} = {wb[t][co.replace('$','')].value}"); break
