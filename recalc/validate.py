import sys, openpyxl
ERRS={'#REF!','#DIV/0!','#VALUE!','#N/A','#NAME?','#NUM!','#NULL!'}
def scan(path):
    wb=openpyxl.load_workbook(path, data_only=True)
    bad=[]
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value,str) and c.value in ERRS:
                    bad.append(f"{ws.title}!{c.coordinate}={c.value}")
    return bad
if __name__=="__main__":
    b=scan(sys.argv[1])
    print(f"ERRORS: {len(b)}")
    for x in b[:80]: print("  ",x)
