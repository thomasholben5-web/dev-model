import sys, time, formulas, warnings
warnings.filterwarnings("ignore")
fn=sys.argv[1]
outs=sys.argv[2].split(",") if len(sys.argv)>2 else None
t=time.time()
xl=formulas.ExcelModel().loads(fn).finish()
base="'[%s]"%__import__("os").path.basename(fn).upper()
targets=[base+o.upper().replace("!","'!") for o in outs] if outs else None
sol=xl.calculate(outputs=targets) if targets else xl.calculate()
print("calc %.1fs, cells=%d"%(time.time()-t,len(sol)))
if outs:
    for o in outs:
        key=base+o.upper().replace("!","'!")
        v=sol.get(key)
        try: v=v.value[0,0]
        except Exception:
            try: v=v.value
            except Exception: pass
        print(f"  {o} = {v}")
