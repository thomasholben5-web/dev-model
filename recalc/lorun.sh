#!/bin/bash
SRC="$1"; WORK="$2"; shift 2; NAMES="$*"
rm -rf "$WORK"; mkdir -p "$WORK/out"
python3 recalc/lo_prep.py "$SRC" "$WORK/in.xlsx" >/dev/null
pkill -9 -f soffice 2>/dev/null; sleep 1
nohup setsid env HOME="$WORK" soffice --headless --nolockcheck --norestore \
  --convert-to "xlsx:Calc MS Excel 2007 XML" --outdir "$WORK/out" "$WORK/in.xlsx" \
  >"$WORK/log" 2>&1 </dev/null &
for i in $(seq 1 90); do [ -f "$WORK/out/in.xlsx" ] && break; sleep 1; done
sleep 2
if [ -f "$WORK/out/in.xlsx" ]; then python3 recalc/lo_read.py "$WORK/out/in.xlsx" "$SRC" "$NAMES"; else echo "NO OUTPUT"; cat "$WORK/log"; fi
