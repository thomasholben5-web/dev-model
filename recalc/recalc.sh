#!/bin/bash
set -e
SRC="$1"; OUT="${2:-recalc/out}"
rm -rf "$OUT"; mkdir -p "$OUT"
soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$OUT" "$SRC" >/dev/null 2>&1
BASE=$(basename "$SRC")
echo "$OUT/$BASE"
