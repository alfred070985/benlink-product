#!/usr/bin/env python3
import json
from pathlib import Path
from openpyxl import load_workbook

path = Path(__file__).with_name("来福商品池_成本及二次定价_2026-10-30.xlsx")
wf = load_workbook(path, read_only=True, data_only=False)
wv = load_workbook(path, read_only=True, data_only=True)
result = {}
for name in ("汇总", "定价参数", "数据来源"):
    sf, sv = wf[name], wv[name]
    rows = []
    for rf, rv in zip(sf.iter_rows(values_only=True), sv.iter_rows(values_only=True)):
        rows.append({"formulas": rf, "cached": rv})
    result[name] = rows
print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
wf.close()
wv.close()
