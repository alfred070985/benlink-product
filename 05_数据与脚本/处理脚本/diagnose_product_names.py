#!/usr/bin/env python3
import json
from pathlib import Path
from openpyxl import load_workbook
root = Path(__file__).resolve().parent
result = []
for filename in ("来福商品池_成本及二次定价_2026-10-30.xlsx", "来福商品池_成本及二次定价_末次采购成本更新_2026-10-30.xlsx"):
    wb = load_workbook(root / filename, read_only=True, data_only=False)
    ws = wb["二次定价"]
    for row_no in (5456, 36169):
        value = ws.cell(row_no, 2).value
        result.append({"file": filename, "row": row_no, "value": value, "repr": repr(value), "data_type": ws.cell(row_no, 2).data_type})
    wb.close()
print(json.dumps(result, ensure_ascii=False, indent=2))
