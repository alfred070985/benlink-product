#!/usr/bin/env python3
import json
import zipfile
from collections import Counter
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
PATH = ROOT / "来福商品池_成本及二次定价_2026-10-30.xlsx"

result = {"path": str(PATH), "size": PATH.stat().st_size, "sheets": []}
with zipfile.ZipFile(PATH) as zf:
    names = zf.namelist()
    result["package_features"] = {
        "charts": sum(n.startswith("xl/charts/chart") for n in names),
        "drawings": sum(n.startswith("xl/drawings/drawing") and n.endswith(".xml") for n in names),
        "images": sum(n.startswith("xl/media/") for n in names),
        "pivots": sum(n.startswith("xl/pivotTables/") and n.endswith(".xml") for n in names),
        "external_links": sum(n.startswith("xl/externalLinks/externalLink") and n.endswith(".xml") for n in names),
    }

wb = load_workbook(PATH, read_only=True, data_only=False)
for ws in wb.worksheets:
    info = {"title": ws.title, "max_row": ws.max_row, "max_column": ws.max_column}
    it = ws.iter_rows(values_only=True)
    headers = list(next(it, ()))
    info["headers"] = headers
    formula_cells = 0
    formula_columns = Counter()
    blank_ids = 0
    ids = Counter()
    samples = []
    id_idx = headers.index("商品编号") if "商品编号" in headers else None
    for row_no, row in enumerate(it, start=2):
        for idx, value in enumerate(row):
            if isinstance(value, str) and value.startswith("="):
                formula_cells += 1
                formula_columns[headers[idx] if idx < len(headers) else str(idx + 1)] += 1
        if id_idx is not None:
            value = row[id_idx] if id_idx < len(row) else None
            if value is None or str(value).strip() == "":
                blank_ids += 1
            else:
                ids[str(value).strip()] += 1
        if row_no <= 3:
            samples.append(row)
    info["formula_cells"] = formula_cells
    info["formula_columns"] = dict(formula_columns)
    info["samples"] = samples
    if id_idx is not None:
        info["products"] = sum(ids.values())
        info["unique_products"] = len(ids)
        info["blank_product_ids"] = blank_ids
        info["duplicate_product_ids"] = sum(n - 1 for n in ids.values() if n > 1)
    result["sheets"].append(info)
wb.close()
print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
