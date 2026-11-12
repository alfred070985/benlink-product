#!/usr/bin/env python3
from collections import Counter
from pathlib import Path
from openpyxl import load_workbook
root = Path(__file__).resolve().parent
for name in ['来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx',
             '来福商品池_成本及二次定价_回填上海履约成本_2026-11-02.xlsx']:
    wb = load_workbook(root / name, read_only=True, data_only=False)
    print(name)
    for ws in wb.worksheets:
        counts = Counter()
        examples = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == 'f':
                    counts[cell.column_letter] += 1
                    examples.setdefault(cell.column_letter, cell.value)
        print(ws.title, '公式数', sum(counts.values()), '按列', dict(counts))
        print('示例', examples)
    wb.close()
