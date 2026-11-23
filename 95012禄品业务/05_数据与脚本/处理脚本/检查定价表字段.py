#!/usr/bin/env python3
from pathlib import Path
from openpyxl import load_workbook
root = Path(__file__).resolve().parent
path = root / '来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx'
wb = load_workbook(path, read_only=True, data_only=True)
ws = wb['二次定价']
it = ws.iter_rows(values_only=True)
headers = next(it)
print('列数：', len(headers))
for index, header in enumerate(headers, 1):
    print(index, repr(header))
for row in it:
    item = dict(zip(headers, row))
    if item.get('定价状态') == '已定价':
        print('首条已定价商品：')
        for key, value in item.items():
            print(repr(key), repr(value))
        break
wb.close()
