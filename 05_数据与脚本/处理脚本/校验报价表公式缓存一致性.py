#!/usr/bin/env python3
import json
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / '来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx'
TARGET = ROOT / '来福商品池_成本及二次定价_回填上海履约成本_2026-11-02.xlsx'
source_wb = load_workbook(SOURCE, read_only=True, data_only=True)
target_wb = load_workbook(TARGET, read_only=True, data_only=True)
assert source_wb.sheetnames == target_wb.sheetnames
compared_cells = 0
compared_rows = 0
for sheet_name in source_wb.sheetnames:
    source_ws = source_wb[sheet_name]
    target_ws = target_wb[sheet_name]
    source_it = source_ws.iter_rows(values_only=True)
    target_it = target_ws.iter_rows(values_only=True)
    rows = 0
    while True:
        source_row = next(source_it, None)
        target_row = next(target_it, None)
        if source_row is None or target_row is None:
            assert source_row is None and target_row is None, f'{sheet_name} 行数变化'
            break
        rows += 1
        if sheet_name == '二次定价':
            assert len(source_row) == 31 and len(target_row) == 34
            assert target_row[:31] == source_row, f'{sheet_name} 第{rows}行原缓存值变化'
            compared_cells += 31
        else:
            assert target_row == source_row, f'{sheet_name} 第{rows}行缓存值变化'
            compared_cells += len(source_row)
    compared_rows += rows
    print(f'{sheet_name}：{rows}行缓存值一致', flush=True)
source_wb.close()
target_wb.close()
summary = {'QA状态': '通过', 'Sheet数': len(source_wb.sheetnames), '对比行数': compared_rows,
           '对比原字段单元格数': compared_cells, '原公式及输入对应的缓存结果变化数': 0}
(ROOT / '回填上海报价表_缓存一致性QA.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
