#!/usr/bin/env python3
import json
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / '来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx'
SHANGHAI = ROOT / '来福商品池_上海履约成本试算_全量_2026-11-02.xlsx'
OUTPUT = ROOT / '来福商品池_成本及二次定价_回填上海履约成本_2026-11-02.xlsx'
TEMP = ROOT / '来福商品池_成本及二次定价_回填上海履约成本_2026-11-02.tmp.xlsx'
MIN_HEADER = '上海履约成本最低（元）'
MAX_HEADER = '上海履约成本最高（元）'
JSON_HEADER = '上海履约成本六区JSON'

wb = load_workbook(SOURCE)
assert '二次定价' in wb.sheetnames
ws = wb['二次定价']
source_headers = [cell.value for cell in ws[1]]
assert len(source_headers) == 31
for header in (MIN_HEADER, MAX_HEADER, JSON_HEADER):
    assert header not in source_headers
code_col = source_headers.index('商品编号') + 1
row_by_code = {}
for row in range(2, ws.max_row + 1):
    code = str(ws.cell(row, code_col).value)
    assert code not in row_by_code
    row_by_code[code] = row
assert len(row_by_code) == 43519

min_col = ws.max_column + 1
max_col = min_col + 1
json_col = min_col + 2
style_source = ws.cell(1, min_col - 1)
for col, header in ((min_col, MIN_HEADER), (max_col, MAX_HEADER), (json_col, JSON_HEADER)):
    cell = ws.cell(1, col, header)
    cell.font = copy(style_source.font)
    cell.fill = copy(style_source.fill)
    cell.border = copy(style_source.border)
    cell.alignment = copy(style_source.alignment)
    cell.protection = copy(style_source.protection)
ws.column_dimensions[get_column_letter(min_col)].width = 23
ws.column_dimensions[get_column_letter(max_col)].width = 23
ws.column_dimensions[get_column_letter(json_col)].width = 100
if ws.auto_filter.ref:
    ws.auto_filter.ref = f'A1:{get_column_letter(json_col)}{ws.max_row}'

result_wb = load_workbook(SHANGHAI, read_only=True, data_only=True)
result_ws = result_wb['上海履约成本试算']
it = result_ws.iter_rows(values_only=True)
headers = next(it)
index = {header: position for position, header in enumerate(headers)}
updated = set()
status_counter = {'有最低最高': 0, '无成功成本': 0}
current_code = None
current_rows = []


def flush(code, rows):
    assert code in row_by_code, f'报价表缺少商品：{code}'
    assert code not in updated, f'商品重复出现：{code}'
    assert len(rows) == 6, f'{code} 区域行数={len(rows)}'
    area_ids = [row[index['收货区域编码']] for row in rows]
    assert len(set(area_ids)) == 6
    first = rows[0]
    costs = []
    area_results = []
    for row in rows:
        cost = row[index['试算履约成本（元）']]
        status = row[index['试算状态']]
        if status == '成功':
            assert type(cost) in (int, float)
            costs.append(cost)
        else:
            assert cost is None
        area_results.append({
            '收货区域编码': row[index['收货区域编码']],
            '收货区域名称': row[index['收货区域名称']],
            '省': {'编码': row[index['试算省份编码']], '名称': row[index['试算省份名称']]},
            '市': {'编码': row[index['试算城市编码']], '名称': row[index['试算城市名称']]},
            '区县': {'编码': row[index['试算区县编码']], '名称': row[index['试算区县名称']]},
            '试算履约成本（元）': cost,
            '试算状态': status,
            '失败或不适用原因': row[index['失败或不适用原因']],
            '成本明细': row[index['成本明细']],
        })
    payload = {
        '商品编号': code,
        '商品名称': first[index['商品名称']],
        '仓库': {'编码': first[index['仓库编码']], '名称': first[index['仓库名称']]},
        '试算条件': {
            '渠道': first[index['渠道']], '数量': first[index['数量']],
            '运输方式': first[index['运输方式']], '计算方式': first[index['计算方式']],
            '承运商': first[index['承运商']],
        },
        '区域结果': area_results,
        'Source Name': first[index['Source Name']],
        'Source URL': first[index['Source URL']],
    }
    json_text = json.dumps(payload, ensure_ascii=False, separators=(',', ':'), allow_nan=False)
    assert len(json_text) <= 32767, f'{code} JSON长度={len(json_text)}'
    target_row = row_by_code[code]
    if costs:
        ws.cell(target_row, min_col, min(costs))
        ws.cell(target_row, max_col, max(costs))
        status_counter['有最低最高'] += 1
    else:
        ws.cell(target_row, min_col, None)
        ws.cell(target_row, max_col, None)
        status_counter['无成功成本'] += 1
    ws.cell(target_row, min_col).number_format = '0.###'
    ws.cell(target_row, max_col).number_format = '0.###'
    ws.cell(target_row, json_col, json_text)
    updated.add(code)


for values in it:
    code = str(values[index['商品编号']])
    if code != current_code:
        if current_code is not None:
            flush(current_code, current_rows)
        current_code, current_rows = code, []
    current_rows.append(values)
if current_code is not None:
    flush(current_code, current_rows)
result_wb.close()
assert len(updated) == 43519
assert updated == set(row_by_code)
assert status_counter['有最低最高'] == 2122
assert status_counter['无成功成本'] == 41397
wb.save(TEMP)
TEMP.replace(OUTPUT)
wb.close()
summary = {'文件': str(OUTPUT), '商品数': len(updated), '新增列': [MIN_HEADER, MAX_HEADER, JSON_HEADER], **status_counter}
print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
