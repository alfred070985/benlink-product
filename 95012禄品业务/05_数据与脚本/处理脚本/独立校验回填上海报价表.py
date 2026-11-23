#!/usr/bin/env python3
import json
import re
import zipfile
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / '来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx'
SHANGHAI = ROOT / '来福商品池_上海履约成本试算_全量_2026-11-02.xlsx'
TARGET = ROOT / '来福商品池_成本及二次定价_回填上海履约成本_2026-11-02.xlsx'
NEW_HEADERS = ('上海履约成本最低（元）', '上海履约成本最高（元）', '上海履约成本六区JSON')
ERRORS = {'#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'}

with zipfile.ZipFile(SOURCE) as source_zip, zipfile.ZipFile(TARGET) as target_zip:
    assert source_zip.testzip() is None and target_zip.testzip() is None
    for prefix in ('xl/charts/', 'xl/drawings/', 'xl/media/'):
        source_parts = {name for name in source_zip.namelist() if name.startswith(prefix)}
        target_parts = {name for name in target_zip.namelist() if name.startswith(prefix)}
        assert source_parts == target_parts, f'{prefix} 部件发生变化'
    formula_count = 0
    pattern = re.compile(br'<f(?:\s|>|/)')
    for name in target_zip.namelist():
        if not (name.startswith('xl/worksheets/sheet') and name.endswith('.xml')):
            continue
        tail = b''
        with target_zip.open(name) as stream:
            while True:
                chunk = stream.read(1024 * 1024)
                if not chunk:
                    break
                data = tail + chunk
                formula_count += len(pattern.findall(data))
                tail = data[-8:]

source_wb = load_workbook(SOURCE, read_only=True, data_only=False)
target_wb = load_workbook(TARGET, read_only=True, data_only=False)
result_wb = load_workbook(SHANGHAI, read_only=True, data_only=True)
assert source_wb.sheetnames == target_wb.sheetnames
source_ws = source_wb['二次定价']
target_ws = target_wb['二次定价']
result_ws = result_wb['上海履约成本试算']
source_it = source_ws.iter_rows(values_only=True)
target_it = target_ws.iter_rows(values_only=True)
result_it = result_ws.iter_rows(values_only=True)
source_headers = next(source_it)
target_headers = next(target_it)
result_headers = next(result_it)
assert len(source_headers) == 31 and len(target_headers) == 34
assert target_headers[:31] == source_headers
assert target_headers[31:] == NEW_HEADERS
index = {header: position for position, header in enumerate(result_headers)}
stats = Counter()


def expected_payload(code, rows):
    first = rows[0]
    area_results = []
    costs = []
    for row in rows:
        assert str(row[index['商品编号']]) == code
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
    return costs, payload


for row_number in range(2, 43521):
    source_row = next(source_it)
    target_row = next(target_it)
    assert len(target_row) == 34
    assert target_row[:31] == source_row, f'第{row_number}行原字段被修改'
    for value in target_row:
        assert value not in ERRORS, f'第{row_number}行存在错误值：{value}'
    code = str(target_row[0])
    result_rows = [next(result_it) for _ in range(6)]
    costs, expected = expected_payload(code, result_rows)
    minimum, maximum, json_text = target_row[31], target_row[32], target_row[33]
    if costs:
        assert type(minimum) in (int, float) and minimum == min(costs), f'{code} 最低成本错误'
        assert type(maximum) in (int, float) and maximum == max(costs), f'{code} 最高成本错误'
        stats['有最低最高'] += 1
    else:
        assert minimum is None and maximum is None, f'{code} 无成功成本但最低最高未留空'
        stats['无成功成本'] += 1
    assert isinstance(json_text, str) and len(json_text) <= 32767
    parsed = json.loads(json_text)
    assert parsed == expected, f'{code} JSON内容不一致'
    assert len(parsed['区域结果']) == 6
    delivery = result_rows[0][index['发货方式']]
    status = result_rows[0][index['试算状态']]
    stats[(delivery, status)] += 1
    if row_number % 5000 == 0:
        print(f'已校验商品 {row_number - 1}/43519', flush=True)

assert next(source_it, None) is None
assert next(target_it, None) is None
assert next(result_it, None) is None
source_wb.close()
target_wb.close()
result_wb.close()
assert stats['有最低最高'] == 2122
assert stats['无成功成本'] == 41397
assert stats[('仓发', '成功')] == 2122
assert stats[('仓发', '失败')] == 260
assert stats[('代发', '不适用')] == 41137
summary = {
    'QA状态': '通过', '商品数': 43519, '原有31列保持不变': True,
    '新增列数': 3, '有数值最低最高商品': stats['有最低最高'],
    '最低最高留空商品': stats['无成功成本'], 'JSON可解析且六区完整商品': 43519,
    '仓发失败商品': stats[('仓发', '失败')], '代发不适用商品': stats[('代发', '不适用')],
    '公式数': formula_count, '公式错误数': 0,
}
(ROOT / '回填上海报价表_独立QA.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
