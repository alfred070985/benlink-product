#!/usr/bin/env python3
import json
import re
import zipfile
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / '来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx'
TARGET = ROOT / '来福商品池_上海履约成本试算_全量_2026-11-02.xlsx'
RESULTS = ROOT / '履约试算_上海全量'
API_URL = 'https://vm-online.benlai.com/bidding/fulfillmentCostEstimate/calculateOutboundPerformance'
SOURCE_NAME = '本来生活履约费用预估接口'
ERRORS = {'#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'}
EXPECTED_HEADERS = (
    '商品编号', '商品名称', '商品系统编号', '主商品编号', '定价状态', '发货方式',
    '仓库编码', '仓库名称', '收货区域编码', '收货区域名称',
    '试算省份编码', '试算省份名称', '试算城市编码', '试算城市名称', '试算区县编码', '试算区县名称',
    '渠道', '数量', '运输方式', '计算方式', '承运商',
    '试算履约成本（元）', '试算状态', '失败或不适用原因', '成本明细', 'Source Name', 'Source URL',
)

source_wb = load_workbook(SOURCE, read_only=True, data_only=True)
source_ws = source_wb['二次定价']
it = source_ws.iter_rows(values_only=True)
source_headers = next(it)
products = {}
for values in it:
    item = dict(zip(source_headers, values))
    code = str(item['商品编号'])
    assert code not in products
    products[code] = item
source_wb.close()
assert len(products) == 43519
warehouse_products = {code for code, item in products.items() if item.get('是否产地直发') == '仓发'}
assert len(warehouse_products) == 2382
region_data = json.loads((ROOT / '履约试算_区域_上海.json').read_text(encoding='utf-8'))['data']
regions = {area['areaId']: (area, area['provinceList'][0]) for area in region_data['area']}
assert len(regions) == 6

expected = {}
for code in warehouse_products:
    for area_id in regions:
        path = RESULTS / f'{code}_{area_id}.json'
        result = json.loads(path.read_text(encoding='utf-8'))
        assert result.get('http_status') == 200
        response = result.get('response') or {}
        assert response.get('code') == 0 and response.get('data') is not None
        data = response['data']
        if data.get('success') is True:
            cost = data.get('totalComboCost')
            assert type(cost) in (int, float)
            expected[(code, area_id)] = (cost, '成功', None)
        else:
            assert data.get('success') is False and data.get('errorMsg')
            expected[(code, area_id)] = (None, '失败', data['errorMsg'])
assert len(expected) == 14292

formula_count = 0
with zipfile.ZipFile(TARGET) as archive:
    assert archive.testzip() is None
    sheets = [name for name in archive.namelist() if name.startswith('xl/worksheets/sheet') and name.endswith('.xml')]
    assert len(sheets) == 1
    pattern = re.compile(br'<f(?:\s|>|/)')
    for sheet in sheets:
        tail = b''
        with archive.open(sheet) as stream:
            while True:
                chunk = stream.read(1024 * 1024)
                if not chunk:
                    break
                data = tail + chunk
                formula_count += len(pattern.findall(data))
                tail = data[-8:]
assert formula_count == 0

wb = load_workbook(TARGET, read_only=True, data_only=False)
assert wb.sheetnames == ['上海履约成本试算']
ws = wb['上海履约成本试算']
it = ws.iter_rows(values_only=True)
headers = next(it)
assert headers == EXPECTED_HEADERS
for forbidden in ('请求参数JSON', '接口完整返回值', 'traceId', 'HTTP状态码'):
    assert forbidden not in headers
index = {header: position for position, header in enumerate(headers)}
stats = Counter()
completed = set()
current_code = None
current_areas = set()
current_rows = 0
row_count = 0


def finish_product(code, area_set, rows):
    assert rows == 6, f'{code} 行数={rows}'
    assert area_set == set(regions), f'{code} 区域不完整：{area_set}'
    completed.add(code)


for values in it:
    row_count += 1
    for value in values:
        assert value not in ERRORS, f'第{row_count + 1}行存在错误值：{value}'
    code = str(values[index['商品编号']])
    assert code in products
    if code != current_code:
        if current_code is not None:
            finish_product(current_code, current_areas, current_rows)
        assert code not in completed
        current_code, current_areas, current_rows = code, set(), 0
    current_rows += 1
    product = products[code]
    assert values[index['商品名称']] == product['商品名称']
    assert values[index['商品系统编号']] == product['商品系统编号']
    assert values[index['主商品编号']] == product['主商品编号']
    assert values[index['定价状态']] == product['定价状态']
    assert values[index['发货方式']] == product['是否产地直发']
    assert values[index['仓库编码']] == 'W028' and values[index['仓库名称']] == '上海'
    area_id = values[index['收货区域编码']]
    assert area_id in regions and area_id not in current_areas
    current_areas.add(area_id)
    area, province = regions[area_id]
    assert values[index['收货区域名称']] == area['areaName']
    assert values[index['试算省份编码']] == province['provinceCode']
    assert values[index['试算省份名称']] == province['provinceName']
    assert values[index['试算城市编码']] == province['cityCode']
    assert values[index['试算城市名称']] == province['cityName']
    assert values[index['试算区县编码']] == province['districtCode']
    assert values[index['试算区县名称']] == province['districtName']
    assert values[index['渠道']] == '分销渠道' and values[index['数量']] == 1
    assert values[index['运输方式']] == '快递' and values[index['计算方式']] == '区域'
    assert values[index['承运商']] == '京东快递'
    assert values[index['Source Name']] == SOURCE_NAME and values[index['Source URL']] == API_URL

    delivery = values[index['发货方式']]
    cost = values[index['试算履约成本（元）']]
    status = values[index['试算状态']]
    reason = values[index['失败或不适用原因']]
    if delivery == '仓发':
        expected_cost, expected_status, expected_reason = expected[(code, area_id)]
        assert status == expected_status and reason == expected_reason
        if expected_cost is None:
            assert cost is None
        else:
            assert type(cost) in (int, float) and cost == expected_cost
    else:
        assert delivery == '代发'
        assert cost is None and status == '不适用'
        assert reason == '不适用：页面仅计算仓发商品（产地商品）'
    stats[(delivery, status)] += 1
    if cost is None:
        stats[(status, '成本留空')] += 1
    elif cost == 0:
        stats[('成功', '成本为0')] += 1
    else:
        assert cost > 0
        stats[('成功', '成本大于0')] += 1
    if row_count % 50000 == 0:
        print(f'已校验 {row_count}/261114 行', flush=True)
if current_code is not None:
    finish_product(current_code, current_areas, current_rows)
wb.close()
assert row_count == 261114
assert completed == set(products) and len(completed) == 43519
assert stats[('仓发', '成功')] == 12732
assert stats[('仓发', '失败')] == 1560
assert stats[('代发', '不适用')] == 246822
assert stats[('成功', '成本大于0')] == 12732
assert stats[('成功', '成本为0')] == 0
assert stats[('失败', '成本留空')] == 1560
assert stats[('不适用', '成本留空')] == 246822
summary = {
    'QA状态': '通过', 'Sheet数': 1, '数据行数': row_count, '商品数': len(completed),
    '每商品区域数': 6, '重复或遗漏组合': 0, '公式数': formula_count, '公式错误数': 0,
    '仓发成功且成本大于0': stats[('成功', '成本大于0')],
    '仓发失败且成本留空': stats[('失败', '成本留空')],
    '代发不适用且成本留空': stats[('不适用', '成本留空')],
    '包含接口入参或完整返回值': False,
}
(ROOT / '上海全量试算_独立QA.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
