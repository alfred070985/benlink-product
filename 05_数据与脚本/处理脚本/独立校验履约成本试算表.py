#!/usr/bin/env python3
import glob
import json
import re
import zipfile
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / '来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx'
TARGET = ROOT / '来福商品池_履约成本试算_2026-11-02.xlsx'
ATTR = ROOT / '已定价商品_北京履约属性'
BEIJING = ROOT / '履约试算_北京仓发'
OTHER = ROOT / '履约试算_其他仓发'
API_URL = 'https://vm-online.benlai.com/bidding/fulfillmentCostEstimate/calculateOutboundPerformance'
SOURCE_NAME = '本来生活履约费用预估接口'
KEYS = ['gross_weight', 'net_weight', 'length', 'width', 'height', 'vol', 'shape', 'attr',
        'storage_method', 'delivery_method', 'is_origin_pack', 'is_mix_pack', 'goods_type', 'package_count']
WAREHOUSE_FILES = {'W005': '北京', 'W006': '广州', 'W028': '上海', 'W010': '武汉'}
ERRORS = {'#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'}

source_wb = load_workbook(SOURCE, read_only=True, data_only=True)
source_ws = source_wb['二次定价']
it = source_ws.iter_rows(values_only=True)
source_headers = next(it)
products = {}
for values in it:
    item = dict(zip(source_headers, values))
    if item.get('定价状态') == '已定价':
        code = str(item['商品编号'])
        assert code not in products
        products[code] = item
source_wb.close()
assert len(products) == 31327
warehouse_products = {code for code, item in products.items() if item.get('是否产地直发') == '仓发'}
assert len(warehouse_products) == 1911

attrs = {}
for filename in glob.glob(str(ATTR / '*.json')):
    data = json.loads(Path(filename).read_text(encoding='utf-8'))['data']
    for values in data.get('rows') or []:
        item = dict(zip(data['column_list'], values))
        if item['goods_code'] in warehouse_products:
            attrs[item['goods_code']] = item
groups = {}
for code, item in attrs.items():
    groups.setdefault(tuple(item.get(key) for key in KEYS), []).append(code)
representative = {}
for members in groups.values():
    for code in members:
        representative[code] = sorted(members)[0]


def raw_result(path):
    result = json.loads(path.read_text(encoding='utf-8'))
    assert result.get('http_status') == 200
    response = result['response']
    assert response['code'] == 0 and response.get('data') is not None
    data = response['data']
    if data.get('success') is True:
        cost = data.get('totalComboCost')
        assert type(cost) in (int, float)
        return cost, '成功', None
    assert data.get('success') is False
    return None, '失败', data.get('errorMsg') or '试算失败'


expected_outcomes = {}
for code in warehouse_products:
    rep = representative.get(code, code)
    expected_outcomes[(code, 'W005')] = raw_result(BEIJING / f'{rep}.json')
    for warehouse_code in ('W006', 'W028', 'W010'):
        expected_outcomes[(code, warehouse_code)] = raw_result(OTHER / f'{warehouse_code}_{code}.json')

regions = {}
expected_keys = set()
for warehouse_code, name in WAREHOUSE_FILES.items():
    data = json.loads((ROOT / f'履约试算_区域_{name}.json').read_text(encoding='utf-8'))['data']
    assert len(data['area']) == 6
    for area in data['area']:
        province = area['provinceList'][0]
        key = (warehouse_code, area['areaId'])
        expected_keys.add(key)
        regions[key] = (warehouse_code, name, area, province)
assert len(expected_keys) == 24

# ZIP 完整性、公式和错误检查（流式，不将整个 XML 读入内存）。
formula_count = 0
with zipfile.ZipFile(TARGET) as archive:
    assert archive.testzip() is None, 'XLSX ZIP 包存在损坏'
    sheet_files = [name for name in archive.namelist() if name.startswith('xl/worksheets/sheet') and name.endswith('.xml')]
    assert len(sheet_files) == 1
    pattern = re.compile(br'<f(?:\s|>|/)')
    for sheet_file in sheet_files:
        tail = b''
        with archive.open(sheet_file) as stream:
            while True:
                chunk = stream.read(1024 * 1024)
                if not chunk:
                    break
                data = tail + chunk
                formula_count += len(pattern.findall(data))
                tail = data[-8:]
assert formula_count == 0, f'发现公式：{formula_count}'

wb = load_workbook(TARGET, read_only=True, data_only=False)
assert wb.sheetnames == ['履约成本试算结果'], wb.sheetnames
ws = wb['履约成本试算结果']
it = ws.iter_rows(values_only=True)
headers = next(it)
index = {header: position for position, header in enumerate(headers)}
required = ['商品编号', '商品名称', '商品系统编号', '主商品编号', '发货方式', '仓库编码', '仓库名称',
            '收货区域编码', '收货区域名称', '试算省份编码', '试算省份名称', '试算城市编码', '试算城市名称',
            '试算区县编码', '试算区县名称', '渠道', '数量', '运输方式', '计算方式', '承运商',
            '试算履约成本（元）', '试算状态', '失败或不适用原因', '成本明细', 'Source Name', 'Source URL']
assert headers == tuple(required), headers

stats = Counter()
completed = set()
current_code = None
current_keys = set()
current_rows = 0
row_count = 0


def finish_product(code, keys, rows):
    assert code is not None
    assert rows == 24, f'{code} 行数={rows}'
    assert keys == expected_keys, f'{code} 组合缺失={expected_keys - keys}，额外={keys - expected_keys}'
    completed.add(code)


for values in it:
    row_count += 1
    for value in values:
        assert value not in ERRORS, f'第{row_count + 1}行存在公式错误值：{value}'
    code = str(values[index['商品编号']])
    assert code in products, f'未知商品：{code}'
    if code != current_code:
        if current_code is not None:
            finish_product(current_code, current_keys, current_rows)
        assert code not in completed, f'商品不是连续24行或重复出现：{code}'
        current_code, current_keys, current_rows = code, set(), 0
    current_rows += 1
    product = products[code]
    assert values[index['商品名称']] == product['商品名称']
    assert values[index['商品系统编号']] == product['商品系统编号']
    assert values[index['主商品编号']] == product['主商品编号']
    assert values[index['发货方式']] == product['是否产地直发']
    warehouse_code = values[index['仓库编码']]
    area_id = values[index['收货区域编码']]
    key = (warehouse_code, area_id)
    assert key in expected_keys, f'未知组合：{key}'
    assert key not in current_keys, f'{code} 重复组合：{key}'
    current_keys.add(key)
    expected_warehouse, warehouse_name, area, province = regions[key]
    assert values[index['仓库名称']] == warehouse_name
    assert values[index['收货区域名称']] == area['areaName']
    assert values[index['试算省份编码']] == province['provinceCode']
    assert values[index['试算省份名称']] == province['provinceName']
    assert values[index['试算城市编码']] == province['cityCode']
    assert values[index['试算城市名称']] == province['cityName']
    assert values[index['试算区县编码']] == province['districtCode']
    assert values[index['试算区县名称']] == province['districtName']
    assert values[index['渠道']] == '分销渠道'
    assert values[index['数量']] == 1
    assert values[index['运输方式']] == '快递'
    assert values[index['计算方式']] == '区域'
    assert values[index['承运商']] == '京东快递'
    assert values[index['Source Name']] == SOURCE_NAME
    assert values[index['Source URL']] == API_URL

    cost = values[index['试算履约成本（元）']]
    status = values[index['试算状态']]
    reason = values[index['失败或不适用原因']]
    delivery = values[index['发货方式']]
    if delivery == '仓发':
        expected_cost, expected_status, expected_reason = expected_outcomes[(code, warehouse_code)]
        assert status == expected_status, f'{code} {key} 状态 {status} != {expected_status}'
        assert reason == expected_reason, f'{code} {key} 原因 {reason} != {expected_reason}'
        if expected_cost is None:
            assert cost is None, f'{code} {key} 失败行成本未留空'
        else:
            assert type(cost) in (int, float), f'{code} {key} 成本不是数值：{cost!r}'
            assert cost == expected_cost, f'{code} {key} 成本 {cost} != {expected_cost}'
    else:
        assert delivery == '代发'
        assert cost is None and status == '不适用'
        assert reason == '不适用：页面仅计算仓发商品（产地商品）'
    stats[(delivery, warehouse_code, status)] += 1
    if cost is None:
        stats[(status, '成本留空')] += 1
    elif cost == 0:
        stats[('成功', '成本为0')] += 1
    else:
        assert cost > 0
        stats[('成功', '成本大于0')] += 1
        assert warehouse_code == 'W005' and cost == 1.16
        assert values[index['成本明细']]
    if row_count % 100000 == 0:
        print(f'已校验 {row_count}/751848 行', flush=True)

if current_code is not None:
    finish_product(current_code, current_keys, current_rows)
wb.close()
assert row_count == 751848, row_count
assert len(completed) == 31327
assert completed == set(products)
assert stats[('成功', '成本大于0')] == 10224
assert stats[('成功', '成本为0')] == 31350
assert stats[('失败', '成本留空')] == 4290
assert stats[('不适用', '成本留空')] == 705984

summary = {
    'QA状态': '通过', 'Sheet数': 1, 'Sheet名称': '履约成本试算结果',
    '数据行数': row_count, '商品数': len(completed), '每商品组合数': 24,
    '组合总数': row_count, '重复或遗漏组合': 0, '公式数': formula_count, '公式错误数': 0,
    '成功且成本大于0': stats[('成功', '成本大于0')],
    '成功且成本为0': stats[('成功', '成本为0')],
    '失败且成本留空': stats[('失败', '成本留空')],
    '不适用且成本留空': stats[('不适用', '成本留空')],
    '分仓状态统计': {'|'.join(map(str, key)): value for key, value in sorted(stats.items(), key=lambda item: str(item[0]))
                  if len(key) == 3},
}
(ROOT / '履约成本试算_独立QA.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
