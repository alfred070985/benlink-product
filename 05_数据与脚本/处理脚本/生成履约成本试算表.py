#!/usr/bin/env python3
import glob
import json
from collections import Counter
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / '来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx'
OUTPUT = ROOT / '来福商品池_履约成本试算_2026-11-02.xlsx'
TEMP = ROOT / '来福商品池_履约成本试算_2026-11-02.tmp.xlsx'
BEIJING_RESULTS = ROOT / '履约试算_北京仓发'
OTHER_RESULTS = ROOT / '履约试算_其他仓发'
ATTR = ROOT / '已定价商品_北京履约属性'
API_URL = 'https://vm-online.benlai.com/bidding/fulfillmentCostEstimate/calculateOutboundPerformance'
SOURCE_NAME = '本来生活履约费用预估接口'
KEYS = ['gross_weight', 'net_weight', 'length', 'width', 'height', 'vol', 'shape', 'attr',
        'storage_method', 'delivery_method', 'is_origin_pack', 'is_mix_pack', 'goods_type', 'package_count']
WAREHOUSES = [
    ('北京', 'W005', '履约试算_区域_北京.json'),
    ('广州', 'W006', '履约试算_区域_广州.json'),
    ('上海', 'W028', '履约试算_区域_上海.json'),
    ('武汉', 'W010', '履约试算_区域_武汉.json'),
]
HEADERS = [
    '商品编号', '商品名称', '商品系统编号', '主商品编号', '发货方式',
    '仓库编码', '仓库名称', '收货区域编码', '收货区域名称',
    '试算省份编码', '试算省份名称', '试算城市编码', '试算城市名称', '试算区县编码', '试算区县名称',
    '渠道', '数量', '运输方式', '计算方式', '承运商',
    '试算履约成本（元）', '试算状态', '失败或不适用原因', '成本明细', 'Source Name', 'Source URL',
]

source_wb = load_workbook(SOURCE, read_only=True, data_only=True)
source_ws = source_wb['二次定价']
it = source_ws.iter_rows(values_only=True)
source_headers = next(it)
products = []
for values in it:
    item = dict(zip(source_headers, values))
    if item.get('定价状态') == '已定价':
        products.append(item)
source_wb.close()
assert len(products) == 31327, f'已定价商品数异常：{len(products)}'
assert len({str(item['商品编号']) for item in products}) == len(products), '商品编号存在重复'
warehouse_codes = {str(item['商品编号']) for item in products if item.get('是否产地直发') == '仓发'}
assert len(warehouse_codes) == 1911, f'仓发商品数异常：{len(warehouse_codes)}'

# 北京：同履约属性商品使用同组代表结果；所有多商品组均已做第二商品复核。
attrs = {}
for filename in glob.glob(str(ATTR / '*.json')):
    data = json.loads(Path(filename).read_text(encoding='utf-8'))['data']
    for values in data.get('rows') or []:
        item = dict(zip(data['column_list'], values))
        if item['goods_code'] in warehouse_codes:
            attrs[item['goods_code']] = item
groups = {}
for code, item in attrs.items():
    signature = tuple(item.get(key) for key in KEYS)
    groups.setdefault(signature, []).append(code)
representative = {}
for members in groups.values():
    rep = sorted(members)[0]
    for code in members:
        representative[code] = rep


def load_result(path):
    result = json.loads(path.read_text(encoding='utf-8'))
    assert result.get('http_status') == 200, f'HTTP异常：{path}'
    response = result.get('response') or {}
    assert response.get('code') == 0 and response.get('data') is not None, f'接口空响应或异常：{path}'
    data = response['data']
    assert data.get('success') in (True, False), f'未知成功状态：{path}'
    if data.get('success'):
        cost = data.get('totalComboCost')
        assert isinstance(cost, (int, float)) and not isinstance(cost, bool), f'成功结果成本不是数值：{path}'
    else:
        assert data.get('errorMsg'), f'失败结果缺少原因：{path}'
    return data


beijing_results = {}
for code in warehouse_codes:
    result_code = representative.get(code, code)
    beijing_results[code] = load_result(BEIJING_RESULTS / f'{result_code}.json')
other_results = {}
for _, warehouse_code, _ in WAREHOUSES[1:]:
    for code in warehouse_codes:
        other_results[(warehouse_code, code)] = load_result(OTHER_RESULTS / f'{warehouse_code}_{code}.json')

regions = {}
for warehouse_name, warehouse_code, filename in WAREHOUSES:
    data = json.loads((ROOT / filename).read_text(encoding='utf-8'))['data']
    areas = []
    for area in data['area']:
        assert len(area['provinceList']) >= 1
        areas.append((area, area['provinceList'][0]))
    assert len(areas) == 6, f'{warehouse_name}收货区域数异常：{len(areas)}'
    regions[warehouse_code] = areas


def detail_text(data):
    if not data.get('success'):
        return None
    details = []
    for detail in data.get('costDetails') or []:
        text = f"{detail.get('costTypeName', '费用')}：{detail.get('amount', 0)}元"
        items = []
        for item in detail.get('items') or []:
            items.append(f"{item.get('itemName')}（{item.get('itemCode')}）×{item.get('qty')}，"
                         f"单价{item.get('price')}元，金额{item.get('amount')}元")
        if items:
            text += '；' + '；'.join(items)
        details.append(text)
    return '；'.join(details) or None


def outcome(product_code, warehouse_code):
    if warehouse_code == 'W005':
        data = beijing_results[product_code]
    else:
        data = other_results[(warehouse_code, product_code)]
    if data.get('success'):
        return data['totalComboCost'], '成功', None, detail_text(data)
    reason = data.get('errorMsg') or '试算失败'
    return None, '失败', reason, None


out_wb = Workbook(write_only=True)
out_ws = out_wb.create_sheet('履约成本试算结果')
out_ws.freeze_panes = 'A2'
header_fill = PatternFill('solid', fgColor='1F4E78')
header_font = Font(color='FFFFFF', bold=True)
header_cells = []
for header in HEADERS:
    cell = WriteOnlyCell(out_ws, value=header)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal='center', vertical='center')
    header_cells.append(cell)
out_ws.append(header_cells)
widths = [16, 38, 16, 16, 12, 12, 12, 14, 14, 14, 18, 14, 18, 14, 18,
          12, 8, 12, 12, 12, 20, 12, 40, 72, 28, 82]
for index, width in enumerate(widths, 1):
    out_ws.column_dimensions[get_column_letter(index)].width = width

stats = Counter()
row_count = 0
for index, product in enumerate(products, 1):
    code = str(product['商品编号'])
    delivery = product.get('是否产地直发')
    for warehouse_name, warehouse_code, _ in WAREHOUSES:
        for area, province in regions[warehouse_code]:
            if delivery == '仓发':
                cost, status, reason, details = outcome(code, warehouse_code)
            else:
                cost, status, reason, details = None, '不适用', '不适用：页面仅计算仓发商品（产地商品）', None
            out_ws.append([
                code, product.get('商品名称'), product.get('商品系统编号'), product.get('主商品编号'), delivery,
                warehouse_code, warehouse_name, area['areaId'], area['areaName'],
                province['provinceCode'], province['provinceName'], province['cityCode'], province['cityName'],
                province['districtCode'], province['districtName'],
                '分销渠道', 1, '快递', '区域', '京东快递',
                cost, status, reason, details, SOURCE_NAME, API_URL,
            ])
            row_count += 1
            stats[(delivery, warehouse_code, status)] += 1
            if cost is None:
                stats[(status, '成本留空')] += 1
            elif cost == 0:
                stats[('成功', '成本为0')] += 1
            else:
                stats[('成功', '成本大于0')] += 1
    if index % 2000 == 0:
        print(f'已生成商品 {index}/{len(products)}，数据行 {row_count}', flush=True)

assert row_count == 751848, f'结果行数异常：{row_count}'
out_ws.auto_filter.ref = f'A1:Z{row_count + 1}'
out_wb.save(TEMP)
TEMP.replace(OUTPUT)
summary = {
    '文件': str(OUTPUT), 'Sheet': '履约成本试算结果', '商品数': len(products),
    '仓发商品数': len(warehouse_codes), '代发商品数': len(products) - len(warehouse_codes),
    '数据行数': row_count, '列数': len(HEADERS),
    '状态统计': {'|'.join(map(str, key)): value for key, value in sorted(stats.items(), key=lambda item: str(item[0]))},
}
(ROOT / '履约成本试算_生成统计.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
