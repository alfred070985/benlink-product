#!/usr/bin/env python3
import json
from collections import Counter
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / '来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx'
RESULTS = ROOT / '履约试算_上海全量'
OUTPUT = ROOT / '来福商品池_上海履约成本试算_全量_2026-11-02.xlsx'
TEMP = ROOT / '来福商品池_上海履约成本试算_全量_2026-11-02.tmp.xlsx'
API_URL = 'https://vm-online.benlai.com/bidding/fulfillmentCostEstimate/calculateOutboundPerformance'
SOURCE_NAME = '本来生活履约费用预估接口'
HEADERS = [
    '商品编号', '商品名称', '商品系统编号', '主商品编号', '定价状态', '发货方式',
    '仓库编码', '仓库名称', '收货区域编码', '收货区域名称',
    '试算省份编码', '试算省份名称', '试算城市编码', '试算城市名称', '试算区县编码', '试算区县名称',
    '渠道', '数量', '运输方式', '计算方式', '承运商',
    '试算履约成本（元）', '试算状态', '失败或不适用原因', '成本明细', 'Source Name', 'Source URL',
]

wb = load_workbook(SOURCE, read_only=True, data_only=True)
ws = wb['二次定价']
it = ws.iter_rows(values_only=True)
source_headers = next(it)
products = [dict(zip(source_headers, values)) for values in it]
wb.close()
assert len(products) == 43519
assert len({str(item['商品编号']) for item in products}) == len(products)
warehouse_codes = {str(item['商品编号']) for item in products if item.get('是否产地直发') == '仓发'}
assert len(warehouse_codes) == 2382
region_data = json.loads((ROOT / '履约试算_区域_上海.json').read_text(encoding='utf-8'))['data']
areas = [(area, area['provinceList'][0]) for area in region_data['area']]
assert len(areas) == 6


def load_result(code, area_id):
    path = RESULTS / f'{code}_{area_id}.json'
    result = json.loads(path.read_text(encoding='utf-8'))
    assert result.get('http_status') == 200, f'HTTP异常：{path}'
    response = result.get('response') or {}
    assert response.get('code') == 0 and response.get('data') is not None, f'接口异常或空响应：{path}'
    data = response['data']
    if data.get('success') is True:
        cost = data.get('totalComboCost')
        assert type(cost) in (int, float), f'成本不是数值：{path}'
        return data
    assert data.get('success') is False and data.get('errorMsg'), f'失败原因缺失：{path}'
    return data


results = {}
for code in warehouse_codes:
    for area, _ in areas:
        results[(code, area['areaId'])] = load_result(code, area['areaId'])
assert len(results) == 14292


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


out_wb = Workbook(write_only=True)
out_ws = out_wb.create_sheet('上海履约成本试算')
out_ws.freeze_panes = 'A2'
header_cells = []
for header in HEADERS:
    cell = WriteOnlyCell(out_ws, value=header)
    cell.fill = PatternFill('solid', fgColor='1F4E78')
    cell.font = Font(color='FFFFFF', bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    header_cells.append(cell)
out_ws.append(header_cells)
widths = [16, 38, 16, 16, 12, 12, 12, 12, 14, 14, 14, 18, 14, 18, 14, 18,
          12, 8, 12, 12, 12, 20, 12, 38, 72, 28, 82]
for index, width in enumerate(widths, 1):
    out_ws.column_dimensions[get_column_letter(index)].width = width

stats = Counter()
row_count = 0
for index, product in enumerate(products, 1):
    code = str(product['商品编号'])
    delivery = product.get('是否产地直发')
    for area, province in areas:
        if delivery == '仓发':
            data = results[(code, area['areaId'])]
            if data.get('success'):
                cost, status, reason = data['totalComboCost'], '成功', None
            else:
                cost, status, reason = None, '失败', data['errorMsg']
            details = detail_text(data)
        else:
            cost, status, reason, details = None, '不适用', '不适用：页面仅计算仓发商品（产地商品）', None
        out_ws.append([
            code, product.get('商品名称'), product.get('商品系统编号'), product.get('主商品编号'),
            product.get('定价状态'), delivery, 'W028', '上海', area['areaId'], area['areaName'],
            province['provinceCode'], province['provinceName'], province['cityCode'], province['cityName'],
            province['districtCode'], province['districtName'], '分销渠道', 1, '快递', '区域', '京东快递',
            cost, status, reason, details, SOURCE_NAME, API_URL,
        ])
        row_count += 1
        stats[(delivery, status)] += 1
        if cost is None:
            stats[(status, '成本留空')] += 1
        elif cost == 0:
            stats[('成功', '成本为0')] += 1
        else:
            stats[('成功', '成本大于0')] += 1
    if index % 4000 == 0:
        print(f'已生成商品 {index}/{len(products)}，数据行 {row_count}', flush=True)
assert row_count == 261114
out_ws.auto_filter.ref = f'A1:AA{row_count + 1}'
out_wb.save(TEMP)
TEMP.replace(OUTPUT)
summary = {
    '文件': str(OUTPUT), 'Sheet': '上海履约成本试算', '商品数': len(products),
    '仓发商品数': len(warehouse_codes), '代发商品数': len(products) - len(warehouse_codes),
    '数据行数': row_count, '列数': len(HEADERS), '接口试算任务数': len(results),
    '状态统计': {'|'.join(map(str, key)): value for key, value in sorted(stats.items(), key=lambda item: str(item[0]))},
}
(ROOT / '上海全量试算_生成统计.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
