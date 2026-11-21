#!/usr/bin/env python3
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / '来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx'
OUTPUT = ROOT / 'P0000322178_履约成本试算_2026-11-02.xlsx'
PRODUCT_CODE = 'P0000322178'
API_URL = 'https://vm-online.benlai.com/bidding/fulfillmentCostEstimate/calculateOutboundPerformance'
WAREHOUSES = [
    ('北京', 'W005', '履约试算_区域_北京.json'),
    ('广州', 'W006', '履约试算_区域_广州.json'),
    ('上海', 'W028', '履约试算_区域_上海.json'),
    ('武汉', 'W010', '履约试算_区域_武汉.json'),
]

wb = load_workbook(SOURCE, read_only=True, data_only=True)
ws = wb['二次定价']
it = ws.iter_rows(values_only=True)
headers = next(it)
product = None
for values in it:
    item = dict(zip(headers, values))
    if str(item.get('商品编号')) == PRODUCT_CODE:
        product = item
        break
wb.close()
assert product is not None, f'定价表未找到 {PRODUCT_CODE}'
assert product.get('定价状态') == '已定价', f"商品定价状态为：{product.get('定价状态')}"
is_origin = product.get('是否产地直发') != '仓发'
print(f"商品={PRODUCT_CODE}，名称={product.get('商品名称')}，发货方式={product.get('是否产地直发')}", flush=True)

tasks = []
for warehouse_name, warehouse_code, filename in WAREHOUSES:
    data = json.loads((ROOT / filename).read_text(encoding='utf-8'))['data']
    assert len(data['area']) == 6
    for area in data['area']:
        province = area['provinceList'][0]
        tasks.append({
            'warehouse_name': warehouse_name, 'warehouse_code': warehouse_code,
            'area_id': area['areaId'], 'area_name': area['areaName'], 'province': province,
        })
assert len(tasks) == 24


def calculate(task):
    province = task['province']
    body = {
        'warehouseCode': task['warehouse_code'], 'businessId': 5, 'orderChannelId': 5,
        'provinceCode': str(province['provinceCode']), 'cityCode': str(province['cityCode']),
        'districtCode': str(province['districtCode']), 'logisticsCompanyCode': 'jingdong',
        'comboQty': 1, 'transportFeeType': 1, 'callSource': 1, 'optId': 0, 'optName': '系统批量试算',
        'goodsList': [{'goodsCode': PRODUCT_CODE, 'qty': 1, 'isProductOrigin': is_origin}],
    }
    result = {'task': task, 'request': body}
    try:
        response = requests.post(API_URL, json=body, timeout=60)
        result['http_status'] = response.status_code
        result['raw_response'] = response.text
        try:
            result['response'] = response.json()
        except Exception:
            result['response'] = None
            result['request_error'] = '接口返回不是有效JSON'
    except Exception as exc:
        result['request_error'] = f'{type(exc).__name__}: {exc}'
        result['raw_response'] = None
        result['response'] = None
    return result


results = [None] * len(tasks)
with ThreadPoolExecutor(max_workers=4) as pool:
    futures = {pool.submit(calculate, task): index for index, task in enumerate(tasks)}
    for future in as_completed(futures):
        index = futures[future]
        result = future.result()
        results[index] = result
        data = (result.get('response') or {}).get('data') or {}
        print(f"{index + 1}/24 {result['task']['warehouse_name']} {result['task']['area_name']} "
              f"HTTP={result.get('http_status')} success={data.get('success')} "
              f"cost={data.get('totalComboCost')} error={data.get('errorMsg') or result.get('request_error')}", flush=True)

headers_out = [
    '商品编号', '商品名称', '商品系统编号', '主商品编号', '发货方式',
    '仓库编码', '仓库名称', '收货区域编码', '收货区域名称',
    '试算省份编码', '试算省份名称', '试算城市编码', '试算城市名称', '试算区县编码', '试算区县名称',
    '渠道', '数量', '运输方式', '计算方式', '承运商',
    'HTTP状态码', '接口code', '接口success', '试算状态', '试算履约成本（元）', 'totalCost',
    '错误原因', '成本明细JSON', 'traceId', '请求参数JSON', '接口完整返回值', 'Source Name', 'Source URL',
]
out_wb = Workbook()
out_ws = out_wb.active
out_ws.title = '试算结果'
out_ws.append(headers_out)
for cell in out_ws[1]:
    cell.fill = PatternFill('solid', fgColor='1F4E78')
    cell.font = Font(color='FFFFFF', bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')

for result in results:
    task = result['task']
    province = task['province']
    response = result.get('response') or {}
    data = response.get('data')
    if result.get('request_error'):
        status = '请求异常'
    elif response.get('code') != 0:
        status = '接口异常'
    elif data is None:
        status = '空响应'
    elif data.get('success') is True:
        status = '成功'
    else:
        status = '失败'
    data = data or {}
    reason = result.get('request_error') or data.get('errorMsg') or response.get('message')
    cost = data.get('totalComboCost') if status == '成功' else None
    out_ws.append([
        PRODUCT_CODE, product.get('商品名称'), product.get('商品系统编号'), product.get('主商品编号'),
        product.get('是否产地直发'), task['warehouse_code'], task['warehouse_name'],
        task['area_id'], task['area_name'], province['provinceCode'], province['provinceName'],
        province['cityCode'], province['cityName'], province['districtCode'], province['districtName'],
        '分销渠道', 1, '快递', '区域', '京东快递',
        result.get('http_status'), response.get('code'), data.get('success'), status, cost, data.get('totalCost'),
        reason, json.dumps(data.get('costDetails'), ensure_ascii=False) if data.get('costDetails') is not None else None,
        response.get('traceId'), json.dumps(result['request'], ensure_ascii=False),
        result.get('raw_response'), '本来生活履约费用预估接口', API_URL,
    ])

out_ws.freeze_panes = 'A2'
out_ws.auto_filter.ref = f'A1:AG{out_ws.max_row}'
widths = [16, 34, 16, 16, 12, 12, 12, 14, 14, 14, 18, 14, 18, 14, 18, 12, 8, 12, 12, 12,
          12, 12, 14, 12, 20, 14, 36, 70, 36, 90, 100, 28, 82]
for index, width in enumerate(widths, 1):
    out_ws.column_dimensions[get_column_letter(index)].width = width
for row in out_ws.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(vertical='top', wrap_text=cell.column in (27, 28, 29, 30, 31))
    row[24].number_format = '0.00'
    row[25].number_format = '0.00'
out_wb.save(OUTPUT)

check_wb = load_workbook(OUTPUT, read_only=True, data_only=True)
assert check_wb.sheetnames == ['试算结果']
check_ws = check_wb['试算结果']
assert check_ws.max_row == 25
assert check_ws.max_column == len(headers_out)
check_wb.close()
print(f'已生成：{OUTPUT}', flush=True)
