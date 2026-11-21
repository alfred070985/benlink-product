#!/usr/bin/env python3
import json
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import requests
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / '来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx'
OUT = ROOT / '履约试算_其他仓发'
URL = 'https://vm-online.benlai.com/bidding/fulfillmentCostEstimate/calculateOutboundPerformance'
WAREHOUSES = [('广州', 'W006'), ('上海', 'W028'), ('武汉', 'W010')]

wb = load_workbook(SOURCE, read_only=True, data_only=True)
ws = wb['二次定价']
it = ws.iter_rows(values_only=True)
headers = next(it)
codes = sorted({str(row[headers.index('商品编号')]) for row in it
                if row[headers.index('定价状态')] == '已定价' and row[headers.index('是否产地直发')] == '仓发'})
wb.close()


def classify(result):
    if result.get('error'):
        return '网络错误'
    if result.get('http_status') != 200:
        return 'HTTP异常'
    response = result.get('response') or {}
    if response.get('code') != 0:
        return '接口异常'
    data = response.get('data')
    if data is None:
        return '空响应'
    if data.get('success') is True:
        return '成功'
    if data.get('success') is False:
        return '未查询到商品信息' if data.get('errorMsg') == '未查询到商品信息' else '试算失败'
    return '未知状态'


def result_key(result):
    data = (result.get('response') or {}).get('data') or {}
    return classify(result), data.get('totalComboCost'), data.get('errorMsg')

results = {}
summary = {'商品数': len(codes), '预期任务数': len(codes) * len(WAREHOUSES), '仓库': {}}
missing = []
for warehouse_name, warehouse_code in WAREHOUSES:
    statuses = Counter()
    costs = Counter()
    success_codes = []
    failed_codes = []
    for code in codes:
        path = OUT / f'{warehouse_code}_{code}.json'
        if not path.exists():
            missing.append(f'{warehouse_code}_{code}')
            continue
        result = json.loads(path.read_text(encoding='utf-8'))
        results[(warehouse_code, code)] = result
        status = classify(result)
        statuses[status] += 1
        if status == '成功':
            success_codes.append(code)
            data = result['response']['data']
            costs[str(data.get('totalComboCost'))] += 1
        elif status == '未查询到商品信息':
            failed_codes.append(code)
    summary['仓库'][warehouse_code] = {'状态统计': dict(statuses), '成功成本分布': dict(costs)}
    summary['仓库'][warehouse_code]['复核商品'] = {
        '成功': [success_codes[0], success_codes[len(success_codes) // 2], success_codes[-1]] if success_codes else [],
        '失败': [failed_codes[0]] if failed_codes else [],
    }
summary['实际结果数'] = len(results)
summary['缺失任务'] = missing
(ROOT / '其他仓试算_核对.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)

samples = []
for warehouse_name, warehouse_code in WAREHOUSES:
    data = json.loads((ROOT / f'履约试算_区域_{warehouse_name}.json').read_text(encoding='utf-8'))['data']
    selected = summary['仓库'][warehouse_code]['复核商品']
    for code in selected['成功'] + selected['失败']:
        for area in data['area']:
            province = area['provinceList'][0]
            samples.append({'warehouse_name': warehouse_name, 'warehouse_code': warehouse_code,
                            'code': code, 'area_id': area['areaId'], 'area_name': area['areaName'],
                            'province': province})


def calculate(sample):
    province = sample['province']
    body = {
        'warehouseCode': sample['warehouse_code'], 'businessId': 5, 'orderChannelId': 5,
        'provinceCode': str(province['provinceCode']), 'cityCode': str(province['cityCode']),
        'districtCode': str(province['districtCode']), 'logisticsCompanyCode': 'jingdong',
        'comboQty': 1, 'transportFeeType': 1, 'callSource': 1, 'optId': 0, 'optName': '系统批量试算',
        'goodsList': [{'goodsCode': sample['code'], 'qty': 1, 'isProductOrigin': False}],
    }
    try:
        response = requests.post(URL, json=body, timeout=45)
        result = {'sample': sample, 'http_status': response.status_code, 'response': response.json()}
    except Exception as exc:
        result = {'sample': sample, 'error': f'{type(exc).__name__}: {exc}'}
    result['与一区结果一致'] = result_key(result) == result_key(results[(sample['warehouse_code'], sample['code'])])
    return result

checks = []
with ThreadPoolExecutor(max_workers=6) as pool:
    futures = [pool.submit(calculate, sample) for sample in samples]
    for future in as_completed(futures):
        checks.append(future.result())
checks.sort(key=lambda item: (item['sample']['warehouse_code'], item['sample']['code'], item['sample']['area_id']))
(ROOT / '其他仓六区_复核.json').write_text(json.dumps(checks, ensure_ascii=False, indent=2), encoding='utf-8')
mismatches = [item for item in checks if not item['与一区结果一致']]
print(f'六区复核={len(checks)}，不一致={len(mismatches)}', flush=True)
if mismatches:
    print(json.dumps(mismatches, ensure_ascii=False, indent=2), flush=True)
