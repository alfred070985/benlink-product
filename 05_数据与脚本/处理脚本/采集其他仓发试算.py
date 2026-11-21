#!/usr/bin/env python3
import argparse
import json
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import requests
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / '来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx'
OUT = ROOT / '履约试算_其他仓发'
OUT.mkdir(exist_ok=True)
URL = 'https://vm-online.benlai.com/bidding/fulfillmentCostEstimate/calculateOutboundPerformance'
WAREHOUSES = [('广州', 'W006'), ('上海', 'W028'), ('武汉', 'W010')]

wb = load_workbook(SOURCE, read_only=True, data_only=True)
ws = wb['二次定价']
it = ws.iter_rows(values_only=True)
headers = next(it)
codes = sorted({str(row[headers.index('商品编号')]) for row in it
                if row[headers.index('定价状态')] == '已定价' and row[headers.index('是否产地直发')] == '仓发'})
wb.close()

regions = {}
for name, code in WAREHOUSES:
    data = json.loads((ROOT / f'履约试算_区域_{name}.json').read_text(encoding='utf-8'))['data']
    area = next(item for item in data['area'] if item['areaId'] == 1)
    regions[code] = area['provinceList'][0]


def resolved(path):
    if not path.exists():
        return False
    try:
        result = json.loads(path.read_text(encoding='utf-8'))
    except Exception:
        return False
    response = result.get('response') or {}
    return result.get('http_status') == 200 and response.get('code') == 0 and response.get('data') is not None


tasks = []
for warehouse_name, warehouse_code in WAREHOUSES:
    for code in codes:
        output = OUT / f'{warehouse_code}_{code}.json'
        if not resolved(output):
            tasks.append({'code': code, 'warehouse_name': warehouse_name, 'warehouse_code': warehouse_code})
parser = argparse.ArgumentParser()
parser.add_argument('--max-tasks', type=int, default=1500)
args = parser.parse_args()
pending = tasks[:args.max_tasks]
print(f'仓发商品={len(codes)}, 总任务={len(codes) * len(WAREHOUSES)}, 待处理={len(tasks)}, 本次={len(pending)}', flush=True)
thread_local = threading.local()


def calculate(task):
    province = regions[task['warehouse_code']]
    body = {
        'warehouseCode': task['warehouse_code'], 'businessId': 5, 'orderChannelId': 5,
        'provinceCode': str(province['provinceCode']), 'cityCode': str(province['cityCode']),
        'districtCode': str(province['districtCode']), 'logisticsCompanyCode': 'jingdong',
        'comboQty': 1, 'transportFeeType': 1, 'callSource': 1, 'optId': 0, 'optName': '系统批量试算',
        'goodsList': [{'goodsCode': task['code'], 'qty': 1, 'isProductOrigin': False}],
    }
    if not hasattr(thread_local, 'session'):
        thread_local.session = requests.Session()
    try:
        response = thread_local.session.post(URL, json=body, timeout=45)
        result = {'task': task, 'http_status': response.status_code, 'response': response.json()}
    except Exception as exc:
        result = {'task': task, 'error': f'{type(exc).__name__}: {exc}'}
    (OUT / f"{task['warehouse_code']}_{task['code']}.json").write_text(
        json.dumps(result, ensure_ascii=False, default=str), encoding='utf-8')
    value = (result.get('response') or {}).get('data') or {}
    return task['warehouse_code'], task['code'], value.get('totalComboCost'), value.get('success'), result.get('error')


with ThreadPoolExecutor(max_workers=8) as pool:
    futures = [pool.submit(calculate, task) for task in pending]
    for index, future in enumerate(as_completed(futures), 1):
        warehouse, code, cost, success, error = future.result()
        if index % 100 == 0 or index == len(pending) or error:
            print(f'{index}/{len(pending)} {warehouse} {code} cost={cost} success={success} error={error}', flush=True)
print('done')
