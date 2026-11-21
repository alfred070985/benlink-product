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
OUT = ROOT / '履约试算_上海全量'
OUT.mkdir(exist_ok=True)
URL = 'https://vm-online.benlai.com/bidding/fulfillmentCostEstimate/calculateOutboundPerformance'

wb = load_workbook(SOURCE, read_only=True, data_only=True)
ws = wb['二次定价']
it = ws.iter_rows(values_only=True)
headers = next(it)
all_codes = []
warehouse_codes = []
for values in it:
    item = dict(zip(headers, values))
    code = str(item['商品编号'])
    all_codes.append(code)
    if item.get('是否产地直发') == '仓发':
        warehouse_codes.append(code)
wb.close()
assert len(all_codes) == len(set(all_codes)), '报价表商品编号存在重复'
warehouse_codes = sorted(warehouse_codes)
region_data = json.loads((ROOT / '履约试算_区域_上海.json').read_text(encoding='utf-8'))['data']
areas = [(area, area['provinceList'][0]) for area in region_data['area']]
assert len(areas) == 6


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
for code in warehouse_codes:
    for area, province in areas:
        output = OUT / f"{code}_{area['areaId']}.json"
        if not resolved(output):
            tasks.append({'code': code, 'area_id': area['areaId'], 'area_name': area['areaName'], 'province': province})
parser = argparse.ArgumentParser()
parser.add_argument('--max-tasks', type=int, default=2500)
args = parser.parse_args()
pending = tasks[:args.max_tasks]
print(f'报价表商品={len(all_codes)}，仓发商品={len(warehouse_codes)}，总任务={len(warehouse_codes) * 6}，'
      f'待处理={len(tasks)}，本次={len(pending)}', flush=True)
thread_local = threading.local()


def calculate(task):
    province = task['province']
    body = {
        'warehouseCode': 'W028', 'businessId': 5, 'orderChannelId': 5,
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
    (OUT / f"{task['code']}_{task['area_id']}.json").write_text(
        json.dumps(result, ensure_ascii=False, default=str), encoding='utf-8')
    data = (result.get('response') or {}).get('data') or {}
    return task['code'], task['area_id'], data.get('success'), data.get('totalComboCost'), result.get('error')


with ThreadPoolExecutor(max_workers=12) as pool:
    futures = [pool.submit(calculate, task) for task in pending]
    for index, future in enumerate(as_completed(futures), 1):
        code, area_id, success, cost, error = future.result()
        if index % 200 == 0 or index == len(pending) or error:
            print(f'{index}/{len(pending)} {code} 区域{area_id} success={success} cost={cost} error={error}', flush=True)
print('done')
