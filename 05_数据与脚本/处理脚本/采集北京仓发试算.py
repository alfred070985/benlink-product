#!/usr/bin/env python3
import argparse
import glob
import json
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import requests
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / '来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx'
OUT = ROOT / '履约试算_北京仓发'
OUT.mkdir(exist_ok=True)
URL = 'https://vm-online.benlai.com/bidding/fulfillmentCostEstimate/calculateOutboundPerformance'
KEYS = ['gross_weight', 'net_weight', 'length', 'width', 'height', 'vol', 'shape', 'attr',
        'storage_method', 'delivery_method', 'is_origin_pack', 'is_mix_pack', 'goods_type', 'package_count']

wb = load_workbook(SOURCE, read_only=True, data_only=True)
ws = wb['二次定价']
it = ws.iter_rows(values_only=True)
headers = next(it)
codes = []
for row in it:
    item = dict(zip(headers, row))
    if item.get('定价状态') == '已定价' and item.get('是否产地直发') == '仓发':
        codes.append(str(item['商品编号']))
wb.close()
code_set = set(codes)
attrs = {}
for filename in glob.glob(str(ROOT / '已定价商品_北京履约属性' / '*.json')):
    data = json.loads(Path(filename).read_text(encoding='utf-8'))['data']
    for values in data.get('rows') or []:
        item = dict(zip(data['column_list'], values))
        if item['goods_code'] in code_set:
            attrs[item['goods_code']] = item

groups = {}
for code, item in attrs.items():
    signature = tuple(item.get(key) for key in KEYS)
    groups.setdefault(signature, []).append(code)
tasks = {}
for members in groups.values():
    members = sorted(members)
    tasks[members[0]] = {'code': members[0], 'group_size': len(members), 'validation_for': None}
    if len(members) > 1:
        tasks[members[1]] = {'code': members[1], 'group_size': len(members),
                             'validation_for': members[0]}
for code in sorted(code_set - set(attrs)):
    tasks[code] = {'code': code, 'group_size': 1, 'validation_for': None}

# 只保留每组一个正式样本，另加每组一个复核样本（若存在）。
parser = argparse.ArgumentParser()
parser.add_argument('--max-tasks', type=int, default=900)
args = parser.parse_args()
pending = [task for code, task in sorted(tasks.items()) if not (OUT / f'{code}.json').exists()]
pending = pending[:args.max_tasks]
print(f'仓发商品={len(codes)}, 属性商品={len(attrs)}, 属性分组={len(groups)}, 总任务={len(tasks)}, 本次={len(pending)}', flush=True)
thread_local = threading.local()

def calculate(task):
    code = task['code']
    body = {
        'warehouseCode': 'W005', 'businessId': 5, 'orderChannelId': 5,
        'provinceCode': '1', 'cityCode': '2', 'districtCode': '3',
        'logisticsCompanyCode': 'jingdong', 'comboQty': 1, 'transportFeeType': 1,
        'callSource': 1, 'optId': 0, 'optName': '系统批量试算',
        'goodsList': [{'goodsCode': code, 'qty': 1, 'isProductOrigin': False}],
    }
    if not hasattr(thread_local, 'session'):
        thread_local.session = requests.Session()
    try:
        response = thread_local.session.post(URL, json=body, timeout=60)
        result = {'task': task, 'http_status': response.status_code, 'response': response.json()}
    except Exception as exc:
        result = {'task': task, 'error': f'{type(exc).__name__}: {exc}'}
    (OUT / f'{code}.json').write_text(json.dumps(result, ensure_ascii=False, default=str), encoding='utf-8')
    value = (result.get('response') or {}).get('data') or {}
    return code, value.get('totalComboCost'), value.get('success'), result.get('error')

with ThreadPoolExecutor(max_workers=8) as pool:
    futures = [pool.submit(calculate, task) for task in pending]
    for index, future in enumerate(as_completed(futures), 1):
        code, cost, success, error = future.result()
        if index % 50 == 0 or index == len(pending) or error:
            print(f'{index}/{len(pending)} {code} cost={cost} success={success} error={error}', flush=True)
print('done')
