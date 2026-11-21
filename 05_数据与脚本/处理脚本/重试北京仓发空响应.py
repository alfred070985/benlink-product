#!/usr/bin/env python3
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
import requests

ROOT = Path(__file__).resolve().parent
OUT = ROOT / '履约试算_北京仓发'
URL = 'https://vm-online.benlai.com/bidding/fulfillmentCostEstimate/calculateOutboundPerformance'
files = []
for path in OUT.glob('*.json'):
    data = json.loads(path.read_text(encoding='utf-8'))
    response = data.get('response') or {}
    if response.get('code') == 0 and response.get('data') is None:
        files.append((path, data))
print(f'空响应重试：{len(files)}', flush=True)

def retry(pair):
    path, data = pair
    code = data['task']['code']
    body = {
        'warehouseCode': 'W005', 'businessId': 5, 'orderChannelId': 5,
        'provinceCode': '1', 'cityCode': '2', 'districtCode': '3',
        'logisticsCompanyCode': 'jingdong', 'comboQty': 1, 'transportFeeType': 1,
        'callSource': 1, 'optId': 0, 'optName': '系统批量试算',
        'goodsList': [{'goodsCode': code, 'qty': 1, 'isProductOrigin': False}],
    }
    try:
        response = requests.post(URL, json=body, timeout=60)
        result = {'task': data['task'], 'http_status': response.status_code, 'response': response.json()}
    except Exception as exc:
        result = {'task': data['task'], 'error': f'{type(exc).__name__}: {exc}'}
    path.write_text(json.dumps(result, ensure_ascii=False, default=str), encoding='utf-8')
    value = (result.get('response') or {}).get('data') or {}
    return code, value.get('success'), value.get('totalComboCost'), value.get('errorMsg')

with ThreadPoolExecutor(max_workers=4) as pool:
    for future in as_completed([pool.submit(retry, pair) for pair in files]):
        print(future.result(), flush=True)
