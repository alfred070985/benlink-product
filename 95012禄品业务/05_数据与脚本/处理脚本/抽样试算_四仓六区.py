#!/usr/bin/env python3
import json
import time
from pathlib import Path
import requests

ROOT = Path(__file__).resolve().parent
URL = 'https://vm-online.benlai.com/bidding/fulfillmentCostEstimate/calculateOutboundPerformance'
WAREHOUSES = [
    (1, '北京', 'W005', '履约试算_区域_北京.json'),
    (2, '广州', 'W006', '履约试算_区域_广州.json'),
    (3, '上海', 'W028', '履约试算_区域_上海.json'),
    (7, '武汉', 'W010', '履约试算_区域_武汉.json'),
]
results = []
for site_id, site_name, warehouse_code, filename in WAREHOUSES:
    payload = json.loads((ROOT / filename).read_text(encoding='utf-8'))
    data = payload.get('data') or payload.get('value') or {}
    for area in data.get('area') or []:
        province = area['provinceList'][0]
        body = {
            'warehouseCode': warehouse_code,
            'businessId': 5,
            'orderChannelId': 5,
            'provinceCode': str(province['provinceCode']),
            'cityCode': str(province['cityCode']),
            'districtCode': str(province['districtCode']),
            'logisticsCompanyCode': 'jingdong',
            'comboQty': 1,
            'transportFeeType': 1,
            'callSource': 1,
            'optId': 0,
            'optName': '系统批量试算',
            'goodsList': [{'goodsCode': 'P0000534174', 'qty': 1, 'isProductOrigin': False}],
        }
        response = requests.post(URL, json=body, timeout=60)
        response.raise_for_status()
        result = response.json()
        value = result.get('data') or result.get('value') or {}
        row = {
            'siteId': site_id, 'siteName': site_name, 'warehouseCode': warehouse_code,
            'areaId': area['areaId'], 'areaName': area['areaName'],
            'province': province, 'response': result,
        }
        results.append(row)
        print(site_name, area['areaName'], result.get('code'), value.get('success'),
              value.get('totalComboCost'), value.get('errorMsg') or result.get('msg'), flush=True)
        time.sleep(0.1)
(ROOT / '抽样试算_四仓六区.json').write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
