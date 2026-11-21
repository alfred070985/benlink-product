#!/usr/bin/env python3
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
from archery_runtime import ARCHERY_DIR
sys.path.insert(0, str(ARCHERY_DIR))
from archery_client import ArcheryClient
from execute_archery_query import execute

INSTANCE = '[online]生产PolarDB-Newretail(prod-newretail.rwlb.rds.aliyuncs.com)'
TABLES = [
    'billing_business', 'billing_business_rel', 'billing_business_rel_config_type',
    'billing_config', 'billing_config_item', 'billing_config_item_param', 'billing_config_item_sub',
    'billing_rule', 'billing_rule_delivery', 'billing_rule_delivery_carrier',
    'billing_rule_delivery_interval', 'billing_rule_delivery_weight',
    'billing_rule_match', 'billing_rule_match_condition', 'billing_rule_param',
    'billing_template', 'billing_template_param', 'billing_param',
    'billing_round_rule', 'billing_round_rule_item',
]
OUT = ROOT / '履约规则数据'
OUT.mkdir(exist_ok=True)
client = ArcheryClient()
if not client.login():
    raise SystemExit('Archery login failed')
for table in TABLES:
    if (OUT / f'{table}.json').exists():
        print(f'{table}: 已存在，跳过', flush=True)
        continue
    rows = []
    columns = None
    for offset in range(0, 100000, 1000):
        sql = f'SELECT * FROM {table} ORDER BY id LIMIT {offset}, 1000'
        result = execute(client, sql, 1000, INSTANCE, 'wmsx_billing')
        if result.get('status') != 0:
            raise RuntimeError(f'{table}: {result}')
        data = result.get('data') or {}
        if data.get('error'):
            raise RuntimeError(f'{table}: {data["error"]}')
        columns = data.get('column_list') or columns
        batch = data.get('rows') or []
        rows.extend(batch)
        print(f'{table}: {len(rows)} rows', flush=True)
        if len(batch) < 1000:
            break
        time.sleep(0.1)
    (OUT / f'{table}.json').write_text(json.dumps({
        'table': table, 'columns': columns, 'rows': rows,
    }, ensure_ascii=False, default=str), encoding='utf-8')
print('done')
