#!/usr/bin/env python3
import json
import sys
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / '来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx'
OUT = ROOT / '已定价商品_北京履约属性'
OUT.mkdir(exist_ok=True)
from archery_runtime import ARCHERY_DIR
sys.path.insert(0, str(ARCHERY_DIR))
from archery_client import ArcheryClient
from execute_archery_query import execute

wb = load_workbook(SOURCE, read_only=True, data_only=True)
ws = wb['二次定价']
it = ws.iter_rows(values_only=True)
headers = next(it)
products = []
for row in it:
    item = dict(zip(headers, row))
    if item.get('定价状态') == '已定价':
        products.append(str(item['商品编号']))
wb.close()
products = sorted(set(products))
(ROOT / '已定价商品编号.json').write_text(json.dumps(products, ensure_ascii=False), encoding='utf-8')
print(f'已定价商品：{len(products)}', flush=True)

client = ArcheryClient()
if not client.login():
    raise SystemExit('Archery login failed')
instance = '[online]生产PolarDB-Newretail(prod-newretail.rwlb.rds.aliyuncs.com)'
for batch_no, start in enumerate(range(0, len(products), 400)):
    output = OUT / f'batch_{batch_no:04d}.json'
    if output.exists():
        continue
    codes = products[start:start + 400]
    in_sql = ','.join("'" + code.replace("'", "''") + "'" for code in codes)
    sql = f"""SELECT g.id, g.goods_code, g.goods_name, g.warehouse_id, g.owner_id,
       g.gross_weight, g.net_weight, g.length, g.width, g.height, g.vol,
       g.inner_pack_length, g.inner_pack_width, g.inner_pack_height, g.inner_pack_vol,
       g.outer_pack_length, g.outer_pack_width, g.outer_pack_height, g.outer_pack_vol,
       g.expand_length, g.expand_width, g.expand_height, g.expand_vol,
       g.shape, g.attr, g.goods_class, g.logistics_attr, g.state,
       e.storage_method, e.delivery_method, e.is_origin_pack, e.is_mix_pack,
       e.goods_type, e.package_count, e.bas_pack_id
FROM wmsx_goods.bas_goods g
LEFT JOIN wmsx_goods.bas_goods_extend e
  ON g.id = e.goods_id AND g.warehouse_id = e.warehouse_id
 AND g.owner_id = e.owner_id AND e.is_delete = 0
WHERE g.goods_code IN ({in_sql})
  AND g.owner_id = 10776 AND g.warehouse_id = 70 AND g.is_delete = 0
ORDER BY g.goods_code, g.id
LIMIT 1000"""
    result = execute(client, sql, 1000, instance, 'wmsx_goods')
    if result.get('status') != 0 or (result.get('data') or {}).get('error'):
        raise RuntimeError(f'batch {batch_no}: {result}')
    output.write_text(json.dumps(result, ensure_ascii=False, default=str), encoding='utf-8')
    count = len((result.get('data') or {}).get('rows') or [])
    print(f'batch {batch_no}: {count} rows', flush=True)
print('done')
