#!/usr/bin/env python3
import glob
import json
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / '来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx'
OUT = ROOT / '履约试算_北京仓发'
ATTR = ROOT / '已定价商品_北京履约属性'
KEYS = ['gross_weight', 'net_weight', 'length', 'width', 'height', 'vol', 'shape', 'attr',
        'storage_method', 'delivery_method', 'is_origin_pack', 'is_mix_pack', 'goods_type', 'package_count']

wb = load_workbook(SOURCE, read_only=True, data_only=True)
ws = wb['二次定价']
it = ws.iter_rows(values_only=True)
headers = next(it)
codes = sorted({str(row[headers.index('商品编号')]) for row in it
                if row[headers.index('定价状态')] == '已定价' and row[headers.index('是否产地直发')] == '仓发'})
wb.close()
code_set = set(codes)
attrs = {}
for filename in glob.glob(str(ATTR / '*.json')):
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
        tasks[members[1]] = {'code': members[1], 'group_size': len(members), 'validation_for': members[0]}
for code in sorted(code_set - set(attrs)):
    tasks[code] = {'code': code, 'group_size': 1, 'validation_for': None}


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

results = {}
for path in OUT.glob('*.json'):
    result = json.loads(path.read_text(encoding='utf-8'))
    results[result['task']['code']] = result
statuses = Counter(classify(result) for result in results.values())
missing = sorted(set(tasks) - set(results))
extra = sorted(set(results) - set(tasks))
costs = Counter()
for result in results.values():
    if classify(result) == '成功':
        value = ((result.get('response') or {}).get('data') or {}).get('totalComboCost')
        costs[str(value)] += 1

mismatches = []
for code, task in tasks.items():
    representative = task.get('validation_for')
    if not representative or code not in results or representative not in results:
        continue
    left = ((results[representative].get('response') or {}).get('data') or {})
    right = ((results[code].get('response') or {}).get('data') or {})
    left_key = (classify(results[representative]), left.get('totalComboCost'), left.get('errorMsg'))
    right_key = (classify(results[code]), right.get('totalComboCost'), right.get('errorMsg'))
    if left_key != right_key:
        mismatches.append({'representative': representative, 'validation': code,
                           'representative_result': left_key, 'validation_result': right_key})

unresolved = []
for code, result in sorted(results.items()):
    status = classify(result)
    if status not in ('成功', '未查询到商品信息'):
        unresolved.append({'code': code, 'status': status, 'result': result})
summary = {
    '仓发商品数': len(codes), '有属性商品数': len(attrs), '缺属性商品数': len(code_set - set(attrs)),
    '属性分组数': len(groups), '预期任务数': len(tasks), '实际结果数': len(results),
    '状态统计': dict(statuses), '成功成本分布': dict(costs),
    '缺失任务': missing, '额外任务': extra,
    '同属性复核样本数': sum(1 for task in tasks.values() if task.get('validation_for')),
    '同属性复核不一致数': len(mismatches), '同属性复核不一致': mismatches,
    '未解决异常数': len(unresolved), '未解决异常': unresolved,
}
(ROOT / '北京仓发试算_核对.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
print(json.dumps({k: v for k, v in summary.items() if k not in ('同属性复核不一致', '未解决异常')}, ensure_ascii=False, indent=2))
if mismatches:
    print('同属性复核不一致：', json.dumps(mismatches, ensure_ascii=False, indent=2))
if unresolved:
    print('未解决异常：', json.dumps(unresolved, ensure_ascii=False, indent=2))
