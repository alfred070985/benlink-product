#!/usr/bin/env python3
import json
from pathlib import Path

P = Path(__file__).resolve().parent / '履约规则数据'
def rows(name):
    d = json.loads((P / f'{name}.json').read_text(encoding='utf-8'))
    cols = d.get('columns') or (d.get('data') or {}).get('column_list')
    values = d.get('rows') or (d.get('data') or {}).get('rows')
    return [dict(zip(cols, row)) for row in values]

params = {x['id']: x for x in rows('billing_param')}
rules = {x['id']: x for x in rows('billing_rule')}
templates = {x['id']: x for x in rows('billing_template')}
conditions = rows('billing_rule_match_condition')
matches = rows('billing_rule_match')
for warehouse_id in (70, 71, 62, 78):
    print('\n' + '=' * 30, '仓库', warehouse_id)
    for rule in rules.values():
        if rule['warehouse_id'] != warehouse_id or rule['owner_id'] != 10776 or rule['is_delete'] or not rule['state']:
            continue
        if rule['billing_type'] not in (12, 13):
            continue
        print('\n规则', rule['id'], rule['rule_name'], '类型', rule['billing_type'])
        for m in matches:
            if m['billing_rule_id'] == rule['id'] and not m['is_delete']:
                param = params.get(m['billing_param_id'], {})
                print('  匹配', m['line_no'], param.get('param_name'), param.get('param_code'), m['param_value'])
        for c in conditions:
            if c['billing_rule_id'] == rule['id'] and not c['is_delete']:
                param = params.get(c['billing_param_id'], {})
                print('  条件', c['line_no'], param.get('param_name'), param.get('param_code'), c['comp_condition'], c['param_value'])
print('\n' + '=' * 30, '当前北京配置75')
items = [x for x in rows('billing_config_item') if x['billing_config_id'] == 75 and not x['is_delete']]
subs = rows('billing_config_item_sub')
item_params = rows('billing_config_item_param')
for item in items:
    print('\n项目', item)
    for sub in subs:
        if sub['billing_config_item_id'] != item['id'] or sub['is_delete']:
            continue
        print('  子项', sub)
        template = templates.get(sub['billing_template_id'])
        if template:
            print('  模板', template['template_name'], template['template_code'], template['template_content'])
        for p in item_params:
            if p['billing_config_item_sub_id'] == sub['id'] and not p['is_delete']:
                param = params.get(p['billing_param_id'], {})
                print('    参数', param.get('param_name'), param.get('param_code'), p['param_value'])
