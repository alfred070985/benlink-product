SELECT m.id, m.billing_rule_id, m.line_no, m.param_value,
       c.param_value AS material_code, r.rule_name, r.state, r.is_delete AS rule_deleted
FROM billing_rule_match m
JOIN billing_rule_match_condition c
  ON m.billing_rule_id = c.billing_rule_id AND m.line_no = c.line_no
 AND c.billing_param_id = 23 AND c.is_delete = 0
JOIN billing_rule r ON r.id = m.billing_rule_id
WHERE m.billing_rule_id = 3 AND m.billing_param_id = 27
  AND c.param_value = 'P0000269326' AND m.is_delete = 0
LIMIT 100
