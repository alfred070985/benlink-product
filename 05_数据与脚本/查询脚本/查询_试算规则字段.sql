SELECT table_name, column_name, column_type, ordinal_position
FROM information_schema.columns
WHERE table_schema = 'wmsx_billing'
  AND table_name IN (
    'billing_business', 'billing_business_rel', 'billing_business_rel_config_type',
    'billing_config', 'billing_config_item', 'billing_config_item_param', 'billing_config_item_sub',
    'billing_rule', 'billing_rule_delivery', 'billing_rule_delivery_area',
    'billing_rule_delivery_carrier', 'billing_rule_delivery_interval', 'billing_rule_delivery_weight',
    'billing_rule_match', 'billing_rule_match_condition', 'billing_rule_param',
    'billing_template', 'billing_template_param', 'billing_round_rule', 'billing_round_rule_item'
  )
ORDER BY table_name, ordinal_position
LIMIT 1000
