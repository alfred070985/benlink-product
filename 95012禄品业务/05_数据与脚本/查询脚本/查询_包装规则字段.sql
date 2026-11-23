SELECT table_name, column_name, column_type, ordinal_position
FROM information_schema.columns
WHERE table_schema = 'wmsx_basic'
  AND table_name IN (
    'rule_delivery', 'rule_delivery_method', 'rule_delivery_package_category',
    'rule_delivery_material', 'rule_delivery_consumables',
    'bas_material_consumables', 'bas_material_consumables_bind',
    'bas_material_consumables_goods', 'bas_material_consumables_temperature',
    'bas_carrier'
  )
ORDER BY table_name, ordinal_position
LIMIT 1000
