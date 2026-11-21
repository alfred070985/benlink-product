SELECT table_schema, table_name, column_name, column_type, ordinal_position
FROM information_schema.columns
WHERE (table_schema = 'wmsx_goods' AND table_name IN ('bas_goods', 'bas_goods_extend', 'bas_goods_consumables'))
   OR (table_schema = 'wmsx_basic' AND table_name IN (
      'bas_material_consumables', 'bas_material_consumables_bind',
      'bas_material_consumables_goods', 'bas_material_consumables_temperature',
      'rule_delivery_material'
   ))
ORDER BY table_schema, table_name, ordinal_position
LIMIT 1000
