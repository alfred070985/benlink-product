SELECT table_name, column_name, column_type, ordinal_position
FROM information_schema.columns
WHERE table_schema = 'wmsx_basic'
  AND table_name IN ('bas_goods', 'bas_warehouse', 'bas_owner', 'bas_carrier', 'bas_delivery_area')
ORDER BY table_name, ordinal_position
LIMIT 1000
