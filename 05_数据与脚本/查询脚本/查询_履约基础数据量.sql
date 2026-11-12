SELECT 'wmsx_goods.bas_goods' AS table_name, COUNT(*) AS row_count
FROM wmsx_goods.bas_goods
WHERE owner_id = 10776 AND warehouse_id IN (62, 70, 71, 78) AND is_delete = 0
UNION ALL
SELECT 'wmsx_goods.bas_goods_extend', COUNT(*)
FROM wmsx_goods.bas_goods_extend
WHERE owner_id = 10776 AND warehouse_id IN (62, 70, 71, 78) AND is_delete = 0
UNION ALL
SELECT 'wmsx_basic.bas_material_consumables', COUNT(*)
FROM wmsx_basic.bas_material_consumables
WHERE owner_id = 10776 AND warehouse_id IN (62, 70, 71, 78) AND is_delete = 0
UNION ALL
SELECT 'wmsx_basic.bas_material_consumables_bind', COUNT(*)
FROM wmsx_basic.bas_material_consumables_bind
WHERE owner_id = 10776 AND warehouse_id IN (62, 70, 71, 78) AND is_delete = 0
UNION ALL
SELECT 'wmsx_basic.rule_delivery_material', COUNT(*)
FROM wmsx_basic.rule_delivery_material
WHERE owner_id = 10776 AND warehouse_id IN (62, 70, 71, 78) AND is_delete = 0
UNION ALL
SELECT 'wmsx_basic.bas_carrier', COUNT(*)
FROM wmsx_basic.bas_carrier
WHERE owner_id = 10776 AND warehouse_id IN (62, 70, 71, 78) AND is_delete = 0
LIMIT 100
