SELECT w.id AS warehouse_id, w.warehouse_code, w.warehouse_name,
       o.id AS owner_id, o.owner_code, o.owner_name
FROM wmsx_basic.bas_warehouse w
JOIN wmsx_basic.bas_owner o ON o.id = 10776
WHERE w.id IN (62, 70, 71, 78) AND w.is_delete = 0 AND o.is_delete = 0
ORDER BY w.id
LIMIT 100
