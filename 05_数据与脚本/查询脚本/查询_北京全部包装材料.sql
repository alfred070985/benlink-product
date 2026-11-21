SELECT DISTINCT rm.owner_id, rm.material_code, rm.material_name, g.id AS material_goods_id,
       g.gross_weight, g.net_weight, g.length, g.width, g.height, g.vol,
       g.state AS goods_state, g.is_delete AS goods_deleted
FROM wmsx_basic.rule_delivery_material rm
LEFT JOIN wmsx_goods.bas_goods g
  ON g.id = rm.material_id AND g.warehouse_id = rm.warehouse_id
WHERE rm.warehouse_id = 70 AND rm.is_delete = 0
ORDER BY rm.owner_id, rm.material_code
LIMIT 1000
