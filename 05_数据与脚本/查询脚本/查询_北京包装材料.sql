SELECT DISTINCT rm.material_code, rm.material_name, g.id AS material_goods_id,
       g.gross_weight, g.net_weight, g.length, g.width, g.height, g.vol,
       g.state AS goods_state, g.is_delete AS goods_deleted
FROM wmsx_basic.rule_delivery_material rm
LEFT JOIN wmsx_goods.bas_goods g
  ON g.id = rm.material_id AND g.warehouse_id = rm.warehouse_id AND g.owner_id = rm.owner_id
WHERE rm.warehouse_id = 70 AND rm.owner_id = 10776 AND rm.is_delete = 0
ORDER BY rm.material_code
LIMIT 1000
