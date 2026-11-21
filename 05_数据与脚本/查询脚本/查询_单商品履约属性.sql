SELECT g.id, g.goods_code, g.goods_name, g.warehouse_id, g.owner_id,
       g.gross_weight, g.net_weight, g.length, g.width, g.height, g.vol,
       g.shape, g.attr, g.goods_class, g.logistics_attr, g.state,
       e.storage_method, e.delivery_method, e.is_origin_pack, e.is_mix_pack,
       e.goods_type, e.package_count, e.bas_pack_id
FROM wmsx_goods.bas_goods g
LEFT JOIN wmsx_goods.bas_goods_extend e
  ON g.id = e.goods_id AND g.warehouse_id = e.warehouse_id
 AND g.owner_id = e.owner_id AND e.is_delete = 0
WHERE g.goods_code = 'P0000534174'
  AND g.owner_id = 10776
  AND g.warehouse_id IN (62, 70, 71, 78)
  AND g.is_delete = 0
ORDER BY g.warehouse_id, g.id
LIMIT 100
