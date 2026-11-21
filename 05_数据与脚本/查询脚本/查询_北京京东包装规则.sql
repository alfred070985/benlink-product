SELECT c.id AS carrier_id, c.carrier_name, c.logistics_company_code,
       r.id AS rule_id, r.rule_name, r.state AS rule_state, r.outer_temperature,
       m.delivery_method,
       pc.id AS rule_package_category_id, pc.package_category_id, pc.package_category_name,
       rm.id AS rule_material_id, rm.material_id, rm.material_code, rm.material_name,
       rc.consumables_code, rc.consumables_name, rc.qty
FROM wmsx_basic.bas_carrier c
JOIN wmsx_basic.rule_delivery r ON r.carrier_id = c.id AND r.is_delete = 0
LEFT JOIN wmsx_basic.rule_delivery_method m ON m.rule_id = r.id AND m.is_delete = 0
LEFT JOIN wmsx_basic.rule_delivery_package_category pc ON pc.rule_id = r.id AND pc.is_delete = 0
LEFT JOIN wmsx_basic.rule_delivery_material rm
  ON rm.rule_id = r.id AND rm.rule_package_category_id = pc.id AND rm.is_delete = 0
LEFT JOIN wmsx_basic.rule_delivery_consumables rc
  ON rc.rule_id = r.id AND rc.rule_package_category_id = pc.id
 AND rc.rule_material_id = rm.id AND rc.is_delete = 0
WHERE c.warehouse_id = 70 AND c.owner_id = 10776
  AND c.logistics_company_code = 'jingdong' AND c.is_delete = 0
ORDER BY c.id, r.id, pc.id, rm.id, rc.id
LIMIT 1000
