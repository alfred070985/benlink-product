SELECT id, owner_id, name, type, weight, length, width, height,
       temperature, state, refrigerant_type
FROM wmsx_basic.bas_material_consumables
WHERE warehouse_id = 70 AND is_delete = 0
ORDER BY type, id
LIMIT 500
