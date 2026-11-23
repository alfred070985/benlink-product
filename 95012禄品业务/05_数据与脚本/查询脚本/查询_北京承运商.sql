SELECT id, carrier_name, logistics_company_code, logistics_company_name,
       carrier_type, express_type, is_default, state, sort_num
FROM wmsx_basic.bas_carrier
WHERE warehouse_id = 70 AND owner_id = 10776 AND is_delete = 0
ORDER BY sort_num, id
LIMIT 200
