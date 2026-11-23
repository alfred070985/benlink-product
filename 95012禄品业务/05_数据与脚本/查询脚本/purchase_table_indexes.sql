SELECT TABLE_NAME, INDEX_NAME, SEQ_IN_INDEX, COLUMN_NAME
FROM information_schema.STATISTICS
WHERE TABLE_SCHEMA = 'newretail_oms'
  AND TABLE_NAME IN ('callback_wmsx_order', 'callback_wmsx_order_detail')
ORDER BY TABLE_NAME, INDEX_NAME, SEQ_IN_INDEX
