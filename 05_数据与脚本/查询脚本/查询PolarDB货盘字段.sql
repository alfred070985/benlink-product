SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE
FROM information_schema.COLUMNS
WHERE TABLE_SCHEMA = 'benlai_business_product'
  AND TABLE_NAME IN (
    'platform_product', 'platform_product_exclusive', 'platform_product_website',
    'product_pool', 'product_pool_item', 'product_pool_rule',
    'product_pool_selection_rule', 'product_pool_type',
    'product_pool_whitelist', 'product_pool_black'
  )
ORDER BY TABLE_NAME, ORDINAL_POSITION;
