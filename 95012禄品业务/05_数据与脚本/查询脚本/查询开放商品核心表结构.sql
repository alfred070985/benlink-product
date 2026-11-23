SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = 'dbo'
  AND TABLE_NAME IN (
    'Platform_Product', 'Platform_Product_WebSite', 'Product_Pool',
    'Product_Pool_Item', 'Product', 'Product_WebSite',
    'Product_Inventory', 'Product_PlatformPrice_Audit'
  )
ORDER BY TABLE_NAME, ORDINAL_POSITION