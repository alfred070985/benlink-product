SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = 'dbo'
  AND TABLE_NAME IN (
    'Product_Pool', 'Product_Pool_Item', 'Product_Pool_Rule', 'Product_Pool_Type',
    'PlatformProductCustom', 'PlatformProductCustomItem',
    'Platform_Product_Exclusive', 'Platform_Product_Delivery',
    'Platform_Product', 'Platform_Product_WebSite'
  )
ORDER BY TABLE_NAME, ORDINAL_POSITION;
