SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = 'dbo'
  AND TABLE_NAME IN ('Product', 'Product_Price', 'Product_Basic', 'ProductBasic')
  AND COLUMN_NAME IN ('SysNo', 'ProductSysNo', 'ProductBasicSysNo', 'ProductID', 'WebSiteSysNo', 'ChannelSysNo', 'SaleChannelSysNo', 'IsDelete', 'Status', 'UnitCost')
ORDER BY TABLE_NAME, COLUMN_NAME
