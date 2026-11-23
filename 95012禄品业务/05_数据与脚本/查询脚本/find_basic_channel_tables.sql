SELECT TABLE_NAME,
       COUNT(DISTINCT COLUMN_NAME) AS matched_columns,
       MAX(CASE WHEN COLUMN_NAME = 'ProductBasicSysNo' THEN 1 ELSE 0 END) AS has_basic,
       MAX(CASE WHEN COLUMN_NAME = 'WebSiteSysNo' THEN 1 ELSE 0 END) AS has_website,
       MAX(CASE WHEN COLUMN_NAME = 'ChannelSysNo' THEN 1 ELSE 0 END) AS has_channel,
       MAX(CASE WHEN COLUMN_NAME = 'SaleChannelSysNo' THEN 1 ELSE 0 END) AS has_sale_channel
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = 'dbo'
  AND COLUMN_NAME IN ('ProductBasicSysNo', 'WebSiteSysNo', 'ChannelSysNo', 'SaleChannelSysNo')
GROUP BY TABLE_NAME
HAVING MAX(CASE WHEN COLUMN_NAME = 'ProductBasicSysNo' THEN 1 ELSE 0 END) = 1
ORDER BY matched_columns DESC, TABLE_NAME
