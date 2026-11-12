SELECT 'Product.SourceChannelSysNo' AS Source, COUNT(*) AS Rows, COUNT(DISTINCT SysNo) AS Products
FROM dbo.Product
WHERE SourceChannelSysNo = 95012
UNION ALL
SELECT 'Product_Pool.SysNo', COUNT(*), COUNT(DISTINCT SysNo)
FROM dbo.Product_Pool
WHERE SysNo = 95012
UNION ALL
SELECT 'Product_PoolItem.PoolSysNo', COUNT(*), COUNT(DISTINCT ProductSysNo)
FROM dbo.Product_Pool_Item
WHERE PoolSysNo = 95012
UNION ALL
SELECT 'PlatformProductCustom.PlatformSysNo', COUNT(*), COUNT(DISTINCT SysNo)
FROM dbo.PlatformProductCustom
WHERE SysNo = 95012