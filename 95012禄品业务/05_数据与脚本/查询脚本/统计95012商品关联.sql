SELECT 'Platform_Product' AS Source, Status, CONVERT(int, IsDelete) AS IsDelete,
       COUNT(*) AS Rows, COUNT(DISTINCT ProductSysNo) AS Products
FROM dbo.Platform_Product
WHERE PlatformSysNo = 95012
GROUP BY Status, CONVERT(int, IsDelete)
UNION ALL
SELECT 'Platform_Product_WebSite', Status, CONVERT(int, IsDelete),
       COUNT(*), COUNT(DISTINCT ProductSysNo)
FROM dbo.Platform_Product_WebSite
WHERE PlatformSysNo = 95012
GROUP BY Status, CONVERT(int, IsDelete)
UNION ALL
SELECT 'Product_Pool_Item_Channel', Status, IsDelete,
       COUNT(*), COUNT(DISTINCT ProductSysNo)
FROM dbo.Product_Pool_Item
WHERE ChannelSysNo = 95012
GROUP BY Status, IsDelete
UNION ALL
SELECT 'Product_Pool_Item_SaleChannel', Status, IsDelete,
       COUNT(*), COUNT(DISTINCT ProductSysNo)
FROM dbo.Product_Pool_Item
WHERE SaleChannelSysNo = 95012
GROUP BY Status, IsDelete
ORDER BY Source, Status, IsDelete