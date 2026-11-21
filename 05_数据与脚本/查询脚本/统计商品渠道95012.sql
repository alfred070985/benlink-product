SELECT 'ChannelSysNo' AS Source, Status, IsDelete,
       COUNT(*) AS RowCount, COUNT(DISTINCT ProductID) AS Products
FROM product_channel_info
WHERE ChannelSysNo = 95012
GROUP BY Status, IsDelete
UNION ALL
SELECT 'SaleChannelSysNo', Status, IsDelete,
       COUNT(*), COUNT(DISTINCT ProductID)
FROM product_channel_info
WHERE SaleChannelSysNo = 95012
GROUP BY Status, IsDelete
ORDER BY Source, Status, IsDelete