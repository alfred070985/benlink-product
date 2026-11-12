SELECT PoolType, Status, IsDelete, COUNT_BIG(1) AS ItemRows,
       COUNT_BIG(DISTINCT ProductSysNo) AS ProductCount,
       COUNT_BIG(DISTINCT ProductId) AS ProductIdCount
FROM dbo.Product_Pool_Item
GROUP BY PoolType, Status, IsDelete
ORDER BY PoolType, Status, IsDelete;
