SELECT PoolSysNo, Status, IsDelete, COUNT_BIG(1) AS ItemRows,
       COUNT_BIG(DISTINCT ProductSysNo) AS ProductCount,
       COUNT_BIG(DISTINCT ProductId) AS ProductIdCount,
       MIN(CreateTime) AS FirstCreateTime, MAX(LastUpdateTime) AS LastUpdateTime
FROM dbo.Product_Pool_Item
WHERE PoolSysNo = 27
GROUP BY PoolSysNo, Status, IsDelete
ORDER BY Status, IsDelete;
