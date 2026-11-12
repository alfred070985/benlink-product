SELECT *
FROM dbo.Product_Pool
WHERE IsDelete = 0 AND PoolAttribute = 1
ORDER BY PoolType, SysNo;
