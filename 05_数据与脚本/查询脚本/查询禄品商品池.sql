SELECT *
FROM dbo.Product_Pool
WHERE PoolName LIKE N'%禄品%'
   OR SysNo = 95012
   OR PoolType = 1
ORDER BY SysNo;
