SELECT *
FROM dbo.Product_Pool
WHERE PoolName LIKE N'%来福%'
   OR PoolName LIKE N'%禄品%'
   OR PoolName LIKE N'%分销%'
   OR PoolName LIKE N'%一件代发%'
ORDER BY SysNo;
