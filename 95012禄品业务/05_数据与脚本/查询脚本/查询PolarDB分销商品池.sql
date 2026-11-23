SELECT SysNo, PoolName, BusinessType, PoolType, IsShow, SkuCount, Status, IsDelete,
       CreateTime, LastUpdateTime, PoolAttribute, PoolScope, ParentSysNo,
       IsUseCustomerPool, MaintainSysNo
FROM product_pool
WHERE PoolName LIKE '%分销%'
   OR PoolName LIKE '%禄品%'
   OR PoolAttribute = 1
   OR SysNo IN (27, 52)
ORDER BY SysNo;
