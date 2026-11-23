SELECT *
FROM dbo.Product_Pool
WHERE SysNo IN (27, 52)
   OR ParentSysNo = 52
ORDER BY SysNo;
