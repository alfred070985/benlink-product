SELECT *
FROM dbo.Platform
WHERE PlatformName LIKE N'%禄品%'
   OR CustomerName LIKE N'%禄品%'
   OR Remark LIKE N'%禄品%'
   OR ChargeSysNo = 95012
   OR SysNo = 95012
   OR CustomerSysNo = 95012
ORDER BY SysNo;
