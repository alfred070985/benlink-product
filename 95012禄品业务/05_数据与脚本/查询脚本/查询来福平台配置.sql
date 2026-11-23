SELECT SysNo, GroupSysNo, PlatformName, ChargeSysNo, CustomerSysNo, CustomerName,
       Status, IsDelete, CreateTime, ApiVersion, SourceSysNo, CompanySysNo,
       ProductPoolType, DockingType, IsPushProduct
FROM dbo.Platform
WHERE PlatformName LIKE N'%来福%'
   OR CustomerName LIKE N'%来福%'
   OR Remark LIKE N'%来福%'
ORDER BY SysNo;
