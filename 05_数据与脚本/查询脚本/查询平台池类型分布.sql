SELECT ProductPoolType, Status, IsDelete, COUNT_BIG(1) AS PlatformCount,
       MIN(SysNo) AS MinPlatformSysNo, MAX(SysNo) AS MaxPlatformSysNo
FROM dbo.Platform
GROUP BY ProductPoolType, Status, IsDelete
ORDER BY ProductPoolType, Status, IsDelete;
