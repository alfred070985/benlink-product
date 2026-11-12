SELECT 'StorePlatform' AS Source, SysNo, Name, Status, CONVERT(int, IsDelete) AS IsDelete
FROM dbo.StorePlatform
WHERE Name LIKE '%禄品%'
UNION ALL
SELECT 'Partner', Sysno, Name, Status, 0
FROM dbo.Partner
WHERE Name LIKE '%禄品%'
ORDER BY Source, SysNo