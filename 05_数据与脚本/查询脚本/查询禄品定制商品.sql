SELECT c.SysNo AS CustomSysNo, c.EntSysNo, c.StoreSysNo, c.Status AS CustomStatus,
       c.IsDelete AS CustomIsDelete, c.Code, c.RecommendName,
       i.SysNo AS ItemSysNo, i.ProductId, i.Status AS ItemStatus, i.IsDelete AS ItemIsDelete, i.Sort
FROM dbo.PlatformProductCustom AS c
LEFT JOIN dbo.PlatformProductCustomItem AS i ON i.RecommendSysNo = c.SysNo
WHERE c.EntSysNo = 95012
ORDER BY c.SysNo, i.Sort, i.SysNo;
