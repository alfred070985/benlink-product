SELECT SysNo, PartnerSysNo, Status, Remark, JsonContent, CreateDate, UpdateDate
FROM dbo.Partner_JsonConfig
WHERE PartnerSysNo = 95012
ORDER BY SysNo