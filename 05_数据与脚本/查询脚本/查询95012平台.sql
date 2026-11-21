SELECT Sysno, Name, GroupSysNo, Status, Remark, ApiVersion,
       integration_platform_id, traceability_platform_id,
       CreateDate, UpdateDate
FROM dbo.Partner
WHERE Sysno = 95012