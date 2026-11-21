SELECT Sysno, PartnerSysno, BusinessSysno, EnableStatus, Status, Remark,
       CreateDate, UpdateDate
FROM dbo.Partner_Business_Relation
WHERE PartnerSysno = 95012
ORDER BY Sysno