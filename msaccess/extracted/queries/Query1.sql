-- Query Name: Query1
-- Extracted: 2026-01-29 16:09:05

SELECT temptbl6.EMPID_I, temptbl6.SKILLNUMBER_I, temptbl6.Skill, temptbl6.EXPIREDSKILL_I, temptbl6.LASTNAME, temptbl6.FRSTNAME, [temptbl6].[FRSTNAME] & ' ' & [temptbl6].[LASTNAME] AS StaffName, temptbl6.LocationName, temptbl6.SUPERVISORCODE_I, temptbl6.SupervisorIndexedName, temptbl6.SupervisorName, temptbl6.RedFlag, temptbl6.OnLeave INTO temptbl8
FROM temptbl6
WHERE (((temptbl6.RedFlag) = True));

