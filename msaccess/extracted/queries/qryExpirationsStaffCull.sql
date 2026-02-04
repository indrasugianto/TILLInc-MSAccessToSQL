-- Query Name: qryExpirationsStaffCull
-- Extracted: 2026-02-04 13:04:22

SELECT tempstaff.DEPRTMNT, tempstaff.LASTNAME, tempstaff.FRSTNAME, tempstaff.JOBTITLE, temptbl2.SkillDesc, temptbl2.EXPIREDSKILL_I
FROM tempstaff INNER JOIN temptbl2 ON tempstaff.EMPLOYID = temptbl2.EMPID_I
ORDER BY tempstaff.DEPRTMNT, tempstaff.LASTNAME;

