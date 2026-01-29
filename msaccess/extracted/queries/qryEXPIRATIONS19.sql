-- Query Name: qryEXPIRATIONS19
-- Extracted: 2026-01-29 16:09:05

SELECT tempstaffskills.EMPID_I, tempstaffskills.SKILLNUMBER_I, tempstaffskills.EXPIREDSKILL_I INTO temptbl1
FROM tempstaffskills INNER JOIN tblStaff ON tempstaffskills.EMPID_I = tblStaff.EMPLOYID
WHERE tempstaffskills.SKILLNUMBER_I=1 Or tempstaffskills.SKILLNUMBER_I=2 Or tempstaffskills.SKILLNUMBER_I=3 Or tempstaffskills.SKILLNUMBER_I=15 Or tempstaffskills.SKILLNUMBER_I=22 Or tempstaffskills.SKILLNUMBER_I=30 Or tempstaffskills.SKILLNUMBER_I=31 Or tempstaffskills.SKILLNUMBER_I=32 Or tempstaffskills.SKILLNUMBER_I=33 Or tempstaffskills.SKILLNUMBER_I=34 Or tempstaffskills.SKILLNUMBER_I=35 Or tempstaffskills.SKILLNUMBER_I=36 Or tempstaffskills.SKILLNUMBER_I=39;

