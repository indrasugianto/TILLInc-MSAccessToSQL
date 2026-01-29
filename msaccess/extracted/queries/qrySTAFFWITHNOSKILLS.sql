-- Query Name: qrySTAFFWITHNOSKILLS
-- Extracted: 2026-01-29 16:09:06 (ADO Method)

SELECT tblStaff.EMPLOYID, tblStaff.LASTNAME, tblStaff.FRSTNAME, tblStaff.DIVISIONCODE_I, tblStaff.DEPRTMNT, tblStaff.JOBTITLE, DCount("EMPID_I","tblStaffSkills","EMPID_I='" & [tblStaff]![EMPLOYID] & "'") AS SkillsCount
FROM tblStaff
WHERE (((tblStaff.LASTNAME)<>"EXAMPLE") And ((DCount("EMPID_I","tblStaffSkills","EMPID_I='" & tblStaff!EMPLOYID & "'"))=0));

