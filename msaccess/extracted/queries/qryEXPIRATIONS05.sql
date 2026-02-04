-- Query Name: qryEXPIRATIONS05
-- Extracted: 2026-02-04 13:04:22

SELECT tblStaffSkills.* INTO tempstaffskills
FROM tblStaff LEFT JOIN tblStaffSkills ON tblStaff.EMPLOYID = tblStaffSkills.EMPID_I
WHERE tblStaffSkills.SKILLNUMBER_I=1  OR 
              tblStaffSkills.SKILLNUMBER_I=2  OR 
              tblStaffSkills.SKILLNUMBER_I=3  OR 
              tblStaffSkills.SKILLNUMBER_I=15 OR 
              tblStaffSkills.SKILLNUMBER_I=22 OR 
              tblStaffSkills.SKILLNUMBER_I=30 OR 
              tblStaffSkills.SKILLNUMBER_I=31 OR 
              tblStaffSkills.SKILLNUMBER_I=32 OR 
              tblStaffSkills.SKILLNUMBER_I=33 OR 
              tblStaffSkills.SKILLNUMBER_I=34 OR 
              tblStaffSkills.SKILLNUMBER_I=35 OR 
              tblStaffSkills.SKILLNUMBER_I=36 OR 
              tblStaffSkills.SKILLNUMBER_I=39;

