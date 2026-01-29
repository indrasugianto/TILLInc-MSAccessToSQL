-- Query Name: qryDeleteSkillsNotTracked
-- Extracted: 2026-01-29 16:09:05

DELETE tblStaffSkills.*
FROM tblStaffSkills
WHERE tblStaffSkills.SKILLNUMBER_I <>1 And 
tblStaffSkills.SKILLNUMBER_I <>2 And 
tblStaffSkills.SKILLNUMBER_I <>3 And 
tblStaffSkills.SKILLNUMBER_I <>35 And 
tblStaffSkills.SKILLNUMBER_I <>15 And 
tblStaffSkills.SKILLNUMBER_I <>22 And 
tblStaffSkills.SKILLNUMBER_I <>30 And 
tblStaffSkills.SKILLNUMBER_I <>31 And 
tblStaffSkills.SKILLNUMBER_I <>32 And 
tblStaffSkills.SKILLNUMBER_I <>33 And 
tblStaffSkills.SKILLNUMBER_I <>34 And 
tblStaffSkills.SKILLNUMBER_I <>36 And
tblStaffSkills.SKILLNUMBER_I <>39;

