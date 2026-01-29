-- Query Name: qryEXPIRATIONS20
-- Extracted: 2026-01-29 16:09:05

SELECT temptbl1.*, DLookUp("Skill","catSkills","SkillID=" & [SKILLNUMBER_I]) AS SkillDesc INTO temptbl2
FROM temptbl1;

