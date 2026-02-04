-- Query Name: qryEXPIRATIONS20
-- Extracted: 2026-02-04 13:04:22

SELECT temptbl1.*, DLookUp("Skill","catSkills","SkillID=" & [SKILLNUMBER_I]) AS SkillDesc INTO temptbl2
FROM temptbl1;

