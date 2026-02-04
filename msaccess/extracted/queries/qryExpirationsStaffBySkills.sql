-- Query Name: qryExpirationsStaffBySkills
-- Extracted: 2026-02-04 13:04:22

TRANSFORM First(qryExpirationsStaffCull.EXPIREDSKILL_I) AS FirstOfEXPIREDSKILL_I
SELECT qryExpirationsStaffCull.DEPRTMNT, qryExpirationsStaffCull.LASTNAME, qryExpirationsStaffCull.FRSTNAME, qryExpirationsStaffCull.JOBTITLE
FROM qryExpirationsStaffCull
GROUP BY qryExpirationsStaffCull.DEPRTMNT, qryExpirationsStaffCull.LASTNAME, qryExpirationsStaffCull.FRSTNAME, qryExpirationsStaffCull.JOBTITLE
PIVOT qryExpirationsStaffCull.SkillDesc;

