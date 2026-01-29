-- Query Name: qryEXPIRATIONS01
-- Extracted: 2026-01-29 16:09:05

SELECT tblStaff.* INTO tempstaff
FROM tblStaff
ORDER BY tblStaff.LASTNAME, tblStaff.FRSTNAME;

