-- Query Name: qryEXPIRATIONS01
-- Extracted: 2026-02-04 13:04:22

SELECT tblStaff.* INTO tempstaff
FROM tblStaff
ORDER BY tblStaff.LASTNAME, tblStaff.FRSTNAME;

