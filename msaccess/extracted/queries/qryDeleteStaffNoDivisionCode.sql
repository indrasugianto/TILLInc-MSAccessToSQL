-- Query Name: qryDeleteStaffNoDivisionCode
-- Extracted: 2026-01-29 16:09:05

DELETE tblStaff.*
FROM tblStaff
WHERE tblStaff.DIVISIONCODE_I Is Null OR Len(tblStaff.DIVISIONCODE_I) = 0;

