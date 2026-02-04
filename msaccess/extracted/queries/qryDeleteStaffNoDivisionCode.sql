-- Query Name: qryDeleteStaffNoDivisionCode
-- Extracted: 2026-02-04 13:04:22

DELETE tblStaff.*
FROM tblStaff
WHERE tblStaff.DIVISIONCODE_I Is Null OR Len(tblStaff.DIVISIONCODE_I) = 0;

