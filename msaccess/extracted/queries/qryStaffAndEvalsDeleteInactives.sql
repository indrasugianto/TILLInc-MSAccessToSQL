-- Query Name: qryStaffAndEvalsDeleteInactives
-- Extracted: 2026-01-29 16:09:05

UPDATE temptbl RIGHT JOIN tblStaffEvalsAndSupervisions ON temptbl.EMPLOYID = tblStaffEvalsAndSupervisions.EmployeeID SET tblStaffEvalsAndSupervisions.DeleteFlag = -1
WHERE (((temptbl.EMPLOYID) Is Null));

