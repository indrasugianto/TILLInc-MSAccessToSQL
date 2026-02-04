-- Query Name: qryStaffAndEvalsDeleteInactives
-- Extracted: 2026-02-04 13:04:22

UPDATE temptbl RIGHT JOIN tblStaffEvalsAndSupervisions ON temptbl.EMPLOYID = tblStaffEvalsAndSupervisions.EmployeeID SET tblStaffEvalsAndSupervisions.DeleteFlag = -1
WHERE (((temptbl.EMPLOYID) Is Null));

