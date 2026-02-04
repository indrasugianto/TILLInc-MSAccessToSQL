-- Query Name: qryRUNREPORT05
-- Extracted: 2026-02-04 13:04:22

INSERT INTO temptbl ( EmplID, LName, FName, SKILLEXPIREDDATE, REDFLAG, LOCATION, SKILL )
SELECT tblStaffEvalsAndSupervisions.EmployeeID, tblStaffEvalsAndSupervisions.LastName, tblStaffEvalsAndSupervisions.FirstName, tblStaffEvalsAndSupervisions.EvalDueBy, True AS Expr1, tblStaffEvalsAndSupervisions.Loc AS LOCATION, 'EvalDueBy' AS SKILL
FROM tblStaffEvalsAndSupervisions INNER JOIN tempstaff ON tblStaffEvalsAndSupervisions.EmployeeID = tempstaff.EMPLOYID
WHERE (((tblStaffEvalsAndSupervisions.EvalDueBy) Is Null));

