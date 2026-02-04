-- Query Name: qryEXPIRATIONS18
-- Extracted: 2026-02-04 13:04:22

INSERT INTO tblExpirations ( Location, RecordType, LastName, FirstName, JobTitle, Supervisor, AdjustedStartDate )
SELECT [tempstaff]![DEPRTMNT] AS Location, "Staff" AS RecordType, tempstaff.LASTNAME, tempstaff.FRSTNAME, tempstaff.JOBTITLE, tempstaff.SUPERVISORCODE_I, tempstaff.BENADJDATE AS AdjustedStartDate
FROM tempstaff INNER JOIN tempstaffskills ON tempstaff.EMPLOYID = tempstaffskills.EMPID_I
WHERE tempstaff.DEPRTMNT Is Not Null And tempstaff.LastName Is Not Null And tempstaff.FRSTNAME Is Not Null
ORDER BY tempstaff.LASTNAME, tempstaff.FRSTNAME;

