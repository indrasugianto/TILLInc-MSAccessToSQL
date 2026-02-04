-- Query Name: qryDeleteNullPeople
-- Extracted: 2026-02-04 13:04:22

DELETE tblPeople.*, tblPeople.IndexedName
FROM tblPeople
WHERE (((tblPeople.IndexedName)='///'));

