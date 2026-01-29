-- Query Name: qryDeleteNullPeople
-- Extracted: 2026-01-29 16:09:05

DELETE tblPeople.*, tblPeople.IndexedName
FROM tblPeople
WHERE (((tblPeople.IndexedName)='///'));

