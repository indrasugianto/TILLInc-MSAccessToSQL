-- Query Name: qryUpdateSpringboardClientsStep5
-- Extracted: 2026-01-29 16:09:05

INSERT INTO temptbl ( IndexedName, SprGroup, LastName, FirstName )
SELECT tblPeopleConsultants.IndexedName, tblPeopleConsultants.SpringboardGroupCode3 AS SprGroup, tblPeople.LastName, tblPeople.FirstName
FROM tblPeopleConsultants INNER JOIN tblPeople ON tblPeopleConsultants.IndexedName = tblPeople.IndexedName
WHERE (((tblPeopleConsultants.SpringboardGroupCode3) Is Not Null) AND ((tblPeopleConsultants.Department)='Springboard') AND ((tblPeopleConsultants.Inactive)=False))
ORDER BY tblPeopleConsultants.SpringboardGroupCode3;

