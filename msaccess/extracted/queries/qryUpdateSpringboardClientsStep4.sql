-- Query Name: qryUpdateSpringboardClientsStep4
-- Extracted: 2026-02-04 13:04:22

INSERT INTO temptbl ( IndexedName, SprGroup, LastName, FirstName )
SELECT tblPeopleConsultants.IndexedName, tblPeopleConsultants.SpringboardGroupCode2 AS SprGroup, tblPeople.LastName, tblPeople.FirstName
FROM tblPeopleConsultants INNER JOIN tblPeople ON tblPeopleConsultants.IndexedName = tblPeople.IndexedName
WHERE (((tblPeopleConsultants.SpringboardGroupCode2) Is Not Null) AND ((tblPeopleConsultants.Department)='Springboard') AND ((tblPeopleConsultants.Inactive)=False))
ORDER BY tblPeopleConsultants.SpringboardGroupCode2;

