-- Query Name: qryUpdateSpringboardLeaders03
-- Extracted: 2026-02-04 13:04:22

SELECT tblPeopleConsultants.IndexedName, tblPeopleConsultants.SpringboardGroupCode1 AS SprGroup, tblPeople.LastName, tblPeople.FirstName INTO temptbl
FROM tblPeopleConsultants INNER JOIN tblPeople ON tblPeopleConsultants.IndexedName = tblPeople.IndexedName
WHERE (((tblPeopleConsultants.SpringboardGroupCode1) Is Not Null) AND ((tblPeopleConsultants.Department)='Springboard') AND ((tblPeopleConsultants.Inactive)=False))
ORDER BY tblPeopleConsultants.SpringboardGroupCode1;

