-- Query Name: qryUpdateAllPeopleTable
-- Extracted: 2026-02-04 13:04:22

SELECT LastName & ', ' & FirstName & ' ' & IIf(Len(MiddleInitial)>0,MiddleInitial,'') AS Person, FirstName & ' ' & IIf(Len(MiddleInitial)<=0,'',MiddleInitial & ' ') & LastName AS PersonByFirstName, tblPeople.* INTO AllPeopleRecords
FROM tblPeople
WHERE Left(IndexedName,3) <> '///'
ORDER BY tblPeople.IndexedName;

