-- Query Name: qryDeleteCorruptedPeopleRecords
-- Extracted: 2026-01-29 16:09:05

DELETE tblPeople.*
FROM tblPeople
WHERE tblPeople.IndexedName = '///'
    OR (
        tblPeople.LastName IS NULL
        AND tblPeople.FirstName IS NOT NULL
        AND tblPeople.CompanyOrganization IS NULL
    )
    OR (
        tblPeople.FirstName IS NULL
        AND tblPeople.LastName IS NOT NULL
        AND tblPeople.CompanyOrganization IS NULL
    );

