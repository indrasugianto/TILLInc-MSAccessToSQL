-- Query Name: qryDeleteCorruptedPeopleRecords
-- Extracted: 2026-02-04 13:04:21

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

