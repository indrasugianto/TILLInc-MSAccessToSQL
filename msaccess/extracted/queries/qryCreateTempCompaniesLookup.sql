-- Query Name: qryCreateTempCompaniesLookup
-- Extracted: 2026-02-04 13:04:21

SELECT DISTINCT tblPeople.CompanyOrganization, tblPeople.IndexedName, Len([CompanyOrganization]) AS Crit INTO temptbl0
FROM tblPeople
WHERE (((Len([CompanyOrganization])) > 0))
ORDER BY tblPeople.CompanyOrganization;

