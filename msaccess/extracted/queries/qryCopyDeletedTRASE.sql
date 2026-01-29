-- Query Name: qryCopyDeletedTRASE
-- Extracted: 2026-01-29 16:09:05

INSERT INTO tblDELETEDPeopleClientsTRASEServices ( IndexedName, RecordDeletedDate, RecordDeletedBy, School, Inactive, DateInactive )
SELECT tblPeopleClientsTRASEServices.IndexedName, [Forms]![frmMainMenu]![TodaysDate] AS RecordDeletedDate, [Forms]![frmMainMenu]![UserName] AS RecordDeletedBy, tblPeopleClientsTRASEServices.School, tblPeopleClientsTRASEServices.Inactive, tblPeopleClientsTRASEServices.DateInactive
FROM tblPeopleClientsTRASEServices
WHERE (((tblPeopleClientsTRASEServices.IndexedName)=[Forms]![frmMainMenu]![RememberIndexedName]));

