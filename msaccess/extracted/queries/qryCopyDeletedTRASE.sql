-- Query Name: qryCopyDeletedTRASE
-- Extracted: 2026-02-04 13:04:21

INSERT INTO tblDELETEDPeopleClientsTRASEServices ( IndexedName, RecordDeletedDate, RecordDeletedBy, School, Inactive, DateInactive )
SELECT tblPeopleClientsTRASEServices.IndexedName, [Forms]![frmMainMenu]![TodaysDate] AS RecordDeletedDate, [Forms]![frmMainMenu]![UserName] AS RecordDeletedBy, tblPeopleClientsTRASEServices.School, tblPeopleClientsTRASEServices.Inactive, tblPeopleClientsTRASEServices.DateInactive
FROM tblPeopleClientsTRASEServices
WHERE (((tblPeopleClientsTRASEServices.IndexedName)=[Forms]![frmMainMenu]![RememberIndexedName]));

