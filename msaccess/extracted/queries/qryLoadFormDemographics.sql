-- Query Name: qryLoadFormDemographics
-- Extracted: 2026-01-29 16:09:05

INSERT INTO tblPeopleClientsDemographics ( IndexedName, RecordAddedDate, RecordAddedBy )
SELECT Left([Form_frmPeople].[IndexedName],160) AS IndexedName, Format(Now(),'mm/dd/yyyy') AS RecordAddedDate, [Form_frmMainMenu].[UserName] AS RecordAddedBy;

