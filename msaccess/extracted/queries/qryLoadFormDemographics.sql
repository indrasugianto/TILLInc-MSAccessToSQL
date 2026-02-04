-- Query Name: qryLoadFormDemographics
-- Extracted: 2026-02-04 13:04:22

INSERT INTO tblPeopleClientsDemographics ( IndexedName, RecordAddedDate, RecordAddedBy )
SELECT Left([Form_frmPeople].[IndexedName],160) AS IndexedName, Format(Now(),'mm/dd/yyyy') AS RecordAddedDate, [Form_frmMainMenu].[UserName] AS RecordAddedBy;

