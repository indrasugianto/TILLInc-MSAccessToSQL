-- Query Name: qryLoadConsultantsTable
-- Extracted: 2026-01-29 16:09:05

INSERT INTO tblPeopleConsultants ( IndexedName, RecordAddedDate, RecordAddedBy )
SELECT """ & Left(Form_frmPeople.IndexedName, 160) & """ AS IndexedName, Now() AS RecordAddedDate, """ & Form_frmMainMenu.UserName & """ AS RecordAddedBy;

