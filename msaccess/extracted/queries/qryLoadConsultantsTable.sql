-- Query Name: qryLoadConsultantsTable
-- Extracted: 2026-02-04 13:04:22

INSERT INTO tblPeopleConsultants ( IndexedName, RecordAddedDate, RecordAddedBy )
SELECT """ & Left(Form_frmPeople.IndexedName, 160) & """ AS IndexedName, Now() AS RecordAddedDate, """ & Form_frmMainMenu.UserName & """ AS RecordAddedBy;

