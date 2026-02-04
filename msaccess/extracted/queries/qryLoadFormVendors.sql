-- Query Name: qryLoadFormVendors
-- Extracted: 2026-02-04 13:04:22

INSERT INTO tblPeopleClientsVendors ( IndexedName, RecordAddedDate, RecordAddedBy, LivingWithParentOrGuardian, LivingIndependently )
SELECT Left([Form_frmPeople].[IndexedName],160) AS IndexedName, Now() AS RecordAddedDate, [Form_frmMainMenu].[UserName] AS RecordAddedBy, False AS LivingWithParentOrGuardian, False AS LivingIndependently;

