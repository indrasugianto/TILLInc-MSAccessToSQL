-- Query Name: qryLoadFormVendors
-- Extracted: 2026-01-29 16:09:05

INSERT INTO tblPeopleClientsVendors ( IndexedName, RecordAddedDate, RecordAddedBy, LivingWithParentOrGuardian, LivingIndependently )
SELECT Left([Form_frmPeople].[IndexedName],160) AS IndexedName, Now() AS RecordAddedDate, [Form_frmMainMenu].[UserName] AS RecordAddedBy, False AS LivingWithParentOrGuardian, False AS LivingIndependently;

