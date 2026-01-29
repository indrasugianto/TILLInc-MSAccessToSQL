-- Query Name: qryCopyDeletedFamilyByClientName
-- Extracted: 2026-01-29 16:09:05

INSERT INTO tblDELETEDPeopleFamily ( IndexedName, ClientIndexedName, RecordDeletedDate, RecordDeletedBy, ClientLastName, ClientFirstName, ClientMiddleInitial, Relationship, Guardian, PrimaryContact, Surrogate, RepPayee, Inactive )
SELECT tblPeopleFamily.IndexedName, tblPeopleFamily.ClientIndexedName, [Forms]![frmMainMenu]![TodaysDate] AS RecordDeletedDate, [Forms]![frmMainMenu]![UserName] AS RecordDeletedBy, tblPeopleFamily.ClientLastName, tblPeopleFamily.ClientFirstName, tblPeopleFamily.ClientMiddleInitial, tblPeopleFamily.Relationship, tblPeopleFamily.Guardian, tblPeopleFamily.PrimaryContact, tblPeopleFamily.Surrogate, tblPeopleFamily.RepPayee, tblPeopleFamily.Inactive
FROM tblPeopleFamily
WHERE (((tblPeopleFamily.ClientIndexedName)=[Forms]![frmMainMenu]![RememberIndexedName]));

