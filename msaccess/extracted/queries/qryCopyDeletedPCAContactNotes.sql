-- Query Name: qryCopyDeletedPCAContactNotes
-- Extracted: 2026-02-04 13:04:21

INSERT INTO tblDELETEDPeopleClientsPCAServicesContactNotes ( IndexedName, RecordNumber, RecordDeletedDate, RecordDeletedBy, DateOfEntry, Staff, ContactType, BillCode, Units, Activity, Comments )
SELECT tblPeopleClientsPCAServicesContactNotes.IndexedName, tblPeopleClientsPCAServicesContactNotes.RecordNumber, [Forms]![frmMainMenu]![TodaysDate] AS RecordDeletedDate, [Forms]![frmMainMenu]![UserName] AS RecordDeletedBy, tblPeopleClientsPCAServicesContactNotes.DateOfEntry, tblPeopleClientsPCAServicesContactNotes.Staff, tblPeopleClientsPCAServicesContactNotes.ContactType, tblPeopleClientsPCAServicesContactNotes.BillCode, tblPeopleClientsPCAServicesContactNotes.Units, tblPeopleClientsPCAServicesContactNotes.Activity, tblPeopleClientsPCAServicesContactNotes.Comments
FROM tblPeopleClientsPCAServicesContactNotes
WHERE (((tblPeopleClientsPCAServicesContactNotes.IndexedName)=[Forms]![frmMainMenu]![RememberIndexedName]));

