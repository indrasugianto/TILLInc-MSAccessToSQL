-- Query Name: ~Ad-hoc Day
-- Extracted: 2026-01-29 16:09:06 (ADO Method)

SELECT tblPeople.IndexedName, tblPeople.Salutation, tblPeople.FirstName, tblPeople.LastName, tblPeople.MailingAddress, tblPeople.MailingCity, tblPeople.MailingState, tblPeople.MailingZIP, tblPeople_1.IndexedName, tblPeople_1.MailingAddress, tblPeopleClientsVendors.LivingWithParentOrGuardian
FROM ((tblPeople INNER JOIN tblPeopleClientsDayServices ON tblPeople.IndexedName = tblPeopleClientsDayServices.IndexedName) INNER JOIN (tblPeople AS tblPeople_1 INNER JOIN tblPeopleFamily ON tblPeople_1.IndexedName = tblPeopleFamily.IndexedName) ON tblPeople.IndexedName = tblPeopleFamily.ClientIndexedName) INNER JOIN tblPeopleClientsVendors ON tblPeople.IndexedName = tblPeopleClientsVendors.IndexedName
WHERE (((tblPeople.IsClientDay)=True) AND ((tblPeople.IsDeceased)=False) AND ((tblPeopleClientsDayServices.Inactive)=False) AND ((tblPeople_1.MailingAddress)=[tblPeople]![MailingAddress]));

