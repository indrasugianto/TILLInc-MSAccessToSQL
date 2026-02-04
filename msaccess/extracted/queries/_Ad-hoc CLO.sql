-- Query Name: ~Ad-hoc CLO
-- Extracted: 2026-02-04 13:04:22 (ADO Method)

SELECT tblPeople.IndexedName, tblPeople.Salutation, tblPeople.FirstName, tblPeople.LastName, tblPeople.MailingAddress, tblPeople.MailingCity, tblPeople.MailingState, tblPeople.MailingZIP
FROM tblPeople INNER JOIN tblPeopleClientsCLOServices ON tblPeople.IndexedName = tblPeopleClientsCLOServices.IndexedName
WHERE (((tblPeople.IsCilentCLO)=True) AND ((tblPeople.IsDeceased)=False) AND ((tblPeopleClientsCLOServices.Inactive)=False));

