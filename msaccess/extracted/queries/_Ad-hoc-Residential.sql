-- Query Name: ~Ad-hoc-Residential
-- Extracted: 2026-02-04 13:04:22 (ADO Method)

SELECT tblPeople.IndexedName, tblPeople.Salutation, tblPeople.FirstName, tblPeople.LastName, tblPeople.MailingAddress, tblPeople.MailingCity, tblPeople.MailingState, tblPeople.MailingZIP
FROM tblPeople INNER JOIN tblPeopleClientsResidentialServices ON tblPeople.IndexedName = tblPeopleClientsResidentialServices.IndexedName
WHERE (((tblPeople.IsClientRes)=True) AND ((tblPeople.IsDeceased)=False) AND ((tblPeopleClientsResidentialServices.Inactive)=False));

