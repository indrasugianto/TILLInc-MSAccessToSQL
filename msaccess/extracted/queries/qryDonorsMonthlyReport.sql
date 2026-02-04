-- Query Name: qryDonorsMonthlyReport
-- Extracted: 2026-02-04 13:04:22

SELECT [tblPeopleDonors]![DateOfDonation] AS DonationDate, tblPeopleDonors.DateReceived, tblPeopleDonors.DateThankYou, tblPeopleDonors.IndexedName, DCount("IndexedName","tblPeopleDonors","Indexedname='" & [tblPeopleDonors].[IndexedName] & "'") AS NumDonations, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, tblPeople.CompanyOrganization, tblPeople.MailingAddress, tblPeople.MailingCity, tblPeople.MailingState, tblPeople.MailingZIP, tblPeopleDonors.DonationType, tblPeopleDonors.SolicitationType, tblPeopleDonors.AppealCode, tblPeopleDonors.IsGrant, tblPeopleDonors.DonationFrom1Salutation, tblPeopleDonors.DonationFrom1FirstName, tblPeopleDonors.DonationFrom1LastName, tblPeopleDonors.DonationFrom2Salutation, tblPeopleDonors.DonationFrom2FirstName, tblPeopleDonors.DonationFrom2LastName, tblPeopleDonors.DonationFromCompany, tblPeopleDonors.Description, tblPeopleDonors.Amount INTO temptbl
FROM tblPeople INNER JOIN tblPeopleDonors ON tblPeople.IndexedName = tblPeopleDonors.IndexedName
ORDER BY tblPeopleDonors.IndexedName;

