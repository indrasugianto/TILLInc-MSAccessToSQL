-- Query Name: qryCopyDeletedDonors
-- Extracted: 2026-01-29 16:09:05

INSERT INTO tblDELETEDPeopleDonors ( IndexedName, [Index], RecordDeletedDate, RecordDeletedBy, DateOfDonation, DateReceived, DateThankYou, DonationType, SolicitationType, AppealCode, IsGrant, IsRestricted, DonationFrom, DonationFrom1Salutation, DonationFrom1FirstName, DonationFrom1LastName, DonationFrom2Salutation, DonationFrom2FirstName, DonationFrom2LastName, DonationFromCompany, Description, Amount, Inactive )
SELECT tblPeopleDonors.IndexedName, tblPeopleDonors.Index, [Forms]![frmMainMenu]![TodaysDate] AS RecordDeletedDate, [Forms]![frmMainMenu]![UserName] AS RecordDeletedBy, tblPeopleDonors.DateOfDonation, tblPeopleDonors.DateReceived, tblPeopleDonors.DateThankYou, tblPeopleDonors.DonationType, tblPeopleDonors.SolicitationType, tblPeopleDonors.AppealCode, tblPeopleDonors.IsGrant, tblPeopleDonors.IsRestricted, tblPeopleDonors.DonationFrom, tblPeopleDonors.DonationFrom1Salutation, tblPeopleDonors.DonationFrom1FirstName, tblPeopleDonors.DonationFrom1LastName, tblPeopleDonors.DonationFrom2Salutation, tblPeopleDonors.DonationFrom2FirstName, tblPeopleDonors.DonationFrom2LastName, tblPeopleDonors.DonationFromCompany, tblPeopleDonors.Description, tblPeopleDonors.Amount, tblPeopleDonors.Inactive
FROM tblPeopleDonors
WHERE (((tblPeopleDonors.IndexedName)=[Forms]![frmMainMenu]![RememberIndexedName]));

