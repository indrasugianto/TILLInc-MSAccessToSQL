-- Query Name: qryDonations
-- Extracted: 2026-02-04 13:04:22 (ADO Method)

SELECT tblPeople.IndexedName, tblPeople.FirstName, tblPeople.LastName, tblPeople.CompanyOrganization, tblPeople.IsDeceased, tblPeopleDonors.Inactive, tblPeopleDonors.DateOfDonation, DateValue([DateOfDonation]) AS DateofDonationNumeric, tblPeopleDonors.DonationType, tblPeopleDonors.IsGrant, tblPeopleDonors.IsRestricted, tblPeopleDonors.Description, tblPeopleDonors.Amount
FROM tblPeopleDonors RIGHT JOIN tblPeople ON tblPeopleDonors.IndexedName = tblPeople.IndexedName
WHERE (((tblPeopleDonors.DateOfDonation) Is Not Null) AND ((Len([DateOfDonation]))>0) AND ((tblPeople.IsDonor)=True));

