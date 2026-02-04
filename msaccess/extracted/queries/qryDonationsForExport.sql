-- Query Name: qryDonationsForExport
-- Extracted: 2026-02-04 13:04:22

SELECT qryDonations.FirstName, qryDonations.LastName, qryDonations.CompanyOrganization, qryDonations.IsDeceased, qryDonations.Inactive, qryDonations.DateOfDonation, qryDonations.DonationType, qryDonations.IsGrant, qryDonations.IsRestricted, qryDonations.Description, qryDonations.Amount
FROM qryDonations
WHERE (((qryDonations.DateOfDonation) Is Not Null) And ((qryDonations.DateofDonationNumeric)>=Forms!frmRptExpMAILINGSANDSPREADSHEETS!DonationsStartDateNumeric And (qryDonations.DateofDonationNumeric)<=Forms!frmRptExpMAILINGSANDSPREADSHEETS!DonationsEndDateNumeric) And ((qryDonations.IndexedName)=Forms!frmRptExpMAILINGSANDSPREADSHEETS!SelectDonor) And ((Len([DateOfDonation]))>0));

