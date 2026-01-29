-- Query Name: qryClientGovernmentAccounts
-- Extracted: 2026-01-29 16:09:06 (ADO Method)

SELECT tblPeople.LastName, tblPeople.FirstName, tblPeopleClientsDemographics.DateOfBirth, tblPeople.ResLocation, tblPeopleClientsDemographics.SocialSecurityNumber, tblPeopleClientsDemographics.MedicaidNumber, tblPeopleClientsDemographics.MedicareNumber, tblPeopleClientsDemographics.FoodStampsCardNumber
FROM (tblPeople INNER JOIN tblPeopleClientsDemographics ON tblPeople.IndexedName = tblPeopleClientsDemographics.IndexedName) LEFT JOIN tblPeopleClientsResidentialServices ON tblPeople.IndexedName = tblPeopleClientsResidentialServices.IndexedName
WHERE (((tblPeople.isclientres)=True) AND ((tblPeople.IsDeceased)=False) AND ((tblPeopleClientsResidentialServices.Inactive)=False))
ORDER BY tblPeople.LastName, tblPeople.FirstName;

