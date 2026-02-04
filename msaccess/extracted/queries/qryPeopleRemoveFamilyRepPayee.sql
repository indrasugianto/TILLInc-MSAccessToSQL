-- Query Name: qryPeopleRemoveFamilyRepPayee
-- Extracted: 2026-02-04 13:04:22

UPDATE (tblPeople INNER JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.IndexedName) INNER JOIN tblPeopleClientsDemographics ON tblPeopleFamily.ClientIndexedName = tblPeopleClientsDemographics.IndexedName SET tblPeopleClientsDemographics.RepresentativePayee = Null, tblPeopleClientsDemographics.RepPayeeAddress = Null, tblPeopleClientsDemographics.RepPayeeCity = Null, tblPeopleClientsDemographics.RepPayeeState = Null, tblPeopleClientsDemographics.RepPayeeZIP = Null, tblPeopleClientsDemographics.RepPayeePhone = Null, tblPeopleClientsDemographics.RepPayeeAddressValidated = False
WHERE (((tblPeopleFamily.IndexedName)=Forms!frmPeople!frmPeopleFamily!IndexedName) And ((tblPeopleClientsDemographics.IndexedName)=Forms!frmPeople!frmPeopleFamily!ClientIndexedName));

