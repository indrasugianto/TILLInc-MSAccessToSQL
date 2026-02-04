-- Query Name: qryPeopleAddFamilyRepPayee
-- Extracted: 2026-02-04 13:04:22

UPDATE (tblPeople INNER JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.IndexedName) INNER JOIN tblPeopleClientsDemographics ON tblPeopleFamily.ClientIndexedName = tblPeopleClientsDemographics.IndexedName SET tblPeopleClientsDemographics.RepresentativePayee = [tblPeople]![FirstName] & " " & [tblPeople]![LastName], tblPeopleClientsDemographics.RepPayeeAddress = [tblPeople]![MailingAddress], tblPeopleClientsDemographics.RepPayeeCity = [tblPeople]![MailingCity], tblPeopleClientsDemographics.RepPayeeState = [tblPeople]![MailingState], tblPeopleClientsDemographics.RepPayeeZIP = [tblPeople]![MailingZIP], tblPeopleClientsDemographics.RepPayeePhone = [tblPeople]![HomePhone], tblPeopleClientsDemographics.RepPayeeAddressValidated = [tblPeople]![MailingAddressValidated]
WHERE (((tblPeopleFamily.IndexedName)=[Forms]![frmPeople]![frmPeopleFamily]![IndexedName]) AND ((tblPeopleClientsDemographics.IndexedName)=[Forms]![frmPeople]![frmPeopleFamily]![ClientIndexedName]));

