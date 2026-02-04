-- Query Name: qryPeopleTILLGamesEmail
-- Extracted: 2026-02-04 13:04:22 (ADO Method)

SELECT [tblPeople]![FirstName] & " " & [tblPeople]![LastName] AS FamilyName, StrConv([tblpeople].[EmailAddress],2) AS FamilyEmailAddress, [tblPeopleFamily]![ClientFirstName] & " " & [tblPeopleFamily]![ClientLastName] AS ClientName, tblPeople_1.ResLocation, tblPeople_1.CLOLocation
FROM (((tblPeople LEFT JOIN tblPeopleFamily ON tblPeople.IndexedName = tblPeopleFamily.IndexedName) LEFT JOIN tblPeople AS tblPeople_1 ON tblPeopleFamily.ClientIndexedName = tblPeople_1.IndexedName) LEFT JOIN tblPeopleClientsResidentialServices ON tblPeople_1.IndexedName = tblPeopleClientsResidentialServices.IndexedName) LEFT JOIN tblPeopleClientsCLOServices ON tblPeople_1.IndexedName = tblPeopleClientsCLOServices.IndexedName
WHERE (((StrConv([tblpeople].[EmailAddress],2)) Like "*@*") AND ((tblPeople_1.IsClientRes)=True) AND ((tblPeopleFamily.Inactive)=False) AND ((tblPeopleClientsResidentialServices.Inactive)=False)) OR (((StrConv([tblpeople].[EmailAddress],2)) Like "*@*") AND ((tblPeople_1.IsCilentCLO)=True) AND ((tblPeopleFamily.Inactive)=False) AND ((tblPeopleClientsCLOServices.Inactive)=False));

