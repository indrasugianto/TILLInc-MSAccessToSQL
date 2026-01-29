-- Query Name: qryCopyDeletedISS
-- Extracted: 2026-01-29 16:09:05

INSERT INTO tblDELETEDPeopleClientsIndividualSupportServices ( IndexedName, RecordDeletedDate, RecordDeletedBy, CityTown, Location, ContractNumber, ActivityCode, ContractNumber2, ActivityCode2, CostCenter, StartDate, EndDate, TerminationReason, CaseManager, FundingSource, Provider, Inactive, DateInactive )
SELECT tblPeopleClientsIndividualSupportServices.IndexedName, [Forms]![frmMainMenu]![TodaysDate] AS RecordDeletedDate, [Forms]![frmMainMenu]![UserName] AS RecordDeletedBy, tblPeopleClientsIndividualSupportServices.CityTown, tblPeopleClientsIndividualSupportServices.Location, tblPeopleClientsIndividualSupportServices.ContractNumber, tblPeopleClientsIndividualSupportServices.ActivityCode, tblPeopleClientsIndividualSupportServices.ContractNumber2, tblPeopleClientsIndividualSupportServices.ActivityCode2, tblPeopleClientsIndividualSupportServices.CostCenter, tblPeopleClientsIndividualSupportServices.StartDate, tblPeopleClientsIndividualSupportServices.EndDate, tblPeopleClientsIndividualSupportServices.TerminationReason, tblPeopleClientsIndividualSupportServices.CaseManager, tblPeopleClientsIndividualSupportServices.FundingSource, tblPeopleClientsIndividualSupportServices.Provider, tblPeopleClientsIndividualSupportServices.Inactive, tblPeopleClientsIndividualSupportServices.DateInactive
FROM tblPeopleClientsIndividualSupportServices
WHERE (((tblPeopleClientsIndividualSupportServices.IndexedName)=[Forms]![frmMainMenu]![RememberIndexedName]));

