-- Query Name: qryCopyDeletedAdultCompanion
-- Extracted: 2026-01-29 16:09:05

INSERT INTO tblDELETEDPeopleClientsAdultCompanion ( IndexedName, RecordDeletedDate, RecordDeletedBy, CityTown, Location, ContractNumber, ActivityCode, ContractNumber2, ActivityCode2, StartDate, EndDate, TerminationReason, CaseManager, FundingSource, Provider, Inactive, DateInactive )
SELECT tblPeopleClientsAdultCompanion.IndexedName, [Forms]![frmMainMenu]![TodaysDate] AS RecordDeletedDate, [Forms]![frmMainMenu]![UserName] AS RecordDeletedBy, tblPeopleClientsAdultCompanion.CityTown, tblPeopleClientsAdultCompanion.Location, tblPeopleClientsAdultCompanion.ContractNumber, tblPeopleClientsAdultCompanion.ActivityCode, tblPeopleClientsAdultCompanion.ContractNumber2, tblPeopleClientsAdultCompanion.ActivityCode2, tblPeopleClientsAdultCompanion.StartDate, tblPeopleClientsAdultCompanion.EndDate, tblPeopleClientsAdultCompanion.TerminationReason, tblPeopleClientsAdultCompanion.CaseManager, tblPeopleClientsAdultCompanion.FundingSource, tblPeopleClientsAdultCompanion.Provider, tblPeopleClientsAdultCompanion.Inactive, tblPeopleClientsAdultCompanion.DateInactive
FROM tblPeopleClientsAdultCompanion
WHERE (((tblPeopleClientsAdultCompanion.IndexedName)=[Forms]![frmMainMenu]![RememberIndexedName]));

