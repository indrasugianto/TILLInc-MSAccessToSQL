-- Query Name: qryCopyDeletedAdultCoaching
-- Extracted: 2026-02-04 13:04:21

INSERT INTO tblDELETEDPeopleClientsAdultCoaching ( IndexedName, RecordDeletedDate, RecordDeletedBy, CityTown, Location, ContractNumber, ActivityCode, ContractNumber2, ActivityCode2, StartDate, EndDate, TerminationReason, CaseManager, FundingSource, Provider, Inactive, DateInactive )
SELECT tblPeopleClientsAdultCoaching.IndexedName, [Forms]![frmMainMenu]![TodaysDate] AS RecordDeletedDate, [Forms]![frmMainMenu]![UserName] AS RecordDeletedBy, tblPeopleClientsAdultCoaching.CityTown, tblPeopleClientsAdultCoaching.Location, tblPeopleClientsAdultCoaching.ContractNumber, tblPeopleClientsAdultCoaching.ActivityCode, tblPeopleClientsAdultCoaching.ContractNumber2, tblPeopleClientsAdultCoaching.ActivityCode2, tblPeopleClientsAdultCoaching.StartDate, tblPeopleClientsAdultCoaching.EndDate, tblPeopleClientsAdultCoaching.TerminationReason, tblPeopleClientsAdultCoaching.CaseManager, tblPeopleClientsAdultCoaching.FundingSource, tblPeopleClientsAdultCoaching.Provider, tblPeopleClientsAdultCoaching.Inactive, tblPeopleClientsAdultCoaching.DateInactive
FROM tblPeopleClientsAdultCoaching
WHERE (((tblPeopleClientsAdultCoaching.IndexedName)=[Forms]![frmMainMenu]![RememberIndexedName]));

