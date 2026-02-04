-- Query Name: qryCopyDeletedVocational
-- Extracted: 2026-02-04 13:04:21

INSERT INTO tblDELETEDPeopleClientsVocationalServices ( IndexedName, RecordDeletedDate, RecordDeletedBy, CityTown, Location, StartDate, EndDate, TerminationReason, Funding, ContractNumber, ActivityCode, ContractNumber2, ActivityCode2, Rate, DMRAnnual, IntensityLevel, Inactive, DateInactive )
SELECT tblPeopleClientsVocationalServices.IndexedName, [Forms]![frmMainMenu]![TodaysDate] AS RecordDeletedDate, [Forms]![frmMainMenu]![UserName] AS RecordDeletedBy, tblPeopleClientsVocationalServices.CityTown, tblPeopleClientsVocationalServices.Location, tblPeopleClientsVocationalServices.StartDate, tblPeopleClientsVocationalServices.EndDate, tblPeopleClientsVocationalServices.TerminationReason, tblPeopleClientsVocationalServices.Funding, tblPeopleClientsVocationalServices.ContractNumber, tblPeopleClientsVocationalServices.ActivityCode, tblPeopleClientsVocationalServices.ContractNumber2, tblPeopleClientsVocationalServices.ActivityCode2, tblPeopleClientsVocationalServices.Rate, tblPeopleClientsVocationalServices.DMRAnnual, tblPeopleClientsVocationalServices.IntensityLevel, tblPeopleClientsVocationalServices.Inactive, tblPeopleClientsVocationalServices.DateInactive
FROM tblPeopleClientsVocationalServices
WHERE (((tblPeopleClientsVocationalServices.IndexedName)=[Forms]![frmMainMenu]![RememberIndexedName]));

