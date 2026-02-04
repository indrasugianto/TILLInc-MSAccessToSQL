-- Query Name: qryCopyDeletedNHDay
-- Extracted: 2026-02-04 13:04:21

INSERT INTO tblDELETEDPeopleClientsNHDay ( IndexedName, RecordDeletedDate, RecordDeletedBy, CityTown, Location, StartDate, EndDate, TerminationReason, Funding, ContractNumber, ActivityCode, ContractNumber2, ActivityCode2, Rate, DMRAnnual, IntensityLevel, Inactive, DateInactive )
SELECT tblPeopleClientsNHDay.IndexedName, [Forms]![frmMainMenu]![TodaysDate] AS RecordDeletedDate, [Forms]![frmMainMenu]![UserName] AS RecordDeletedBy, tblPeopleClientsNHDay.CityTown, tblPeopleClientsNHDay.Location, tblPeopleClientsNHDay.StartDate, tblPeopleClientsNHDay.EndDate, tblPeopleClientsNHDay.TerminationReason, tblPeopleClientsNHDay.Funding, tblPeopleClientsNHDay.ContractNumber, tblPeopleClientsNHDay.ActivityCode, tblPeopleClientsNHDay.ContractNumber2, tblPeopleClientsNHDay.ActivityCode2, tblPeopleClientsNHDay.Rate, tblPeopleClientsNHDay.DMRAnnual, tblPeopleClientsNHDay.IntensityLevel, tblPeopleClientsNHDay.Inactive, tblPeopleClientsNHDay.DateInactive
FROM tblPeopleClientsNHDay
WHERE (((tblPeopleClientsNHDay.IndexedName)=[Forms]![frmMainMenu]![RememberIndexedName]));

