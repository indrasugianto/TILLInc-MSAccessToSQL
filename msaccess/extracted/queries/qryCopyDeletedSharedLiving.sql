-- Query Name: qryCopyDeletedSharedLiving
-- Extracted: 2026-01-29 16:09:05

INSERT INTO tblDELETEDPeopleClientsSharedLivingServices ( IndexedName, RecordDeletedDate, RecordDeletedBy, CityTown, Location, ContractNumber, ActivityCode, ContractNumber2, ActivityCode2, CostCenter, StartDate, EndDate, TerminationReason, Portion, CaseManager, Inactive, DateInactive )
SELECT tblPeopleClientsSharedLivingServices.IndexedName, [Forms]![frmMainMenu]![TodaysDate] AS RecordDeletedDate, [Forms]![frmMainMenu]![UserName] AS RecordDeletedBy, tblPeopleClientsSharedLivingServices.CityTown, tblPeopleClientsSharedLivingServices.Location, tblPeopleClientsSharedLivingServices.ContractNumber, tblPeopleClientsSharedLivingServices.ActivityCode, tblPeopleClientsSharedLivingServices.ContractNumber2, tblPeopleClientsSharedLivingServices.ActivityCode2, tblPeopleClientsSharedLivingServices.CostCenter, tblPeopleClientsSharedLivingServices.StartDate, tblPeopleClientsSharedLivingServices.EndDate, tblPeopleClientsSharedLivingServices.TerminationReason, tblPeopleClientsSharedLivingServices.Portion, tblPeopleClientsSharedLivingServices.CaseManager, tblPeopleClientsSharedLivingServices.Inactive, tblPeopleClientsSharedLivingServices.DateInactive
FROM tblPeopleClientsSharedLivingServices
WHERE (((tblPeopleClientsSharedLivingServices.IndexedName)=[Forms]![frmMainMenu]![RememberIndexedName]));

