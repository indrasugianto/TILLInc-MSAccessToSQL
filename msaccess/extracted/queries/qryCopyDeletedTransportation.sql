-- Query Name: qryCopyDeletedTransportation
-- Extracted: 2026-01-29 16:09:05

INSERT INTO tblDELETEDPeopleClientsTransportationServices ( IndexedName, RecordDeletedDate, RecordDeletedBy, Company, PhoneNumber, RouteNumber, FundingSource, ContractNumber, ActivityCode, ContractNumber2, ActivityCode2, DDSFunding, Comments, Inactive, DateInactive )
SELECT tblPeopleClientsTransportationServices.IndexedName, [Forms]![frmMainMenu]![TodaysDate] AS RecordDeletedDate, [Forms]![frmMainMenu]![UserName] AS RecordDeletedBy, tblPeopleClientsTransportationServices.Company, tblPeopleClientsTransportationServices.PhoneNumber, tblPeopleClientsTransportationServices.RouteNumber, tblPeopleClientsTransportationServices.FundingSource, tblPeopleClientsTransportationServices.ContractNumber, tblPeopleClientsTransportationServices.ActivityCode, tblPeopleClientsTransportationServices.ContractNumber2, tblPeopleClientsTransportationServices.ActivityCode2, tblPeopleClientsTransportationServices.DDSFunding, tblPeopleClientsTransportationServices.Comments, tblPeopleClientsTransportationServices.Inactive, tblPeopleClientsTransportationServices.DateInactive
FROM tblPeopleClientsTransportationServices
WHERE (((tblPeopleClientsTransportationServices.IndexedName)=[Forms]![frmMainMenu]![RememberIndexedName]));

