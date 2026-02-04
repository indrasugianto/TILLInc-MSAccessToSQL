-- Query Name: qryCopyDeletedNHRes
-- Extracted: 2026-02-04 13:04:21

INSERT INTO tblDELETEDPeopleClientsNHRes ( IndexedName, RecordDeletedDate, RecordDeletedBy, StartDate, EndDate, Funding, ResidentialRate, RoomAndBoard, ChargesForCare, ContractNumber, ActivityCode, ContractNumber2, ActivityCode2, Portion, CaseManager, Inactive, DateInactive )
SELECT tblDELETEDPeopleClientsNHRes.IndexedName, [Forms]![frmMainMenu]![TodaysDate] AS RecordDeletedDate, [Forms]![frmMainMenu]![UserName] AS RecordDeletedBy, tblDELETEDPeopleClientsNHRes.StartDate, tblDELETEDPeopleClientsNHRes.EndDate, tblDELETEDPeopleClientsNHRes.Funding, tblDELETEDPeopleClientsNHRes.ResidentialRate, tblDELETEDPeopleClientsNHRes.RoomAndBoard, tblDELETEDPeopleClientsNHRes.ChargesForCare, tblDELETEDPeopleClientsNHRes.ContractNumber, tblDELETEDPeopleClientsNHRes.ActivityCode, tblDELETEDPeopleClientsNHRes.ContractNumber2, tblDELETEDPeopleClientsNHRes.ActivityCode2, tblDELETEDPeopleClientsNHRes.Portion, tblDELETEDPeopleClientsNHRes.CaseManager, tblDELETEDPeopleClientsNHRes.Inactive, tblDELETEDPeopleClientsNHRes.DateInactive
FROM tblDELETEDPeopleClientsNHRes
WHERE (((tblDELETEDPeopleClientsNHRes.IndexedName)=[Forms]![frmMainMenu]![RememberIndexedName]));

