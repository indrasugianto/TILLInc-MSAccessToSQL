-- Query Name: qryCopyDeletedSpringboard
-- Extracted: 2026-02-04 13:04:21

INSERT INTO tblDELETEDPeopleClientsSpringboardServices ( IndexedName, RecordDeletedDate, RecordDeletedBy, CustomerID, Age, DateJoined, DateTerminated, ReasonForTermination, BeginBillingDate, GroupCode, LeaderIndexedName, Leader, Inactive, DateInactive )
SELECT tblPeopleClientsSpringboardServices.IndexedName, [Forms]![frmMainMenu]![TodaysDate] AS RecordDeletedDate, [Forms]![frmMainMenu]![UserName] AS RecordDeletedBy, tblPeopleClientsSpringboardServices.CustomerID, tblPeopleClientsSpringboardServices.Age, tblPeopleClientsSpringboardServices.DateJoined, tblPeopleClientsSpringboardServices.DateTerminated, tblPeopleClientsSpringboardServices.ReasonForTermination, tblPeopleClientsSpringboardServices.BeginBillingDate, tblPeopleClientsSpringboardServices.GroupCode, tblPeopleClientsSpringboardServices.LeaderIndexedName, tblPeopleClientsSpringboardServices.Leader, tblPeopleClientsSpringboardServices.Inactive, tblPeopleClientsSpringboardServices.DateInactive
FROM tblPeopleClientsSpringboardServices
WHERE (((tblPeopleClientsSpringboardServices.IndexedName)=[Forms]![frmMainMenu]![RememberIndexedName]));

