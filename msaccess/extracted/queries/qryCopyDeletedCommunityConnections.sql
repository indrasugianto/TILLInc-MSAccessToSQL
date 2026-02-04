-- Query Name: qryCopyDeletedCommunityConnections
-- Extracted: 2026-02-04 13:04:21

INSERT INTO tblDELETEDPeopleClientsCommunityConnectionsServices ( IndexedName, RecordDeletedDate, RecordDeletedBy, Bowlers, BowlingTeam, Inactive, DateInactive )
SELECT tblPeopleClientsCommunityConnectionsServices.IndexedName, [Forms]![frmMainMenu]![TodaysDate] AS RecordDeletedDate, [Forms]![frmMainMenu]![UserName] AS RecordDeletedBy, tblPeopleClientsCommunityConnectionsServices.Bowlers, tblPeopleClientsCommunityConnectionsServices.BowlingTeam, tblPeopleClientsCommunityConnectionsServices.Inactive, tblPeopleClientsCommunityConnectionsServices.DateInactive
FROM tblPeopleClientsCommunityConnectionsServices
WHERE (((tblPeopleClientsCommunityConnectionsServices.IndexedName)=[Forms]![frmMainMenu]![RememberIndexedName]));

