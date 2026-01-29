-- Query Name: qryCopyDeletedCommunityConnections
-- Extracted: 2026-01-29 16:09:05

INSERT INTO tblDELETEDPeopleClientsCommunityConnectionsServices ( IndexedName, RecordDeletedDate, RecordDeletedBy, Bowlers, BowlingTeam, Inactive, DateInactive )
SELECT tblPeopleClientsCommunityConnectionsServices.IndexedName, [Forms]![frmMainMenu]![TodaysDate] AS RecordDeletedDate, [Forms]![frmMainMenu]![UserName] AS RecordDeletedBy, tblPeopleClientsCommunityConnectionsServices.Bowlers, tblPeopleClientsCommunityConnectionsServices.BowlingTeam, tblPeopleClientsCommunityConnectionsServices.Inactive, tblPeopleClientsCommunityConnectionsServices.DateInactive
FROM tblPeopleClientsCommunityConnectionsServices
WHERE (((tblPeopleClientsCommunityConnectionsServices.IndexedName)=[Forms]![frmMainMenu]![RememberIndexedName]));

