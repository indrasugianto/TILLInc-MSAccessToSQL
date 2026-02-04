-- Query Name: qryUpdateAsstDirectors
-- Extracted: 2026-02-04 13:04:22

UPDATE (tblLocations INNER JOIN catClusters ON tblLocations.Cluster = catClusters.ClusterID) INNER JOIN tblPeople ON (catClusters.ClusterDirectorFirstName = tblPeople.FirstName) AND (catClusters.ClusterDirectorLastName = tblPeople.LastName) SET tblLocations.StaffTertiaryContactIndexedName = [tblPeople].[IndexedName], tblLocations.StaffTertiaryContactLastName = [catClusters].[ClusterDirectorLastName], tblLocations.StaffTertiaryContactFirstName = [catClusters].[ClusterDirectorFirstName], tblLocations.StaffTertiaryContactMiddleInitial = [catClusters].[ClusterDirectorMiddleInitial]
WHERE (((tblLocations.Department)="Residential Services"));

