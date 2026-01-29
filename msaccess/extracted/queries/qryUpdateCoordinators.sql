-- Query Name: qryUpdateCoordinators
-- Extracted: 2026-01-29 16:09:05

UPDATE (tblLocations INNER JOIN catClusters ON tblLocations.Cluster = catClusters.ClusterID) INNER JOIN tblPeople ON (catClusters.ClusterManagerFirstName = tblPeople.FirstName) AND (catClusters.ClusterManagerLastName = tblPeople.LastName) SET tblLocations.StaffSecondaryContactIndexedName = [tblPeople].[IndexedName], tblLocations.StaffSecondaryContactLastName = [catClusters].[ClusterManagerLastName], tblLocations.StaffSecondaryContactFirstName = [catClusters].[ClusterManagerFirstName], tblLocations.StaffSecondaryContactMiddleInitial = [catClusters].[ClusterManagerMiddleInitial]
WHERE (((tblLocations.Department)="Residential Services"));

