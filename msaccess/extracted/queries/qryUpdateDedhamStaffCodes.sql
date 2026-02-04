-- Query Name: qryUpdateDedhamStaffCodes
-- Extracted: 2026-02-04 13:04:22

UPDATE qrytblStaffDedhamManagers INNER JOIN tblLocations ON qrytblStaffDedhamManagers.IndexedName = tblLocations.StaffPrimaryContactIndexedName SET tblLocations.GPName = [qrytblStaffDedhamManagers].[NewLocation]
WHERE (((tblLocations.CityTown)="HQ"));

