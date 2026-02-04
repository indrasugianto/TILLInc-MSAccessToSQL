-- Query Name: qryCensus
-- Extracted: 2026-02-04 13:04:22 (ADO Method)

SELECT tblLocations.CityTown, tblLocations.LocationName, tblLocations.Address, tblLocations.City, tblLocations.State, tblLocations.ZIP, tblLocations.County, tblLocations.NumClients
FROM tblLocations
WHERE (tblLocations.NumClients>0 AND tblLocations.Department='Residential Services') OR (tblLocations.City<>'Dedham' AND tblLocations.NumClients>0 AND tblLocations.Department='Individualized Support Options');

