-- Query Name: qryCensus
-- Extracted: 2026-01-29 16:09:06 (ADO Method)

SELECT tblLocations.CityTown, tblLocations.LocationName, tblLocations.Address, tblLocations.City, tblLocations.State, tblLocations.ZIP, tblLocations.County, tblLocations.NumClients
FROM tblLocations
WHERE (tblLocations.NumClients>0 AND tblLocations.Department='Residential Services') OR (tblLocations.City<>'Dedham' AND tblLocations.NumClients>0 AND tblLocations.Department='Individualized Support Options');

