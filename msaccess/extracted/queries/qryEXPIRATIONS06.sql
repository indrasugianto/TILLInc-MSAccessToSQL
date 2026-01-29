-- Query Name: qryEXPIRATIONS06
-- Extracted: 2026-01-29 16:09:05

SELECT [CityTown] & " - " & [LocationName] AS Location, tblLocations.CityTown, tblLocations.LocationName, tblLocations.GPName, tblPeople.GPSuperCode INTO temptbl
FROM tblLocations INNER JOIN tblPeople ON (tblLocations.CityTown = tblPeople.OfficeCityTown) AND (tblLocations.LocationName = tblPeople.OfficeLocationName)
WHERE (((tblLocations.GPName) Is Not Null) AND ((tblPeople.IsStaff)=True))
ORDER BY [CityTown] & "-" & [LocationName];

