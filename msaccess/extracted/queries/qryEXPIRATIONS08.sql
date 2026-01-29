-- Query Name: qryEXPIRATIONS08
-- Extracted: 2026-01-29 16:09:05

INSERT INTO temptbl ( Location, CityTown, LocationName, GPName, GPSuperCode )
SELECT [CityTown] & " - " & [LocationName] AS Location, tblLocations.CityTown, tblLocations.LocationName, tblLocations.GPName, DLookUp("GPSuperCode","tblPeople","FirstName='" & tblLocations.StaffPrimaryContactFirstName & "' AND LastName='" & tblLocations.StaffPrimaryContactLastName & "'") AS Expr1
FROM tblLocations
WHERE (((tblLocations.CityTown)<>"Dedham") AND ((tblLocations.GPName) Is Not Null) AND ((tblLocations.Department)="Individualized Support Options"))
ORDER BY [CityTown] & " - " & [LocationName];

