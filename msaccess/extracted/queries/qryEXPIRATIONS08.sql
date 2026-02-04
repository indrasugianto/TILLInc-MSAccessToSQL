-- Query Name: qryEXPIRATIONS08
-- Extracted: 2026-02-04 13:04:22

INSERT INTO temptbl ( Location, CityTown, LocationName, GPName, GPSuperCode )
SELECT [CityTown] & " - " & [LocationName] AS Location, tblLocations.CityTown, tblLocations.LocationName, tblLocations.GPName, DLookUp("GPSuperCode","tblPeople","FirstName='" & tblLocations.StaffPrimaryContactFirstName & "' AND LastName='" & tblLocations.StaffPrimaryContactLastName & "'") AS Expr1
FROM tblLocations
WHERE (((tblLocations.CityTown)<>"Dedham") AND ((tblLocations.GPName) Is Not Null) AND ((tblLocations.Department)="Individualized Support Options"))
ORDER BY [CityTown] & " - " & [LocationName];

