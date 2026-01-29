-- Query Name: qryTILLLocations
-- Extracted: 2026-01-29 16:09:06 (ADO Method)

SELECT tblLocations.CityTown, tblLocations.LocationName, tblLocations.Department, tblLocations.CostCenter, tblLocations.Cluster, tblLocations.ABI, tblLocations.NumClients, tblLocations.ResCapacity, tblLocations.ResTILLOwned, tblLocations.Address, tblLocations.City, tblLocations.State, tblLocations.ZIP, tblLocations.County, tblLocations.PhoneNumber
FROM tblLocations
WHERE (((tblLocations.Department)="Residential Services" Or (tblLocations.Department)="Day Services")) OR (((tblLocations.CityTown)="Charlestown")) OR (((tblLocations.CityTown)="Dedham") AND ((tblLocations.LocationName)="HQ")) OR (((tblLocations.CityTown)<>"Dedham") AND ((tblLocations.Department)="Vocational Services")) OR (((tblLocations.CityTown)<>"Dedham") AND ((tblLocations.Department)="Individualized Support Options")) OR (((tblLocations.CityTown)<>"Dedham") AND ((tblLocations.Department)="Day Services"))
ORDER BY tblLocations.CityTown, tblLocations.LocationName;

