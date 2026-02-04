-- Query Name: qryResClientsBySite
-- Extracted: 2026-02-04 13:04:22 (ADO Method)

SELECT tblLocations.Cluster, tblLocations.CityTown, tblLocations.LocationName, tblLocations.NumClients, tblLocations.ResCapacity, [ResCapacity]-[NumClients] AS NumVacancies, tblPeople.LastName, tblPeople.FirstName, tblPeopleClientsDemographics.UsesWheelchair, tblPeopleClientsDemographics.UsesWalker
FROM ((tblPeople INNER JOIN tblPeopleClientsResidentialServices ON tblPeople.IndexedName = tblPeopleClientsResidentialServices.IndexedName) INNER JOIN tblLocations ON (tblPeopleClientsResidentialServices.Location = tblLocations.LocationName) AND (tblPeopleClientsResidentialServices.CityTown = tblLocations.CityTown)) LEFT JOIN tblPeopleClientsDemographics ON tblPeople.IndexedName = tblPeopleClientsDemographics.IndexedName
WHERE (((tblPeople.IsClientRes)=True) AND ((tblPeople.IsDeceased)=False) AND ((tblPeopleClientsResidentialServices.Inactive)=False))
ORDER BY tblLocations.Cluster, tblLocations.CityTown, tblLocations.LocationName, tblPeople.LastName, tblPeople.FirstName;

