-- Query Name: qryRefreshResidentialContacts
-- Extracted: 2026-01-29 16:09:05

UPDATE qrytblLocations LEFT JOIN qrytblPeople ON (qrytblLocations.CityTown = qrytblPeople.OfficeCityTown) AND (qrytblLocations.LocationName = qrytblPeople.OfficeLocationName) SET qrytblLocations.StaffPrimaryContactIndexedName = qrytblPeople.IndexedName, qrytblLocations.StaffPrimaryContactLastName = qrytblPeople.LastName, qrytblLocations.StaffPrimaryContactFirstName = qrytblPeople.FirstName
WHERE (((qrytblLocations.StaffPrimaryContactIndexedName)<>qrytblPeople.IndexedName) And ((qrytblLocations.Department)='Residential Services') And ((qrytblPeople.StaffTitle)='Residence Manager')) Or (((qrytblLocations.StaffPrimaryContactIndexedName)<>qrytblPeople.IndexedName) And ((qrytblLocations.Department)='Residential Services') And ((qrytblPeople.StaffTitle)="Site Coordinator"));

