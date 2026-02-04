-- Query Name: qryRefreshResidentialContacts
-- Extracted: 2026-02-04 13:04:22

UPDATE qrytblLocations LEFT JOIN qrytblPeople ON (qrytblLocations.CityTown = qrytblPeople.OfficeCityTown) AND (qrytblLocations.LocationName = qrytblPeople.OfficeLocationName) SET qrytblLocations.StaffPrimaryContactIndexedName = qrytblPeople.IndexedName, qrytblLocations.StaffPrimaryContactLastName = qrytblPeople.LastName, qrytblLocations.StaffPrimaryContactFirstName = qrytblPeople.FirstName
WHERE (((qrytblLocations.StaffPrimaryContactIndexedName)<>qrytblPeople.IndexedName) And ((qrytblLocations.Department)='Residential Services') And ((qrytblPeople.StaffTitle)='Residence Manager')) Or (((qrytblLocations.StaffPrimaryContactIndexedName)<>qrytblPeople.IndexedName) And ((qrytblLocations.Department)='Residential Services') And ((qrytblPeople.StaffTitle)="Site Coordinator"));

