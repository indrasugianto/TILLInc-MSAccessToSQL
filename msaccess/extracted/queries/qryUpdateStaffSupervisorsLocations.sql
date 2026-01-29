-- Query Name: qryUpdateStaffSupervisorsLocations
-- Extracted: 2026-01-29 16:09:05

UPDATE qrytblPeopleStaffSupervisors INNER JOIN tblPeople ON qrytblPeopleStaffSupervisors.INDEXEDNAME = tblPeople.IndexedName SET qrytblPeopleStaffSupervisors.LOCATION = tblPeople.OfficeCityTown & " - " & tblPeople.OfficeLocationName
WHERE (((tblPeople.OfficeCityTown) IS NOT NULL));

