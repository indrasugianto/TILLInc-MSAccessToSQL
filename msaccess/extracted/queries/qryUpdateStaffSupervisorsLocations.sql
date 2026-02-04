-- Query Name: qryUpdateStaffSupervisorsLocations
-- Extracted: 2026-02-04 13:04:22

UPDATE qrytblPeopleStaffSupervisors INNER JOIN tblPeople ON qrytblPeopleStaffSupervisors.INDEXEDNAME = tblPeople.IndexedName SET qrytblPeopleStaffSupervisors.LOCATION = tblPeople.OfficeCityTown & " - " & tblPeople.OfficeLocationName
WHERE (((tblPeople.OfficeCityTown) IS NOT NULL));

