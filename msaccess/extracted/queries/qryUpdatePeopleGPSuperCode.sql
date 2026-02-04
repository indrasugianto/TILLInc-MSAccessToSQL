-- Query Name: qryUpdatePeopleGPSuperCode
-- Extracted: 2026-02-04 13:04:22

UPDATE qrytblPeople INNER JOIN qrytblPeopleStaffSupervisors ON qrytblPeople.IndexedName = qrytblPeopleStaffSupervisors.INDEXEDNAME SET qrytblPeople.GPSuperCode = [qrytblPeopleStaffSupervisors].[SUPERVISORCODE_I];

