-- Query Name: qryUpdatePeopleGPSuperCode
-- Extracted: 2026-01-29 16:09:05

UPDATE qrytblPeople INNER JOIN qrytblPeopleStaffSupervisors ON qrytblPeople.IndexedName = qrytblPeopleStaffSupervisors.INDEXEDNAME SET qrytblPeople.GPSuperCode = [qrytblPeopleStaffSupervisors].[SUPERVISORCODE_I];

