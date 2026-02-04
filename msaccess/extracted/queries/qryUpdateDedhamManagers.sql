-- Query Name: qryUpdateDedhamManagers
-- Extracted: 2026-02-04 13:04:22

INSERT INTO tblStaffDedhamManagers ( SUPERVISORCODE_I, SupervisorName, NewLocation, Location, IndexedName )
SELECT qrytblPeopleStaffSupervisors.SUPERVISORCODE_I, tblPeople.FirstName & " " & tblPeople.LastName AS SupervisorName, "DED-" & Left(tblPeople.FirstName,1) & Left(tblPeople.LastName,1) AS NewLocation, "Dedham - " & tblpeople.FirstName & ' ' & tblpeople.LastName & " Direct Reports" AS Location, tblPeople.IndexedName AS IndexedName
FROM qrytblPeopleStaffSupervisors INNER JOIN tblPeople ON qrytblPeopleStaffSupervisors.INDEXEDNAME = tblPeople.IndexedName
WHERE (((qrytblPeopleStaffSupervisors.LOCATION)="Dedham - HQ") AND ((qrytblPeopleStaffSupervisors.STAFFCOUNT)>0));

