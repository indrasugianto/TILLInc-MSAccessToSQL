-- Query Name: qryDeleteSupervisorsWithNoStaff
-- Extracted: 2026-01-29 16:09:05

DELETE tblStaffDedhamManagers.*, tblPeopleStaffSupervisors.STAFFCOUNT
FROM tblPeopleStaffSupervisors INNER JOIN tblStaffDedhamManagers ON tblPeopleStaffSupervisors.SUPERVISORCODE_I = tblStaffDedhamManagers.SUPERVISORCODE_I
WHERE (((tblPeopleStaffSupervisors.STAFFCOUNT)=0));

