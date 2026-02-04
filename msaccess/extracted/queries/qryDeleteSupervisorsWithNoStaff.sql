-- Query Name: qryDeleteSupervisorsWithNoStaff
-- Extracted: 2026-02-04 13:04:22

DELETE tblStaffDedhamManagers.*, tblPeopleStaffSupervisors.STAFFCOUNT
FROM tblPeopleStaffSupervisors INNER JOIN tblStaffDedhamManagers ON tblPeopleStaffSupervisors.SUPERVISORCODE_I = tblStaffDedhamManagers.SUPERVISORCODE_I
WHERE (((tblPeopleStaffSupervisors.STAFFCOUNT)=0));

