-- Query Name: qryUpdateStaffSupervisors
-- Extracted: 2026-02-04 13:04:22

UPDATE (tblPeople INNER JOIN tblStaff ON (tblPeople.LastName = tblStaff.LASTNAME) AND (tblPeople.FirstName = tblStaff.FRSTNAME)) INNER JOIN tblPeopleStaffSupervisors ON tblStaff.SUPERVISORCODE_I = tblPeopleStaffSupervisors.SUPERVISORCODE_I SET tblPeople.ManagerSuperCode = [tblPeopleStaffSupervisors]![SUPERVISORCODE_I]
WHERE tblPeople.isStaff=True;

