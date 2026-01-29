-- Query Name: qryUpdateStaffSupervisors
-- Extracted: 2026-01-29 16:09:05

UPDATE (tblPeople INNER JOIN tblStaff ON (tblPeople.LastName = tblStaff.LASTNAME) AND (tblPeople.FirstName = tblStaff.FRSTNAME)) INNER JOIN tblPeopleStaffSupervisors ON tblStaff.SUPERVISORCODE_I = tblPeopleStaffSupervisors.SUPERVISORCODE_I SET tblPeople.ManagerSuperCode = [tblPeopleStaffSupervisors]![SUPERVISORCODE_I]
WHERE tblPeople.isStaff=True;

