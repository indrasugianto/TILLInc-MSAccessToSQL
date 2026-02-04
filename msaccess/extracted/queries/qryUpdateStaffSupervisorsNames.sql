-- Query Name: qryUpdateStaffSupervisorsNames
-- Extracted: 2026-02-04 13:04:22

UPDATE qrytblPeopleStaffSupervisors INNER JOIN qrytblStaff ON qrytblPeopleStaffSupervisors.SUPEMPLID = qrytblStaff.EMPLOYID SET qrytblPeopleStaffSupervisors.LASTNAME = StrConv (Trim([qrytblStaff].[LASTNAME]), 3), qrytblPeopleStaffSupervisors.FIRSTNAME = StrConv (Trim([qrytblStaff].[FRSTNAME]), 3), qrytblPeopleStaffSupervisors.INDEXEDNAME = StrConv (Trim([qrytblStaff].[LASTNAME]), 3) & '/' & StrConv (Trim([qrytblStaff].[FRSTNAME]), 3) & '//';

