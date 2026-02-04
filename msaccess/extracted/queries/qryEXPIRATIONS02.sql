-- Query Name: qryEXPIRATIONS02
-- Extracted: 2026-02-04 13:04:22

UPDATE qrytblStaffDedhamManagers INNER JOIN tempstaff ON qrytblStaffDedhamManagers.SUPERVISORCODE_I = tempstaff.SUPERVISORCODE_I SET tempstaff.DIVISIONCODE_I = 'DEDHAM', tempstaff.DEPRTMNT = qrytblStaffDedhamManagers.NewLocation;

