-- Query Name: qryEXPIRATIONS02
-- Extracted: 2026-01-29 16:09:05

UPDATE qrytblStaffDedhamManagers INNER JOIN tempstaff ON qrytblStaffDedhamManagers.SUPERVISORCODE_I = tempstaff.SUPERVISORCODE_I SET tempstaff.DIVISIONCODE_I = 'DEDHAM', tempstaff.DEPRTMNT = qrytblStaffDedhamManagers.NewLocation;

