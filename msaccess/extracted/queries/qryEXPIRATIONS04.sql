-- Query Name: qryEXPIRATIONS04
-- Extracted: 2026-02-04 13:04:22

INSERT INTO [~TempSuperCodes] ( GPCode, GPSuperCode, JobTitle )
SELECT tblStaff.DEPRTMNT, tblStaff.SUPERVISORCODE_I, tblStaff.JobTitle
FROM tblStaff
WHERE (((tblStaff.JobTitle)='RESUNT' Or (tblStaff.JobTitle)='RESUPR' Or (tblStaff.JobTitle)='ASDRRE' Or (tblStaff.JobTitle)='DASUPR' Or (tblStaff.JobTitle)='SENDPM')) OR (((tblStaff.DEPRTMNT)='CHELSE') AND ((tblStaff.JobTitle)='PRGMGR')) OR (((tblStaff.DEPRTMNT)='NEWTON') AND ((tblStaff.JobTitle)='PRGMGR')) OR (((tblStaff.JobTitle)='RESMGR')) OR (((tblStaff.JobTitle)='SITECO'))
ORDER BY tblStaff.DEPRTMNT;

