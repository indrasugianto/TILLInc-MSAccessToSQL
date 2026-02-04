-- Query Name: qryEXPIRATIONS07
-- Extracted: 2026-02-04 13:04:22

UPDATE temptbl SET temptbl.GPSuperCode = DLookUp("GPSuperCode","~TempSuperCodes","GPCode='" & [temptbl].[GPName] & "'")
WHERE ((([temptbl].[GPSuperCode]) Is Null));

