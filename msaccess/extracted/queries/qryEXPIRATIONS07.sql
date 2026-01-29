-- Query Name: qryEXPIRATIONS07
-- Extracted: 2026-01-29 16:09:05

UPDATE temptbl SET temptbl.GPSuperCode = DLookUp("GPSuperCode","~TempSuperCodes","GPCode='" & [temptbl].[GPName] & "'")
WHERE ((([temptbl].[GPSuperCode]) Is Null));

