-- Query Name: qryEXPIRATIONS15
-- Extracted: 2026-01-29 16:09:05

INSERT INTO tblExpirations ( Location, RecordType, LastName, FirstName, Supervisor, DateISP, DateConsentFormsSigned, DateBMMExpires, DateBMMAccessSignedHRC, DateBMMAccessSigned, DateSPDAuthExpires, DateSignaturesDueBy, AllSPDSignaturesReceived )
SELECT DLookUp("GPName","temptbl","Location='" & [LocCLO] & "'") AS Location, 'Client' AS RecordType, temptbl0.LastName, temptbl0.FirstName, DLookUp("GPSuperCode","temptbl","Location='" & [LocCLO] & "'") AS Supervisor, temptbl0.DateISP, temptbl0.DateConsentFormsSigned, temptbl0.DateBMMExpires, temptbl0.DateBMMAccessSignedHRC, temptbl0.DateBMMAccessSigned, temptbl0.DateSPDAuthExpires, temptbl0.DateSignaturesDueBy, temptbl0.AllSPDSignaturesReceived
FROM temptbl0
WHERE (((DLookUp("GPName","temptbl","Location='" & [LocCLO] & "'")) Is Not Null) AND ((temptbl0.LastName) Is Not Null) AND ((temptbl0.FirstName) Is Not Null));

