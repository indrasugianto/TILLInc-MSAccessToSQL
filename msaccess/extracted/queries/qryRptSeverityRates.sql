-- Query Name: qryRptSeverityRates
-- Extracted: 2026-02-04 13:04:22

UPDATE temptbl SET temptbl.Rate = DLookUp("MedicaidRate","catSeverityRates","Severity='" & [temptbl].[Sev] & "'");

