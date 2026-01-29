-- Query Name: qryRptSeverityRates
-- Extracted: 2026-01-29 16:09:05

UPDATE temptbl SET temptbl.Rate = DLookUp("MedicaidRate","catSeverityRates","Severity='" & [temptbl].[Sev] & "'");

