-- Query Name: qryCurrentFYContracts
-- Extracted: 2026-01-29 16:09:05

SELECT tblContracts.*
FROM tblContracts
WHERE (((tblContracts.FY)=[Forms]![frmRptFinancialAndLetters]![MRFY]));

