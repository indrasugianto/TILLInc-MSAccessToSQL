-- Query Name: qryCurrentFYContracts
-- Extracted: 2026-02-04 13:04:21

SELECT tblContracts.*
FROM tblContracts
WHERE (((tblContracts.FY)=[Forms]![frmRptFinancialAndLetters]![MRFY]));

