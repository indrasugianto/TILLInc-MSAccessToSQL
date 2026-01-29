-- Query Name: qryDDSMaxObligation
-- Extracted: 2026-01-29 16:09:05

SELECT tblContracts.ContractID, tblContracts.FY, IIf(Left([tblContracts]![ContractID],1)="1" Or Left([tblContracts]![ContractID],2)="IN","DDS","Non-DDS"), tblContracts.MaximumObligationAsAmended
FROM tblContracts
WHERE (((tblContracts.FY)=Forms!frmRptFinancialAndLetters!MRFY) And ((IIf(Left(tblContracts!ContractID,1)="1" Or Left(tblContracts!ContractID,2)="IN","DDS","Non-DDS"))="DDS"));

