-- Query Name: qryDDSMaxObligation
-- Extracted: 2026-02-04 13:04:21

SELECT tblContracts.ContractID, tblContracts.FY, IIf(Left([tblContracts]![ContractID],1)="1" Or Left([tblContracts]![ContractID],2)="IN","DDS","Non-DDS"), tblContracts.MaximumObligationAsAmended
FROM tblContracts
WHERE (((tblContracts.FY)=Forms!frmRptFinancialAndLetters!MRFY) And ((IIf(Left(tblContracts!ContractID,1)="1" Or Left(tblContracts!ContractID,2)="IN","DDS","Non-DDS"))="DDS"));

