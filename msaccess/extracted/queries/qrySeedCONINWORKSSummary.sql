-- Query Name: qrySeedCONINWORKSSummary
-- Extracted: 2026-02-04 13:04:22

INSERT INTO [~CONINWORKSSummary] ( FY, ContractID, DDSMaxObligation, ContractUnit, NumLocations, NumClients, Units, BillingRate, FundingSource, AccountingStaff )
SELECT qryCurrentFYContracts.FY, qryCurrentFYContracts.ContractID, qryCurrentFYContracts.MaximumObligationAsAmended, qryCurrentFYContracts.Units, DCount("ContractID","qryCONINWORKS","ContractID=""" & [qryCurrentFYContracts].[ContractID] & """") AS NumLocations, qryCurrentFYContracts.TotalClients, qryCurrentFYContracts.TotalUnitsAsAmended, DLookUp("BillingRate","tblContractsBillingBook","FY=" & [Forms]![frmMainMenu]![SelectFY] & " AND ContractID=""" & [qryCurrentFYContracts].[ContractID] & """") AS BillingRate, DLookUp("FundingSource","qryCONINWORKS","ContractID=""" & [qryCurrentFYContracts].[ContractID] & """") AS FundingSource, qryCurrentFYContracts.AccountingStaff
FROM qryCurrentFYContracts
WHERE (((IIf(Left([ContractID],1)="1",True,IIf(Left([ContractID],2)="IN",True,False)))=True));

