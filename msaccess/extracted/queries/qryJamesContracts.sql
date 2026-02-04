-- Query Name: qryJamesContracts
-- Extracted: 2026-02-04 13:04:22 (ADO Method)

SELECT tblContracts.ContractID, tblContractsBillingBook.BIllingBookNumber, tblContractsBillingBook.ProgramName, tblContractsBillingBook.CostCenter, tblContractsBillingBook.MaximumObligation, tblContractsBillingBook.Units, tblContractsBillingBook.BillingRate, tblContractsBillingBook.NumberOfClients, tblContractsBillingBook.InternalRate, tblContractsBillingBook.FundingSource, tblContractsBillingBook.Staff
FROM tblContracts INNER JOIN tblContractsBillingBook ON (tblContracts.ContractID = tblContractsBillingBook.ContractID) AND (tblContracts.FY = tblContractsBillingBook.FY)
WHERE (((tblContracts.FY)=2023));

