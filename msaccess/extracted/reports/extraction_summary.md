# MS Access Database Extraction Report

**Database:** `c:\GitHub\TILLInc-MSAccessToSQL\msaccess\TILLDB_V9.14_20260203d - WEB.accdb`  
**Extraction Date:** 2026-02-04 13:04:22

## Connection Information
- **Server:** tillsqlserver.database.windows.net
- **User:** tillsqladmin
- **Database Type:** Azure SQL Database

## Summary
- **Total Tables Extracted:** 48
- **Total Queries Extracted:** 167
- **VBA Modules:** Not extracted (requires MS Access application)

## Tables Extracted

| Table Name | Type | Columns | Output File |
|------------|------|---------|-------------|
| catLogonAs | TABLE | 1 | catLogonAs_schema.txt |
| qryappParameters | VIEW | 3 | qryappParameters_schema.txt |
| qryAutismIndividualsAndFamilyForDDS | VIEW | 27 | qryAutismIndividualsAndFamilyForDDS_schema.txt |
| qrycatClusters | VIEW | 12 | qrycatClusters_schema.txt |
| qryCensus | VIEW | 8 | qryCensus_schema.txt |
| qryClientGovernmentAccounts | VIEW | 8 | qryClientGovernmentAccounts_schema.txt |
| qryDonations | VIEW | 13 | qryDonations_schema.txt |
| qryJamesContracts | VIEW | 11 | qryJamesContracts_schema.txt |
| qryJamesExport | VIEW | 14 | qryJamesExport_schema.txt |
| qryPeopleAllPeopleRecords | VIEW | 86 | qryPeopleAllPeopleRecords_schema.txt |
| qryPeopleTILLGamesEmail | VIEW | 5 | qryPeopleTILLGamesEmail_schema.txt |
| qryResClientsBySite | VIEW | 10 | qryResClientsBySite_schema.txt |
| qrySECTION8 | VIEW | 27 | qrySECTION8_schema.txt |
| qrySECTION8_Original | VIEW | 27 | qrySECTION8_Original_schema.txt |
| qrySTAFFWITHNOSKILLS | VIEW | 7 | qrySTAFFWITHNOSKILLS_schema.txt |
| qrySTAFFWITHNOSKILLS_Original | VIEW | 7 | qrySTAFFWITHNOSKILLS_Original_schema.txt |
| qrytblContracts | VIEW | 20 | qrytblContracts_schema.txt |
| qrytblContractsAmendments | VIEW | 24 | qrytblContractsAmendments_schema.txt |
| qrytblContractsBillingBook | VIEW | 48 | qrytblContractsBillingBook_schema.txt |
| qrytblExpirations | VIEW | 48 | qrytblExpirations_schema.txt |
| qrytblLocations | VIEW | 69 | qrytblLocations_schema.txt |
| qrytblPeople | VIEW | 91 | qrytblPeople_schema.txt |
| qrytblPeopleClientsAFCServices | VIEW | 16 | qrytblPeopleClientsAFCServices_schema.txt |
| qrytblPeopleClientsAutismServices | VIEW | 23 | qrytblPeopleClientsAutismServices_schema.txt |
| qrytblPeopleClientsCLOServices | VIEW | 47 | qrytblPeopleClientsCLOServices_schema.txt |
| qrytblPeopleClientsDayServices | VIEW | 51 | qrytblPeopleClientsDayServices_schema.txt |
| qrytblPeopleClientsDemographics | VIEW | 121 | qrytblPeopleClientsDemographics_schema.txt |
| qrytblPeopleClientsIndividualSupportServices | VIEW | 20 | qrytblPeopleClientsIndividualSupportServices_schema.txt |
| qrytblPeopleClientsPCAServices | VIEW | 38 | qrytblPeopleClientsPCAServices_schema.txt |
| qrytblPeopleClientsPCAServicesContactNotes | VIEW | 14 | qrytblPeopleClientsPCAServicesContactNotes_schema.txt |
| qrytblPeopleClientsResidentialServices | VIEW | 53 | qrytblPeopleClientsResidentialServices_schema.txt |
| qrytblPeopleClientsSharedLivingServices | VIEW | 19 | qrytblPeopleClientsSharedLivingServices_schema.txt |
| qrytblPeopleClientsSpringboardServices | VIEW | 16 | qrytblPeopleClientsSpringboardServices_schema.txt |
| qrytblPeopleClientsVendors | VIEW | 58 | qrytblPeopleClientsVendors_schema.txt |
| qrytblPeopleClientsVocationalServices | VIEW | 20 | qrytblPeopleClientsVocationalServices_schema.txt |
| qrytblPeopleConsultants | VIEW | 10 | qrytblPeopleConsultants_schema.txt |
| qrytblPeopleDonors | VIEW | 26 | qrytblPeopleDonors_schema.txt |
| qrytblPeopleFamily | VIEW | 16 | qrytblPeopleFamily_schema.txt |
| qrytblPeopleStaffSupervisors | VIEW | 9 | qrytblPeopleStaffSupervisors_schema.txt |
| qrytblStaff | VIEW | 13 | qrytblStaff_schema.txt |
| qrytblStaffDedhamManagers | VIEW | 5 | qrytblStaffDedhamManagers_schema.txt |
| qrytblStaffEvalsAndSupervisions | VIEW | 15 | qrytblStaffEvalsAndSupervisions_schema.txt |
| qrytblStaffSkills | VIEW | 4 | qrytblStaffSkills_schema.txt |
| qryTILLLocations | VIEW | 15 | qryTILLLocations_schema.txt |
| RedReport | TABLE | 13 | RedReport_schema.txt |
| tblDBChangeHistoryOLD | TABLE | 3 | tblDBChangeHistoryOLD_schema.txt |
| tblLocationsDedhamManagers | TABLE | 67 | tblLocationsDedhamManagers_schema.txt |
| tblTILLMonthlyMasterArchive | TABLE | 147 | tblTILLMonthlyMasterArchive_schema.txt |

## Queries Extracted

| Query Name | Output File |
|------------|-------------|
| qryAppendDeletedPerson | qryAppendDeletedPerson.sql |
| qryAutismServicesRawDemographics_Crosstab | qryAutismServicesRawDemographics_Crosstab.sql |
| qryClientFamilyMailing | qryClientFamilyMailing.sql |
| qryCONINWORKS | qryCONINWORKS.sql |
| qryCONINWORKSExport | qryCONINWORKSExport.sql |
| qryCopyDeletedAdultCoaching | qryCopyDeletedAdultCoaching.sql |
| qryCopyDeletedAdultCompanion | qryCopyDeletedAdultCompanion.sql |
| qryCopyDeletedAutism | qryCopyDeletedAutism.sql |
| qryCopyDeletedClientDemographics | qryCopyDeletedClientDemographics.sql |
| qryCopyDeletedClientVendors | qryCopyDeletedClientVendors.sql |
| qryCopyDeletedCLO | qryCopyDeletedCLO.sql |
| qryCopyDeletedCommunityConnections | qryCopyDeletedCommunityConnections.sql |
| qryCopyDeletedDay | qryCopyDeletedDay.sql |
| qryCopyDeletedDonors | qryCopyDeletedDonors.sql |
| qryCopyDeletedFamilyByClientName | qryCopyDeletedFamilyByClientName.sql |
| qryCopyDeletedFamilyByFamilyName | qryCopyDeletedFamilyByFamilyName.sql |
| qryCopyDeletedISS | qryCopyDeletedISS.sql |
| qryCopyDeletedNHDay | qryCopyDeletedNHDay.sql |
| qryCopyDeletedNHRes | qryCopyDeletedNHRes.sql |
| qryCopyDeletedPCA | qryCopyDeletedPCA.sql |
| qryCopyDeletedPCAContactNotes | qryCopyDeletedPCAContactNotes.sql |
| qryCopyDeletedResidential | qryCopyDeletedResidential.sql |
| qryCopyDeletedSharedLiving | qryCopyDeletedSharedLiving.sql |
| qryCopyDeletedSpringboard | qryCopyDeletedSpringboard.sql |
| qryCopyDeletedTransportation | qryCopyDeletedTransportation.sql |
| qryCopyDeletedTRASE | qryCopyDeletedTRASE.sql |
| qryCopyDeletedVocational | qryCopyDeletedVocational.sql |
| qryCreateAllPeopleTable | qryCreateAllPeopleTable.sql |
| qryCreateTempCompaniesLookup | qryCreateTempCompaniesLookup.sql |
| qryCurrentFYContracts | qryCurrentFYContracts.sql |
| qryDDSMaxObligation | qryDDSMaxObligation.sql |
| qryDeleteCorruptedPeopleRecords | qryDeleteCorruptedPeopleRecords.sql |
| qryDeleteNullPeople | qryDeleteNullPeople.sql |
| qryDeletePhoneDirectory | qryDeletePhoneDirectory.sql |
| qryDeleteSkillsNotTracked | qryDeleteSkillsNotTracked.sql |
| qryDeleteStaffNoDivisionCode | qryDeleteStaffNoDivisionCode.sql |
| qryDeleteSupervisorsWithNoStaff | qryDeleteSupervisorsWithNoStaff.sql |
| qryDonationsForExport | qryDonationsForExport.sql |
| qryDonorAppealCreateMostRecentDonations | qryDonorAppealCreateMostRecentDonations.sql |
| qryDonorsMonthlyReport | qryDonorsMonthlyReport.sql |
| qryEXPIRATIONS00 | qryEXPIRATIONS00.sql |
| qryEXPIRATIONS01 | qryEXPIRATIONS01.sql |
| qryEXPIRATIONS02 | qryEXPIRATIONS02.sql |
| qryEXPIRATIONS02A | qryEXPIRATIONS02A.sql |
| qryEXPIRATIONS03 | qryEXPIRATIONS03.sql |
| qryEXPIRATIONS03A | qryEXPIRATIONS03A.sql |
| qryEXPIRATIONS04 | qryEXPIRATIONS04.sql |
| qryEXPIRATIONS04A | qryEXPIRATIONS04A.sql |
| qryEXPIRATIONS05 | qryEXPIRATIONS05.sql |
| qryEXPIRATIONS05A | qryEXPIRATIONS05A.sql |
| qryEXPIRATIONS05B | qryEXPIRATIONS05B.sql |
| qryEXPIRATIONS06 | qryEXPIRATIONS06.sql |
| qryEXPIRATIONS07 | qryEXPIRATIONS07.sql |
| qryEXPIRATIONS08 | qryEXPIRATIONS08.sql |
| qryEXPIRATIONS09 | qryEXPIRATIONS09.sql |
| qryEXPIRATIONS10 | qryEXPIRATIONS10.sql |
| qryEXPIRATIONS11 | qryEXPIRATIONS11.sql |
| qryEXPIRATIONS12 | qryEXPIRATIONS12.sql |
| qryEXPIRATIONS13 | qryEXPIRATIONS13.sql |
| qryEXPIRATIONS14 | qryEXPIRATIONS14.sql |
| qryEXPIRATIONS15 | qryEXPIRATIONS15.sql |
| qryEXPIRATIONS16 | qryEXPIRATIONS16.sql |
| qryEXPIRATIONS17 | qryEXPIRATIONS17.sql |
| qryEXPIRATIONS18 | qryEXPIRATIONS18.sql |
| qryEXPIRATIONS19 | qryEXPIRATIONS19.sql |
| qryEXPIRATIONS20 | qryEXPIRATIONS20.sql |
| qryEXPIRATIONS21 | qryEXPIRATIONS21.sql |
| qryEXPIRATIONS22 | qryEXPIRATIONS22.sql |
| qryEXPIRATIONS23 | qryEXPIRATIONS23.sql |
| qryEXPIRATIONS24 | qryEXPIRATIONS24.sql |
| qryEXPIRATIONS25 | qryEXPIRATIONS25.sql |
| qryEXPIRATIONS26 | qryEXPIRATIONS26.sql |
| qryExpirationsStaffBySkills | qryExpirationsStaffBySkills.sql |
| qryExpirationsStaffCull | qryExpirationsStaffCull.sql |
| qryISPInfo | qryISPInfo.sql |
| qryLoadConsultantsTable | qryLoadConsultantsTable.sql |
| qryLoadFormDemographics | qryLoadFormDemographics.sql |
| qryLoadFormVendors | qryLoadFormVendors.sql |
| qryLoadPhoneDirectory | qryLoadPhoneDirectory.sql |
| qryLoadTemporaryFamilyTable | qryLoadTemporaryFamilyTable.sql |
| qryMSAMISC | qryMSAMISC.sql |
| qryMSAMISCTEMP | qryMSAMISCTEMP.sql |
| qryPeopleAddFamilyRepPayee | qryPeopleAddFamilyRepPayee.sql |
| qryPeopleRemoveFamilyRepPayee | qryPeopleRemoveFamilyRepPayee.sql |
| qryPeopleResidentialClientsReportExport | qryPeopleResidentialClientsReportExport.sql |
| qryProgramExpirations | qryProgramExpirations.sql |
| qryRefreshResidentialContacts | qryRefreshResidentialContacts.sql |
| qryRepairAnomalies1 | qryRepairAnomalies1.sql |
| qryRepairAnomalies2 | qryRepairAnomalies2.sql |
| qryRepairAnomalies3 | qryRepairAnomalies3.sql |
| qryRepairAnomalies4 | qryRepairAnomalies4.sql |
| qryRptSeverityRates | qryRptSeverityRates.sql |
| qryRUNREPORT05 | qryRUNREPORT05.sql |
| qrySeedCONINWORKSSummary | qrySeedCONINWORKSSummary.sql |
| qryStaffAndEvalsDeleteInactives | qryStaffAndEvalsDeleteInactives.sql |
| qrySurvey | qrySurvey.sql |
| qryUpdateAllPeopleTable | qryUpdateAllPeopleTable.sql |
| qryUpdateAsstDirectors | qryUpdateAsstDirectors.sql |
| qryUpdateCoordinators | qryUpdateCoordinators.sql |
| qryUpdateDedhamHQLocations | qryUpdateDedhamHQLocations.sql |
| qryUpdateDedhamManagers | qryUpdateDedhamManagers.sql |
| qryUpdateDedhamStaffCodes | qryUpdateDedhamStaffCodes.sql |
| qryUpdatePeopleGPSuperCode | qryUpdatePeopleGPSuperCode.sql |
| qryUpdateSpringboardClientsStep1 | qryUpdateSpringboardClientsStep1.sql |
| qryUpdateSpringboardClientsStep2 | qryUpdateSpringboardClientsStep2.sql |
| qryUpdateSpringboardClientsStep3 | qryUpdateSpringboardClientsStep3.sql |
| qryUpdateSpringboardClientsStep4 | qryUpdateSpringboardClientsStep4.sql |
| qryUpdateSpringboardClientsStep5 | qryUpdateSpringboardClientsStep5.sql |
| qryUpdateSpringboardClientsStep6 | qryUpdateSpringboardClientsStep6.sql |
| qryUpdateSpringboardLeaders01 | qryUpdateSpringboardLeaders01.sql |
| qryUpdateSpringboardLeaders02 | qryUpdateSpringboardLeaders02.sql |
| qryUpdateSpringboardLeaders03 | qryUpdateSpringboardLeaders03.sql |
| qryUpdateSpringboardLeaders04 | qryUpdateSpringboardLeaders04.sql |
| qryUpdateSpringboardLeaders05 | qryUpdateSpringboardLeaders05.sql |
| qryUpdateSpringboardLeaders06 | qryUpdateSpringboardLeaders06.sql |
| qryUpdateStaffSupervisors | qryUpdateStaffSupervisors.sql |
| qryUpdateStaffSupervisorsLocations | qryUpdateStaffSupervisorsLocations.sql |
| qryUpdateStaffSupervisorsNames | qryUpdateStaffSupervisorsNames.sql |
| qryUpdateStaffTable | qryUpdateStaffTable.sql |
| Query1 | Query1.sql |
| ~Ad-hoc CLO | _Ad-hoc CLO.sql |
| ~Ad-hoc Day | _Ad-hoc Day.sql |
| ~Ad-hoc Vocational | _Ad-hoc Vocational.sql |
| ~Ad-hoc-Residential | _Ad-hoc-Residential.sql |
| qryappParameters | qryappParameters.sql |
| qryAutismIndividualsAndFamilyForDDS | qryAutismIndividualsAndFamilyForDDS.sql |
| qrycatClusters | qrycatClusters.sql |
| qryCensus | qryCensus.sql |
| qryClientGovernmentAccounts | qryClientGovernmentAccounts.sql |
| qryDonations | qryDonations.sql |
| qryJamesContracts | qryJamesContracts.sql |
| qryJamesExport | qryJamesExport.sql |
| qryPeopleAllPeopleRecords | qryPeopleAllPeopleRecords.sql |
| qryPeopleTILLGamesEmail | qryPeopleTILLGamesEmail.sql |
| qryResClientsBySite | qryResClientsBySite.sql |
| qrySECTION8 | qrySECTION8.sql |
| qrySECTION8_Original | qrySECTION8_Original.sql |
| qrySTAFFWITHNOSKILLS | qrySTAFFWITHNOSKILLS.sql |
| qrySTAFFWITHNOSKILLS_Original | qrySTAFFWITHNOSKILLS_Original.sql |
| qrytblContracts | qrytblContracts.sql |
| qrytblContractsAmendments | qrytblContractsAmendments.sql |
| qrytblContractsBillingBook | qrytblContractsBillingBook.sql |
| qrytblExpirations | qrytblExpirations.sql |
| qrytblLocations | qrytblLocations.sql |
| qrytblPeople | qrytblPeople.sql |
| qrytblPeopleClientsAFCServices | qrytblPeopleClientsAFCServices.sql |
| qrytblPeopleClientsAutismServices | qrytblPeopleClientsAutismServices.sql |
| qrytblPeopleClientsCLOServices | qrytblPeopleClientsCLOServices.sql |
| qrytblPeopleClientsDayServices | qrytblPeopleClientsDayServices.sql |
| qrytblPeopleClientsDemographics | qrytblPeopleClientsDemographics.sql |
| qrytblPeopleClientsIndividualSupportServices | qrytblPeopleClientsIndividualSupportServices.sql |
| qrytblPeopleClientsPCAServices | qrytblPeopleClientsPCAServices.sql |
| qrytblPeopleClientsPCAServicesContactNotes | qrytblPeopleClientsPCAServicesContactNotes.sql |
| qrytblPeopleClientsResidentialServices | qrytblPeopleClientsResidentialServices.sql |
| qrytblPeopleClientsSharedLivingServices | qrytblPeopleClientsSharedLivingServices.sql |
| qrytblPeopleClientsSpringboardServices | qrytblPeopleClientsSpringboardServices.sql |
| qrytblPeopleClientsVendors | qrytblPeopleClientsVendors.sql |
| qrytblPeopleClientsVocationalServices | qrytblPeopleClientsVocationalServices.sql |
| qrytblPeopleConsultants | qrytblPeopleConsultants.sql |
| qrytblPeopleDonors | qrytblPeopleDonors.sql |
| qrytblPeopleFamily | qrytblPeopleFamily.sql |
| qrytblPeopleStaffSupervisors | qrytblPeopleStaffSupervisors.sql |
| qrytblStaff | qrytblStaff.sql |
| qrytblStaffDedhamManagers | qrytblStaffDedhamManagers.sql |
| qrytblStaffEvalsAndSupervisions | qrytblStaffEvalsAndSupervisions.sql |
| qrytblStaffSkills | qrytblStaffSkills.sql |
| qryTILLLocations | qryTILLLocations.sql |


## Note about VBA Code

VBA code extraction requires MS Access to be fully installed and configured for COM automation.
If you need VBA code extraction, please ensure:
1. MS Access is installed (not just the Access Database Engine)
2. The database is not password protected
3. You have necessary permissions for COM automation

Alternatively, you can:
- Open the database in MS Access and manually export VBA modules
- Use the Access application's built-in export features
- Contact your database administrator for assistance
