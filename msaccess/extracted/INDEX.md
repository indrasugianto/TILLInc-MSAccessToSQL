# Quick Reference Index - Extracted MS Access Content

## 📊 Summary
- **Tables:** 47 schema files
- **Queries:** 166 SQL files
- **VBA Modules:** 144 code files
- **Total:** 357 extracted files

---

## 📂 Directory Structure

### `/tables/` - Table Schemas (47 files)
All table and view schemas with column definitions.

### `/queries/` - SQL Queries (166 files)
All Access queries converted to SQL format.

### `/vba/` - VBA Code (144 files)
All forms, reports, and modules with complete source code.

### `/reports/` - Documentation
- `COMPLETE_EXTRACTION_SUMMARY.md` - Full extraction report
- `extraction_summary.md` - Initial extraction summary

---

## 🔍 Quick Find Guide

### People Management
**Forms:**
- `vba/Form_frmPeople.vba` - Main people form
- `vba/Form_frmPeopleSelectPerson.vba` - Person selector
- `vba/Form_frmPeopleEnterAndValidatePerson.vba` - Person validation

**Queries:**
- `queries/qryPeopleAllPeopleRecords.sql` - All people
- `queries/qrytblPeople.sql` - People table view

### Client Services
**Forms (19 services):**
- `vba/Form_frmPeopleClientsServiceResidential.vba`
- `vba/Form_frmPeopleClientsServiceDay.vba`
- `vba/Form_frmPeopleClientsServiceAutism.vba`
- `vba/Form_frmPeopleClientsServicePCA.vba`
- `vba/Form_frmPeopleClientsServiceVocational.vba`
- And 14 more service forms...

**Queries:**
- `queries/qrytblPeopleClientsResidentialServices.sql`
- `queries/qrytblPeopleClientsDayServices.sql`
- `queries/qrytblPeopleClientsAutismServices.sql`
- And more...

### Locations
**Forms:**
- `vba/Form_frmLocations.vba` - Main locations
- `vba/Form_frmLocationsContacts.vba` - Location contacts
- `vba/Form_frmLocationsSelectStaff.vba` - Staff assignment

**Queries:**
- `queries/qryTILLLocations.sql`
- `queries/qrytblLocations.sql`

**Tables:**
- `tables/tblLocationsDedhamManagers_schema.txt`

### Contracts & Billing
**Forms:**
- `vba/Form_frmContracts.vba`
- `vba/Form_frmContractsAmendments.vba`
- `vba/Form_frmContractsBillingBook.vba`

**Queries:**
- `queries/qrytblContracts.sql`
- `queries/qrytblContractsAmendments.sql`
- `queries/qrytblContractsBillingBook.sql`
- `queries/qryCurrentFYContracts.sql`

### Expirations System
**VBA Module:**
- `vba/Expirations_Updated.vba` (Expirations module) - Expiration management code

**Queries (27):**
- `queries/qryEXPIRATIONS00.sql` through `queries/qryEXPIRATIONS26.sql`
- `queries/qryExpirationsStaffBySkills.sql`
- `queries/qryExpirationsStaffCull.sql`
- `queries/qryProgramExpirations.sql`
- `queries/qrytblExpirations.sql`

### Donors & Fundraising
**Forms:**
- `vba/Form_frmPeopleDonors.vba`
- `vba/Form_frmPeopleDonorsNewDonation.vba`

**Queries:**
- `queries/qryDonations.sql`
- `queries/qryDonationsForExport.sql`
- `queries/qryDonorAppealCreateMostRecentDonations.sql`
- `queries/qryDonorsMonthlyReport.sql`
- `queries/qrytblPeopleDonors.sql`

### Family Management
**Forms:**
- `vba/Form_frmPeopleFamily.vba`
- `vba/Form_frmPeopleFamilyEnterAndValidatePerson.vba`

**Queries:**
- `queries/qrytblPeopleFamily.sql`
- `queries/qryClientFamilyMailing.sql`

### Reports
**Form Menus:**
- `vba/Form_frmRpt.vba` - Main report menu
- `vba/Form_frmRptExpDNRTILLEGRAM.vba`
- `vba/Form_frmRptExpMAILINGSANDSPREADSHEETS.vba`
- `vba/Form_frmRptFinancialAndLetters.vba`
- `vba/Form_frmRptPCAReportsAndExports.vba`

**Report Modules (16):**
- `vba/Report_rptRESCLIENTS.vba` - Residential clients
- `vba/Report_rptRESCLIENTSBYSITE.vba` - Clients by site
- `vba/Report_rptSECTION8.vba` - Section 8 report
- `vba/Report_rptSPRINGBOARD.vba` - Springboard report
- `vba/Report_rptTILLLocations.vba` - Locations report
- And 11 more...

### Maintenance & Admin
**Forms:**
- `vba/Form_frmMainMenu.vba` - Main application menu (489 lines)
- `vba/Form_frmDBChanges.vba` - Database changes
- `vba/Form_frmMaintEditReferenceTables.vba`
- `vba/Form_frmMaintMonthlyMaster.vba`
- `vba/Form_frmMaintUserPermissions.vba`

**Queries:**
- `queries/qryDeleteCorruptedPeopleRecords.sql`
- `queries/qryDeleteNullPeople.sql`
- `queries/qryRepairAnomalies1.sql` through `qryRepairAnomalies4.sql`

### Utility Code
**VBA Modules:**
- `vba/Utilities.vba` - 232 lines of utility functions
- `vba/AddressValidation.vba` - Address validation functions

### Data Loading & Refresh
**Queries:**
- `queries/qryLoadConsultantsTable.sql`
- `queries/qryLoadFormDemographics.sql`
- `queries/qryLoadFormVendors.sql`
- `queries/qryLoadPhoneDirectory.sql`
- `queries/qryRefreshResidentialContacts.sql`
- `queries/qryCreateAllPeopleTable.sql`
- `queries/qryUpdateAllPeopleTable.sql`

### Data Deletion & Archival (27 queries)
All deletion queries copy data to archive before deleting:
- `queries/qryCopyDeletedAdultCoaching.sql`
- `queries/qryCopyDeletedAutism.sql`
- `queries/qryCopyDeletedClientDemographics.sql`
- `queries/qryCopyDeletedResidential.sql`
- `queries/qryCopyDeletedPCA.sql`
- And 22 more...

### Update Queries
**Springboard (12 queries):**
- `queries/qryUpdateSpringboardClientsStep1.sql` through `Step6.sql`
- `queries/qryUpdateSpringboardLeaders01.sql` through `06.sql`

**Staff & Supervisors:**
- `queries/qryUpdateStaffTable.sql`
- `queries/qryUpdateStaffSupervisors.sql`
- `queries/qryUpdateStaffSupervisorsLocations.sql`
- `queries/qryUpdateStaffSupervisorsNames.sql`

**Organizational:**
- `queries/qryUpdateAsstDirectors.sql`
- `queries/qryUpdateCoordinators.sql`
- `queries/qryUpdateDedhamHQLocations.sql`
- `queries/qryUpdateDedhamManagers.sql`
- `queries/qryUpdateDedhamStaffCodes.sql`

---

## 🔧 Tools Used for Extraction

1. **Python (ADOX)** - `extract_access_adox.py`
   - Extracted table schemas and queries
   
2. **VBScript** - `extract_vba.vbs`
   - Extracted all VBA code

---

## 📖 Additional Documentation

See `reports/COMPLETE_EXTRACTION_SUMMARY.md` for:
- Detailed extraction methodology
- Database architecture insights
- Migration recommendations
- Security considerations
- Next steps

---

**Last Updated:** January 29, 2026
