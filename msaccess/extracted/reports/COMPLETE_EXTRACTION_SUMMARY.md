# MS Access Database Complete Extraction Report

## Database Information
- **Database File:** `TILLDB_V9.14_20260128 - WEB.accdb`
- **Source Path:** `c:\GitHub\TILLInc-MSAccessToSQL\msaccess\TILLDB_V9.14_20260128 - WEB.accdb`
- **Extraction Date:** January 29, 2026
- **Total Database Size:** 53,365 lines (original .accdb file)

## Azure SQL Connection Details
- **Server:** tillsqlserver.database.windows.net
- **Database:** TILLDBWEB_Prod (configure credentials in `.env`; see project root `.env.example`)

---

## Extraction Summary

### ✅ Successfully Extracted

| Category | Count | Output Directory |
|----------|-------|------------------|
| **Table Schemas** | 47 | `extracted/tables/` |
| **SQL Queries** | 166 | `extracted/queries/` |
| **VBA Modules** | 144 | `extracted/vba/` |
| **Reports** | Multiple | `extracted/reports/` |

---

## Detailed Breakdown

### 1. Table Schemas (47 files)

Table schemas were extracted with complete column information including:
- Column names
- Data types
- Field sizes
- Table types (TABLE, VIEW, LINK)

**Key Tables Extracted:**
- `catLogonAs` - 1 column (TABLE)
- `RedReport` - 13 columns (TABLE)
- `tblDBChangeHistoryOLD` - 3 columns (TABLE)
- `tblLocationsDedhamManagers` - 67 columns (TABLE)
- `tblTILLMonthlyMasterArchive` - 147 columns (TABLE)

**Key Views Extracted:**
- `qryPeopleAllPeopleRecords` - 86 columns
- `qryPeopleClientsDemographics` - 121 columns
- `qrytblPeople` - 91 columns
- `qrytblLocations` - 69 columns
- `qrytblPeopleClientsVendors` - 58 columns
- And 42 more views...

**Output Files:** Each table has a `*_schema.txt` file with complete column definitions.

---

### 2. SQL Queries (166 files)

All queries were extracted with their complete SQL definitions. The queries include:

#### Query Categories:

**Expiration Queries (27 queries)**
- `qryEXPIRATIONS00` through `qryEXPIRATIONS26`
- `qryExpirationsStaffBySkills`
- `qryExpirationsStaffCull`
- `qryProgramExpirations`

**Delete/Archive Queries (27 queries)**
- `qryCopyDeletedAdultCoaching`
- `qryCopyDeletedAdultCompanion`
- `qryCopyDeletedAutism`
- `qryCopyDeletedClientDemographics`
- `qryCopyDeletedClientVendors`
- `qryCopyDeletedCLO`
- And 21 more deletion/archival queries...

**Update Queries (20 queries)**
- `qryUpdateAllPeopleTable`
- `qryUpdateAsstDirectors`
- `qryUpdateCoordinators`
- `qryUpdateDedhamHQLocations`
- `qryUpdateDedhamManagers`
- `qryUpdateSpringboardClientsStep1-6`
- `qryUpdateSpringboardLeaders01-06`
- And more...

**Service Queries**
- `qrytblPeopleClientsAFCServices`
- `qrytblPeopleClientsAutismServices`
- `qrytblPeopleClientsCLOServices`
- `qrytblPeopleClientsDayServices`
- `qrytblPeopleClientsIndividualSupportServices`
- `qrytblPeopleClientsPCAServices`
- `qrytblPeopleClientsResidentialServices`
- `qrytblPeopleClientsSharedLivingServices`
- `qrytblPeopleClientsSpringboardServices`
- `qrytblPeopleClientsVocationalServices`

**Report Queries**
- `qryDonationsForExport`
- `qryDonorAppealCreateMostRecentDonations`
- `qryDonorsMonthlyReport`
- `qryPeopleResidentialClientsReportExport`
- `qryCONINWORKS`
- `qryCONINWORKSExport`
- `qryRUNREPORT05`

**Demographic & Data Queries**
- `qryAutismIndividualsAndFamilyForDDS`
- `qryCensus`
- `qryClientGovernmentAccounts`
- `qrySECTION8`
- `qrySTAFFWITHNOSKILLS`

**Maintenance Queries**
- `qryCreateAllPeopleTable`
- `qryLoadConsultantsTable`
- `qryLoadFormDemographics`
- `qryLoadFormVendors`
- `qryLoadPhoneDirectory`
- `qryRefreshResidentialContacts`

**Output Files:** Each query is saved as a `.sql` file with the complete SQL statement.

---

### 3. VBA Modules (144 files)

Complete VBA code was successfully extracted including all forms, reports, and standard modules.

#### VBA Module Categories:

**Standard Modules (2)**
- `Utilities.vba` - 232 lines (utility functions)
- `AddressValidation.vba` - Address validation functions
- `Expirations_Updated.vba` (Expirations module) - Expiration management functions

**Form Modules (53)**

**Main Navigation:**
- `Form_frmMainMenu.vba` - 489 lines (Main menu form)
- `Form_frmRpt.vba` - Reports menu
- `Form_frmPleaseWaitCustom.vba` - Progress display
- `Form_frmProgressMessages.vba` - Message display

**People Management Forms:**
- `Form_frmPeople.vba` - Main people form
- `Form_frmPeopleAddressMaintenance.vba` - Address management
- `Form_frmPeopleChangeName.vba` - Name changes
- `Form_frmPeopleClientsContacts.vba` - Client contacts
- `Form_frmPeopleClientsDemographics.vba` - Demographics
- `Form_frmPeopleClientsSelectLocation.vba` - Location selection
- `Form_frmPeopleEnterAndValidatePerson.vba` - Person validation
- `Form_frmPeopleSelectPerson.vba` - Person selection
- `Form_frmPeopleScheduleStaffChanges.vba` - Staff scheduling

**Client Service Forms (19):**
- `Form_frmPeopleClientsServiceAdultCoaching.vba`
- `Form_frmPeopleClientsServiceAdultCompanion.vba`
- `Form_frmPeopleClientsServiceAutism.vba`
- `Form_frmPeopleClientsServiceCLO.vba`
- `Form_frmPeopleClientsServiceCommunityConnections.vba`
- `Form_frmPeopleClientsServiceDay.vba`
- `Form_frmPeopleClientsServiceIndividualSupport.vba`
- `Form_frmPeopleClientsServiceNHDay.vba`
- `Form_frmPeopleClientsServiceNHRes.vba`
- `Form_frmPeopleClientsServicePCA.vba`
- `Form_frmPeopleClientsServicePCAContactNotes.vba`
- `Form_frmPeopleClientsServiceResidential.vba`
- `Form_frmPeopleClientsServiceSharedLiving.vba`
- `Form_frmPeopleClientsServiceSpringboard.vba`
- `Form_frmPeopleClientsServiceTransportation.vba`
- `Form_frmPeopleClientsServiceTRASE.vba`
- `Form_frmPeopleClientsServiceVocational.vba`
- `Form_frmPeopleClientsVendors.vba`

**Family and Donor Forms:**
- `Form_frmPeopleFamily.vba` - Family management
- `Form_frmPeopleFamilyEnterAndValidatePerson.vba` - Family validation
- `Form_frmPeopleConsultants.vba` - Consultants
- `Form_frmPeopleDonors.vba` - Donor management
- `Form_frmPeopleDonorsNewDonation.vba` - New donations

**Location Management Forms:**
- `Form_frmLocations.vba` - Main locations form
- `Form_frmLocationsAddressMaintenance.vba` - Address maintenance
- `Form_frmLocationsContacts.vba` - Location contacts
- `Form_frmLocationsSelectStaff.vba` - Staff selection

**Contract Management Forms:**
- `Form_frmContracts.vba` - Contracts
- `Form_frmContractsAmendments.vba` - Contract amendments
- `Form_frmContractsBillingBook.vba` - Billing

**Maintenance Forms:**
- `Form_frmDBChanges.vba` - Database changes
- `Form_frmMaintEditReferenceTables.vba` - Reference tables
- `Form_frmMaintMonthlyMaster.vba` - Monthly master
- `Form_frmMaintMonthlyMasterComments.vba` - Comments
- `Form_frmMaintUserPermissions.vba` - User permissions

**Report Forms:**
- `Form_frmRptExpDNRTILLEGRAM.vba` - DNR/Tillegram exports
- `Form_frmRptExpMAILINGSANDSPREADSHEETS.vba` - Mailings
- `Form_frmRptFinancialAndLetters.vba` - Financial reports
- `Form_frmRptPCAReportsAndExports.vba` - PCA reports

**Report Modules (16)**
- `Report_rptRESBANKING.vba` - Banking report
- `Report_rptRESBENEFITS.vba` - Benefits report
- `Report_rptRESCLIENTS.vba` - Clients report
- `Report_rptRESCLIENTSBYSITE.vba` - Clients by site
- `Report_rptRESCLIENTSNoContacts.vba` - Clients without contacts
- `Report_rptRESCLUSTERS.vba` - Clusters report
- `Report_rptRESEMERGENCYPHONES.vba` - Emergency phones
- `Report_rptRESEMERGENCYPHONESSUMMARY.vba` - Emergency phones summary
- `Report_rptRESRMBD.vba` - RMBD report
- `Report_rptSECTION8.vba` - Section 8 report
- `Report_rptSPRINGBOARD.vba` - Springboard report
- `Report_rptSPRINGBOARDClients.vba` - Springboard clients
- `Report_rptTILLLocations.vba` - TILL locations
- `Report_rptVALIDCLI.vba` - Valid clients
- `Report_rptVALIDFAM.vba` - Valid families
- `Report_rptVICTORIABILLS.vba` - Victoria bills

**Output Files:** Each module is saved as a `.vba` file with complete source code and metadata.

---

## Extraction Methods Used

### 1. ADOX (ActiveX Data Objects Extensions)
- Used for extracting table schemas and queries
- Works with Microsoft Access Database Engine
- Does not require full MS Access installation

### 2. ADO (ActiveX Data Objects)
- Used for alternative query extraction
- Extracted additional views and stored procedures
- Connection string: `Provider=Microsoft.ACE.OLEDB.12.0`

### 3. VBScript with Access.Application COM
- Used for VBA code extraction
- Requires MS Access to be installed
- Successfully extracted all 144 VBA modules

---

## File Organization

```
c:\GitHub\TILLInc-MSAccessToSQL\
├── msaccess/
│   ├── TILLDB_V9.14_20260128 - WEB.accdb (Original database)
│   └── extracted/
│       ├── tables/                  (47 table schema files)
│       │   ├── catLogonAs_schema.txt
│       │   ├── RedReport_schema.txt
│       │   ├── tblDBChangeHistoryOLD_schema.txt
│       │   └── ...
│       ├── queries/                 (166 SQL query files)
│       │   ├── qryAppendDeletedPerson.sql
│       │   ├── qryCONINWORKS.sql
│       │   ├── qryEXPIRATIONS00.sql
│       │   └── ...
│       ├── vba/                     (144 VBA module files)
│       │   ├── Utilities.vba
│       │   ├── Form_frmMainMenu.vba
│       │   ├── Form_frmPeople.vba
│       │   └── ...
│       └── reports/
│           ├── extraction_summary.md
│           ├── COMPLETE_EXTRACTION_SUMMARY.md
│           ├── INDEX.md
│           └── vba_extraction_report.txt (if generated)
├── extract_access_content.ps1   (PowerShell extraction script)
├── extract_access_python.py     (Python extraction script)
├── extract_access_adox.py       (ADOX extraction script - used)
└── extract_vba.vbs              (VBScript VBA extraction - used)
```

---

## Key Findings & Insights

### Database Structure
- **Hybrid Architecture:** This is a split Access database that uses linked tables to Azure SQL Server
- **Complex Service Tracking:** Multiple service types tracked (AFC, Autism, CLO, Day, ISS, PCA, Residential, Shared Living, Springboard, Vocational, etc.)
- **Comprehensive People Management:** Tracks clients, family members, staff, donors, and consultants

### VBA Code Characteristics
- **Main Form:** `frmMainMenu` with 489 lines - initializes database connection and global parameters
- **SQL Connection:** Uses parameterized connection to Azure SQL (`TILLDataBase = CurrentDb`)
- **Email Notifications:** Configured with password stored in `appParameters` table
- **Fiscal Year Logic:** Automatic FY calculation based on current date
- **Progress Indicators:** Custom progress message system for long operations
- **Address Validation:** Dedicated module for address validation
- **Expiration Management:** Complex expiration tracking system

### Query Patterns
- **Staged Updates:** Springboard updates use 6-step process
- **Deletion Archival:** All deletions copy to archive tables first
- **Expiration Tracking:** 27 different expiration queries for various certifications/documents
- **Data Refresh:** Automated table refresh queries for consultants, staff, contacts
- **Export Functionality:** Multiple export queries for reporting and external systems

---

## Next Steps & Recommendations

### 1. Database Migration to Pure Azure SQL
Since this is already connected to Azure SQL, consider:
- Migrate remaining local tables to Azure SQL
- Remove Access frontend dependency
- Create web-based or Windows application frontend
- Maintain business logic from VBA in new application layer

### 2. VBA Code Analysis
- Review utility functions in `Utilities.vba` for reusable business logic
- Document email notification system configuration
- Map form dependencies and navigation flow
- Identify stored passwords and credentials (security review)

### 3. Query Optimization
- Review complex multi-step update queries
- Consider converting to stored procedures in Azure SQL
- Optimize expiration queries (27 separate queries could be consolidated)
- Review deletion/archival process for efficiency

### 4. Data Model Documentation
- Create ER diagram from table schemas
- Document relationships between tables
- Identify primary keys and foreign keys
- Document business rules embedded in VBA

### 5. Security Review
- Credentials are stored in plain text in code
- Review `appParameters` table for sensitive data
- Implement proper credential management (Azure Key Vault)
- Review user permissions system

---

## Technical Notes

### Character Encoding
All files extracted with UTF-8 encoding to preserve special characters.

### Line Endings
Windows-style line endings (CRLF) used throughout.

### SQL Dialect
Queries use MS Access SQL dialect which may need translation for pure T-SQL in Azure SQL Server.

### VBA Dependencies
- Requires MS Access 2007 or later
- Uses Office COM objects
- Connection string uses ACE.OLEDB.12.0 provider

---

## Extraction Scripts

Three extraction scripts were created:

1. **extract_access_adox.py** (✅ Successfully used)
   - Python script using ADOX
   - Extracted tables and queries
   - No MS Access installation required

2. **extract_vba.vbs** (✅ Successfully used)
   - VBScript using Access.Application COM
   - Extracted all VBA modules
   - Requires MS Access installation

3. **extract_access_content.ps1** (❌ Not used - timeout issues)
   - PowerShell alternative script

---

## Support & Contact

For questions about the extracted code or migration assistance, refer to:
- **Database location:** `c:\GitHub\TILLInc-MSAccessToSQL\msaccess\`
- **Extracted files location:** `c:\GitHub\TILLInc-MSAccessToSQL\msaccess\extracted\`
- **Extraction date:** January 29, 2026
- **Azure SQL Server:** tillsqlserver.database.windows.net

---

**Report Generated:** January 29, 2026
**Total Files Extracted:** 357 files (47 tables + 166 queries + 144 VBA modules)
**Extraction Status:** ✅ Complete and Successful
