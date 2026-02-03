# Extracted MS Access Database Content

This folder contains all extracted content from the TILL MS Access database.

## 📂 Folder Structure

- **`tables/`** - 47 table schema files with column definitions
- **`queries/`** - 166 SQL query files 
- **`vba/`** - 144 VBA code modules (forms, reports, and standard modules)
- **`reports/`** - Documentation and analysis reports

## 📖 Documentation

**Start here for navigation:**
- **[INDEX.md](INDEX.md)** - Quick reference guide organized by functionality
- **[COMPLETE_EXTRACTION_SUMMARY.md](reports/COMPLETE_EXTRACTION_SUMMARY.md)** - Comprehensive analysis and recommendations

## 📊 Extraction Summary

| Component | Count | Location |
|-----------|-------|----------|
| Table Schemas | 47 | `tables/` |
| SQL Queries | 166 | `queries/` |
| VBA Modules | 144 | `vba/` |
| **Total Files** | **357** | All folders |

## 🔍 Quick Links

### Find by Category

- **People Management** - `vba/Form_frmPeople.vba`, `queries/qryPeopleAllPeopleRecords.sql`
- **Client Services (19 types)** - `vba/Form_frmPeopleClientsService*.vba`
- **Locations** - `vba/Form_frmLocations.vba`, `queries/qryTILLLocations.sql`
- **Contracts & Billing** - `vba/Form_frmContracts*.vba`
- **Expirations (27 queries)** - `queries/qryEXPIRATIONS*.sql`
- **Donors & Fundraising** - `vba/Form_frmPeopleDonors*.vba`
- **Reports (16 modules)** - `vba/Report_rpt*.vba`
- **Main Menu** - `vba/Form_frmMainMenu.vba` (489 lines - entry point)
- **Utilities** - `vba/Utilities.vba` (232 lines of helper functions)

## 🗄️ Original Database

**File:** `../TILLDB_V9.14_20260128 - WEB.accdb`  
**Size:** 53,365 lines  
**Type:** MS Access 2007+ (.accdb)

## 🔗 Azure SQL Connection

This Access database connects to Azure SQL Server:
- **Server:** tillsqlserver.database.windows.net
- **User:** tillsqladmin
- **Type:** Hybrid split database (Access frontend + SQL backend)

## ⚙️ Extraction Tools Used

1. **Python (ADOX)** - Extracted table schemas and queries
2. **VBScript** - Extracted VBA code modules

Scripts available in parent directory:
- `../../extract_access_adox.py`
- `../../extract_vba.vbs`

## 📅 Extraction Date

**January 29, 2026**

---

For detailed information, see the documentation in the `reports/` folder.
