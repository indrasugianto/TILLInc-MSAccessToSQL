# TILL MS Access to SQL Migration Project

This repository contains the extracted content and migration tools for the TILL MS Access database.

## 📋 Project Overview

**Purpose:** Extract and document all content from the TILL MS Access database (`TILLDB_V9.14_20260128 - WEB.accdb`) to facilitate migration to a modern stack.

**Database Type:** Hybrid MS Access frontend connected to Azure SQL Server backend

## 🗄️ Azure SQL Connection

- **Server:** tillsqlserver.database.windows.net
- **User:** tillsqladmin
- **Database Type:** Azure SQL Database

## 📂 Repository Structure

```
TILLInc-MSAccessToSQL/
├── msaccess/
│   ├── TILLDB_V9.14_20260128 - WEB.accdb  (Original database - not in repo)
│   └── extracted/                          (All extracted content)
│       ├── README.md                       (Quick start guide)
│       ├── INDEX.md                        (Quick reference)
│       ├── tables/                         (47 table schemas)
│       ├── queries/                        (166 SQL queries)
│       ├── vba/                            (144 VBA modules)
│       └── reports/                        (Documentation)
├── extract_access_adox.py                  (Python extraction script)
├── extract_vba.vbs                         (VBScript VBA extraction)
├── extract_access_content.ps1              (PowerShell alternative)
└── README.md                               (This file)
```

## 📊 Extracted Content Summary

| Component | Count | Location |
|-----------|-------|----------|
| **Table Schemas** | 47 | `msaccess/extracted/tables/` |
| **SQL Queries** | 166 | `msaccess/extracted/queries/` |
| **VBA Modules** | 144 | `msaccess/extracted/vba/` |
| **Total Files** | **357** | All folders |

### Key Components Extracted

- **Service Types:** Adult Coaching, Autism, CLO, Day, ISS, PCA, Residential, Shared Living, Springboard, TRASE, Vocational
- **Management Systems:** People, Locations, Contracts, Donors, Family, Staff
- **Reporting:** 16 report modules with comprehensive reporting functionality
- **Maintenance:** 27 expiration tracking queries, 27 data archival queries
- **Business Logic:** 144 VBA modules with forms, reports, and utilities

## 🚀 Getting Started

### View Extracted Content

1. Navigate to `msaccess/extracted/`
2. Start with `README.md` for overview
3. Use `INDEX.md` for quick reference by functionality
4. See `reports/COMPLETE_EXTRACTION_SUMMARY.md` for comprehensive analysis

### Run Extraction Scripts (if needed)

**Prerequisites:**
- Python 3.x with `pywin32` package
- MS Access 2007+ or Access Database Engine
- VBScript support (Windows)

**Extract Queries and Table Schemas:**
```bash
python extract_access_adox.py
```

**Extract VBA Code:**
```bash
cscript extract_vba.vbs
```

## 📖 Documentation

- **[Quick Start](msaccess/extracted/README.md)** - Get started with extracted content
- **[Quick Reference](msaccess/extracted/INDEX.md)** - Find files by functionality
- **[Complete Analysis](msaccess/extracted/reports/COMPLETE_EXTRACTION_SUMMARY.md)** - Full extraction report with recommendations

## 🔍 Key Findings

### Database Architecture
- **Type:** Hybrid split database (Access frontend + Azure SQL backend)
- **Entry Point:** `Form_frmMainMenu.vba` (489 lines)
- **Main Utility Module:** `Utilities.vba` (232 lines)
- **Address Validation:** SmartyStreets API integration

### Service Management
The database tracks 15+ different service types:
- Residential Services (53 columns)
- Day Services (51 columns)
- Demographics (121 columns)
- PCA Services (38 columns)
- And many more...

### Business Systems
- Expiration tracking (27 queries)
- Data archival (27 queries)
- Donor management
- Contract and billing
- Staff scheduling
- Comprehensive reporting

## 🛠️ Technology Stack

### Current (MS Access)
- **Frontend:** MS Access 2007+ (.accdb)
- **Backend:** Azure SQL Server
- **Language:** VBA (Visual Basic for Applications)
- **APIs:** SmartyStreets (address validation)

### Extraction Tools
- **Python:** ADOX (ActiveX Data Objects Extensions)
- **VBScript:** Access.Application COM automation
- **PowerShell:** Alternative extraction method

## 🎯 Migration Goals

1. **Extract all database content** ✅ Complete
   - Table schemas
   - SQL queries
   - VBA business logic

2. **Document architecture** ✅ Complete
   - System analysis
   - Business logic documentation
   - Data flow mapping

3. **Plan migration** 🔄 Next Steps
   - Convert Access queries to T-SQL stored procedures
   - Port VBA business logic to modern language
   - Design new frontend (Web/Desktop)
   - Implement modern authentication

## 🔐 Security Considerations

⚠️ **Important:** The extracted code contains hardcoded credentials:
- Database connection strings
- API keys (SmartyStreets)
- Email passwords

**Recommendations:**
- Use Azure Key Vault for secrets
- Implement proper authentication
- Review and rotate all credentials
- Remove hardcoded passwords from code

## 📅 Project Timeline

- **Extraction Date:** January 29, 2026
- **Database Version:** TILLDB_V9.14_20260128 - WEB
- **Total Files Extracted:** 357

## 🤝 Contributing

This is an internal migration project. For questions or access, contact the TILL database team.

## 📝 License

Internal TILL Inc. project - All rights reserved

---

**Last Updated:** January 29, 2026
