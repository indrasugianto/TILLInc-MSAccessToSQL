# Database Assessment Tools

This directory contains Python tools to connect to and assess the Azure SQL database **TILLDBWEB_Prod**.

## 🔐 Security Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure Database Connection

Copy `.env.example` to `.env` and set your Azure SQL credentials:
```
AZURE_SQL_SERVER=your_server.database.windows.net
AZURE_SQL_DATABASE=your_database
AZURE_SQL_USER=your_username
AZURE_SQL_PASSWORD=your_password
```

⚠️ **IMPORTANT:** The `.env` file is in `.gitignore` and will NOT be committed. See `.env.example` for the template.

## 🛠️ Available Tools

### 1. `db_connection.py` - Database Connection Utility

Provides secure connection management to Azure SQL.

**Test the connection:**
```bash
python db_connection.py
```

**Use in your scripts:**
```python
from db_connection import AzureSQLConnection

# Use as context manager (recommended)
with AzureSQLConnection() as db:
    columns, results = db.execute_query("SELECT * FROM YourTable")
    
# Or manual connection
db = AzureSQLConnection()
db.connect()
columns, results = db.execute_query("SELECT * FROM YourTable")
db.disconnect()
```

---

### 2. `assess_database.py` - Comprehensive Database Assessment

Analyzes the Azure SQL database and generates a detailed report.

**Run assessment:**
```bash
python assess_database.py
```

**What it analyzes:**
- ✅ Database overview and version
- ✅ Table statistics (row counts, sizes)
- ✅ Views
- ✅ Stored procedures
- ✅ Functions
- ✅ Indexes (including primary keys)
- ✅ Foreign key relationships
- ✅ Triggers
- ✅ Linked servers
- ✅ Comparison with Access extraction

**Output:**
- Console display of all findings
- Report saved to `assessment_reports/database_assessment_YYYYMMDD_HHMMSS.txt`

---

## 📊 Understanding the Architecture

### Hybrid Database Setup

The MS Access database (`TILLDB_V9.14_20260128 - WEB.accdb`) is a **frontend application** that connects to the Azure SQL database (`TILLDBWEB_Prod`).

```
┌─────────────────────────┐
│   MS Access Frontend    │
│  (User Interface + VBA) │
│                         │
│  - Forms                │
│  - Reports              │
│  - Business Logic       │
│  - Local Queries        │
└────────────┬────────────┘
             │
             │ Connection
             ↓
┌─────────────────────────┐
│   Azure SQL Database    │
│   TILLDBWEB_Prod        │
│                         │
│  - Tables (data)        │
│  - Views                │
│  - Stored Procedures    │
│  - Relationships        │
└─────────────────────────┘
```

### Data Flow

1. **Access Database** contains:
   - Forms and UI components
   - VBA business logic
   - Some local queries that reference SQL tables
   - Linked tables to Azure SQL

2. **Azure SQL Database** contains:
   - All actual data tables
   - Views
   - Stored procedures (if any)
   - Data relationships

## 🔍 Common Tasks

### Check Connection

```bash
python db_connection.py
```

### Run Full Assessment

```bash
python assess_database.py
```

### Query Specific Tables

```python
from db_connection import AzureSQLConnection

with AzureSQLConnection() as db:
    # Get all tables
    columns, results = db.execute_query("""
        SELECT TABLE_SCHEMA, TABLE_NAME 
        FROM INFORMATION_SCHEMA.TABLES 
        WHERE TABLE_TYPE = 'BASE TABLE'
        ORDER BY TABLE_NAME
    """)
    
    for row in results:
        print(f"{row[0]}.{row[1]}")
```

### Export Table Schema

```python
from db_connection import AzureSQLConnection

with AzureSQLConnection() as db:
    # Get columns for a specific table
    columns, results = db.execute_query("""
        SELECT 
            COLUMN_NAME,
            DATA_TYPE,
            CHARACTER_MAXIMUM_LENGTH,
            IS_NULLABLE
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = 'YourTableName'
        ORDER BY ORDINAL_POSITION
    """)
    
    for row in results:
        print(f"{row[0]}: {row[1]} - Nullable: {row[3]}")
```

## 🚨 Troubleshooting

### Error: "No suitable ODBC driver found"

**Solution:** Install Microsoft ODBC Driver for SQL Server

**Windows:**
```
Download from: https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server
Install: ODBC Driver 18 for SQL Server
```

### Error: "Login failed for user"

**Check:**
1. Credentials in `.env` file are correct
2. Firewall rules allow your IP address
3. User has necessary permissions

**Azure SQL Firewall:**
- Go to Azure Portal
- Navigate to your SQL Server
- Settings → Networking
- Add your client IP address

### Error: "Cannot open database"

**Check:**
1. Database name is correct: `TILLDBWEB_Prod`
2. Database is online in Azure Portal
3. User has access to the specific database

### Error: Module not found (pyodbc, etc.)

**Solution:**
```bash
pip install -r requirements.txt
```

## 📝 Adding New Assessment Scripts

You can create additional scripts using the connection utility:

```python
from db_connection import AzureSQLConnection

def my_custom_analysis():
    """Your custom database analysis"""
    with AzureSQLConnection() as db:
        # Your SQL queries here
        columns, results = db.execute_query("YOUR SQL HERE")
        
        # Process results
        for row in results:
            print(row)

if __name__ == "__main__":
    my_custom_analysis()
```

## 🔗 Related Files

- **`.env`** - Database credentials (not in git)
- **`.env.example`** - Template for credentials
- **`requirements.txt`** - Python dependencies
- **`db_connection.py`** - Connection utility
- **`assess_database.py`** - Assessment tool
- **`assessment_reports/`** - Generated reports

## 📚 Next Steps

1. ✅ Test connection: `python db_connection.py`
2. ✅ Run assessment: `python assess_database.py`
3. 📊 Review generated reports in `assessment_reports/`
4. 🔍 Compare SQL tables with Access queries
5. 📝 Document data model and relationships
6. 🚀 Plan migration strategy

---

**Last updated:** February 2, 2026  
**Database:** TILLDBWEB_Prod  
**Server:** tillsqlserver.database.windows.net
