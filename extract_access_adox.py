"""
Extract queries from MS Access database using ADOX (no MS Access installation needed)
For VBA extraction, we'll need to use a different method
"""
import os
import sys
from pathlib import Path
from datetime import datetime

try:
    import win32com.client
    print("win32com.client imported successfully")
except ImportError:
    print("ERROR: pywin32 not installed. Installing...")
    os.system("pip install pywin32")
    import win32com.client

# Configuration
DATABASE_PATH = r"c:\GitHub\TILLInc-MSAccessToSQL\msaccess\TILLDB_V9.14_20260203d - WEB.accdb"
OUTPUT_DIR = r"c:\GitHub\TILLInc-MSAccessToSQL\msaccess\extracted"

# Create output directories
queries_dir = Path(OUTPUT_DIR) / "queries"
tables_dir = Path(OUTPUT_DIR) / "tables"
reports_dir = Path(OUTPUT_DIR) / "reports"

queries_dir.mkdir(parents=True, exist_ok=True)
tables_dir.mkdir(parents=True, exist_ok=True)
reports_dir.mkdir(parents=True, exist_ok=True)

print(f"Starting extraction from: {DATABASE_PATH}")
print(f"Output directory: {OUTPUT_DIR}\n")

# Initialize counters and lists
query_count = 0
table_count = 0
query_list = []
table_list = []

# Connection string for Access 2007+ (.accdb)
conn_string = f"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={DATABASE_PATH};Persist Security Info=False;"

print("Attempting to connect using ADOX (Access Database Engine)...")
print(f"Connection string: {conn_string}\n")

catalog = None
connection = None

try:
    # Create ADOX Catalog object
    catalog = win32com.client.Dispatch("ADOX.Catalog")
    
    # Open the database
    print("Opening database catalog...")
    catalog.ActiveConnection = conn_string
    
    print("Successfully connected!\n")
    
    # Extract Tables Information
    print("=" * 60)
    print("EXTRACTING TABLE INFORMATION")
    print("=" * 60)
    
    tables = catalog.Tables
    print(f"Total tables found: {tables.Count}\n")
    
    for i in range(tables.Count):
        table = tables.Item(i)
        table_name = table.Name
        table_type = table.Type
        
        # Skip system tables
        if table_name.startswith("MSys") or table_name.startswith("~"):
            continue
        
        # Focus on USER tables and LINK tables
        if table_type not in ["TABLE", "LINK", "VIEW"]:
            continue
        
        table_count += 1
        
        # Get column information
        columns = table.Columns
        column_info = []
        
        for j in range(columns.Count):
            col = columns.Item(j)
            col_name = col.Name
            col_type = col.Type
            try:
                col_size = col.DefinedSize
            except:
                col_size = "N/A"
            
            column_info.append({
                'name': col_name,
                'type': col_type,
                'size': col_size
            })
        
        # Sanitize filename
        safe_filename = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in table_name)
        file_path = tables_dir / f"{safe_filename}_schema.txt"
        
        # Create table schema content
        table_content = f"""Table Name: {table_name}
Table Type: {table_type}
Column Count: {len(column_info)}
Extracted: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

Columns:
{'='*60}
"""
        
        for col in column_info:
            table_content += f"  - {col['name']}: Type={col['type']}, Size={col['size']}\n"
        
        # Write to file
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(table_content)
        
        table_list.append({
            'name': table_name,
            'type': table_type,
            'columns': len(column_info),
            'file': f"{safe_filename}_schema.txt"
        })
        
        print(f"  [+] {table_name} ({table_type}) - {len(column_info)} columns")
    
    print(f"\nTotal Tables Extracted: {table_count}")
    
    # Extract Queries using Procedures (Views and Queries)
    print("\n" + "=" * 60)
    print("EXTRACTING QUERIES")
    print("=" * 60)
    
    procedures = catalog.Procedures
    print(f"Total procedures found: {procedures.Count}\n")
    
    for i in range(procedures.Count):
        proc = procedures.Item(i)
        query_name = proc.Name
        
        # Skip system queries
        if query_name.startswith("~") or query_name.startswith("MSys"):
            continue
        
        query_count += 1
        
        # Get command
        command = proc.Command
        command_text = command.CommandText if hasattr(command, 'CommandText') else str(command)
        
        # Sanitize filename
        safe_filename = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in query_name)
        file_path = queries_dir / f"{safe_filename}.sql"
        
        # Create query content
        query_content = f"""-- Query Name: {query_name}
-- Extracted: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

{command_text}
"""
        
        # Write to file
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(query_content)
        
        query_list.append({
            'name': query_name,
            'file': f"{safe_filename}.sql"
        })
        
        print(f"  [+] {query_name}")
    
    print(f"\nTotal Queries Extracted: {query_count}")
    
except Exception as e:
    print(f"\nERROR during extraction: {e}")
    import traceback
    traceback.print_exc()

finally:
    # Cleanup
    if catalog:
        try:
            catalog.ActiveConnection = None
        except:
            pass
    
    del catalog

# Try alternative method for queries using ADO connection
print("\n" + "=" * 60)
print("TRYING ALTERNATIVE METHOD FOR QUERIES (ADO)")
print("=" * 60)

try:
    # Create ADO Connection
    connection = win32com.client.Dispatch("ADODB.Connection")
    connection.Open(conn_string)
    
    print("ADO Connection opened successfully\n")
    
    # Get schema for queries
    try:
        # OpenSchema for Views
        views_recordset = connection.OpenSchema(23)  # adSchemaViews = 23
        
        alternative_query_count = 0
        while not views_recordset.EOF:
            try:
                view_name = views_recordset.Fields("TABLE_NAME").Value
                view_def = views_recordset.Fields("VIEW_DEFINITION").Value
                
                # Skip if already extracted
                if not any(q['name'] == view_name for q in query_list):
                    alternative_query_count += 1
                    
                    # Sanitize filename
                    safe_filename = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in view_name)
                    file_path = queries_dir / f"{safe_filename}.sql"
                    
                    # Create query content
                    query_content = f"""-- Query Name: {view_name}
-- Extracted: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} (ADO Method)

{view_def}
"""
                    
                    # Write to file
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(query_content)
                    
                    query_list.append({
                        'name': view_name,
                        'file': f"{safe_filename}.sql"
                    })
                    
                    print(f"  [+] {view_name}")
            except:
                pass
            
            views_recordset.MoveNext()
        
        if alternative_query_count > 0:
            print(f"\nAdditional Queries Found: {alternative_query_count}")
            query_count += alternative_query_count
        else:
            print("No additional queries found via ADO method")
            
    except Exception as e:
        print(f"Could not retrieve views via ADO: {e}")
    
    connection.Close()
    
except Exception as e:
    print(f"Could not use ADO method: {e}")

# Create summary report
print("\n" + "=" * 60)
print("CREATING SUMMARY REPORT")
print("=" * 60)

summary = f"""# MS Access Database Extraction Report

**Database:** `{DATABASE_PATH}`  
**Extraction Date:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## Connection Information
- **Server:** tillsqlserver.database.windows.net
- **User:** tillsqladmin
- **Database Type:** Azure SQL Database

## Summary
- **Total Tables Extracted:** {table_count}
- **Total Queries Extracted:** {query_count}
- **VBA Modules:** Not extracted (requires MS Access application)

## Tables Extracted

"""

if table_list:
    summary += "| Table Name | Type | Columns | Output File |\n"
    summary += "|------------|------|---------|-------------|\n"
    for t in table_list:
        summary += f"| {t['name']} | {t['type']} | {t['columns']} | {t['file']} |\n"
else:
    summary += "*No tables found.*\n"

summary += "\n## Queries Extracted\n\n"

if query_list:
    summary += "| Query Name | Output File |\n"
    summary += "|------------|-------------|\n"
    for q in query_list:
        summary += f"| {q['name']} | {q['file']} |\n"
else:
    summary += "*No queries found.*\n"

summary += """

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
"""

# Write summary
report_path = reports_dir / "extraction_summary.md"
with open(report_path, 'w', encoding='utf-8') as f:
    f.write(summary)

print(f"\n[+] Summary report saved to: {report_path}")

print("\n" + "=" * 60)
print("EXTRACTION COMPLETED!")
print("=" * 60)
print(f"Output directory: {OUTPUT_DIR}")
print(f"  - Tables: {tables_dir} ({table_count} files)")
print(f"  - Queries: {queries_dir} ({query_count} files)")
print(f"  - Reports: {reports_dir}")
print("\nDone!")
