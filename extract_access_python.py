"""
Extract queries and VBA code from MS Access database
"""
import os
import sys
from pathlib import Path

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
vba_dir = Path(OUTPUT_DIR) / "vba"
reports_dir = Path(OUTPUT_DIR) / "reports"

queries_dir.mkdir(parents=True, exist_ok=True)
vba_dir.mkdir(parents=True, exist_ok=True)
reports_dir.mkdir(parents=True, exist_ok=True)

print(f"Starting extraction from: {DATABASE_PATH}")
print(f"Output directory: {OUTPUT_DIR}")

# Initialize counters and lists
query_count = 0
vba_count = 0
query_list = []
module_list = []

access = None
try:
    # Create Access Application object
    print("\nCreating Access application object...")
    access = win32com.client.Dispatch("Access.Application")
    access.Visible = False
    
    # Open the database
    print("Opening database...")
    access.OpenCurrentDatabase(DATABASE_PATH, False)
    
    db = access.CurrentDb()
    
    # Extract Queries
    print("\n" + "="*60)
    print("EXTRACTING QUERIES")
    print("="*60)
    
    query_defs = db.QueryDefs
    print(f"Total query definitions found: {query_defs.Count}")
    
    for i in range(query_defs.Count):
        qry = query_defs.Item(i)
        query_name = qry.Name
        
        # Skip system queries
        if query_name.startswith("~") or query_name.startswith("MSys"):
            continue
            
        query_count += 1
        
        try:
            sql = qry.SQL
            query_type = qry.Type
            
            # Sanitize filename
            safe_filename = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in query_name)
            file_path = queries_dir / f"{safe_filename}.sql"
            
            # Create query content
            query_content = f"""-- Query Name: {query_name}
-- Query Type: {query_type}
-- Extracted: {Path(__file__).name}

{sql}
"""
            
            # Write to file
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(query_content)
            
            query_list.append({
                'name': query_name,
                'type': query_type,
                'file': f"{safe_filename}.sql"
            })
            
            print(f"  [OK] Extracted: {query_name}")
            
        except Exception as e:
            print(f"  [ERROR] Error extracting query '{query_name}': {e}")
    
    print(f"\nTotal Queries Extracted: {query_count}")
    
    # Extract VBA Code
    print("\n" + "="*60)
    print("EXTRACTING VBA CODE")
    print("="*60)
    
    try:
        vba_project = access.VBE.VBProjects(1)
        print(f"VBA Project name: {vba_project.Name}")
        print(f"Total components found: {vba_project.VBComponents.Count}")
        
        for i in range(vba_project.VBComponents.Count):
            component = vba_project.VBComponents.Item(i + 1)  # VBA is 1-indexed
            module_name = component.Name
            module_type = component.Type
            
            # Get module type name
            type_names = {
                1: "Standard Module",
                2: "Class Module",
                3: "Form Module",
                100: "Document Module"
            }
            module_type_name = type_names.get(module_type, f"Unknown ({module_type})")
            
            code_module = component.CodeModule
            line_count = code_module.CountOfLines
            
            if line_count > 0:
                vba_count += 1
                
                # Get code
                code = code_module.Lines(1, line_count)
                
                # Sanitize filename
                safe_filename = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in module_name)
                file_path = vba_dir / f"{safe_filename}.vba"
                
                # Create module content
                module_content = f"""' Module Name: {module_name}
' Module Type: {module_type_name}
' Lines of Code: {line_count}
' Extracted: {Path(__file__).name}

{code}
"""
                
                # Write to file
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(module_content)
                
                module_list.append({
                    'name': module_name,
                    'type': module_type_name,
                    'lines': line_count,
                    'file': f"{safe_filename}.vba"
                })
                
                print(f"  [OK] Extracted: {module_name} ({module_type_name}) - {line_count} lines")
            else:
                print(f"  - Skipped: {module_name} (empty)")
                
    except Exception as e:
        print(f"\n[WARNING] Warning: Could not access VBA project")
        print(f"  Error: {e}")
        print("  The database may be password protected or VBA may not be accessible.")
    
    print(f"\nTotal VBA Modules Extracted: {vba_count}")
    
    # Create summary report
    print("\n" + "="*60)
    print("CREATING SUMMARY REPORT")
    print("="*60)
    
    summary = f"""# MS Access Database Extraction Report

**Database:** `{DATABASE_PATH}`  
**Extraction Date:** {Path(__file__).name}

## Connection Information
- **Server:** tillsqlserver.database.windows.net
- **User:** tillsqladmin
- **Database Type:** Azure SQL Database

## Summary
- **Total Queries Extracted:** {query_count}
- **Total VBA Modules Extracted:** {vba_count}

## Queries Extracted

"""
    
    if query_list:
        summary += "| Query Name | Type | Output File |\n"
        summary += "|------------|------|-------------|\n"
        for q in query_list:
            summary += f"| {q['name']} | {q['type']} | {q['file']} |\n"
    else:
        summary += "*No queries found.*\n"
    
    summary += "\n## VBA Modules Extracted\n\n"
    
    if module_list:
        summary += "| Module Name | Type | Lines | Output File |\n"
        summary += "|-------------|------|-------|-------------|\n"
        for m in module_list:
            summary += f"| {m['name']} | {m['type']} | {m['lines']} | {m['file']} |\n"
    else:
        summary += "*No VBA modules found or VBA not accessible.*\n"
    
    # Write summary
    report_path = reports_dir / "extraction_summary.md"
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(summary)
    
    print(f"\n[OK] Summary report saved to: {report_path}")
    
    # Close database
    access.CloseCurrentDatabase()
    access.Quit()
    
    print("\n" + "="*60)
    print("EXTRACTION COMPLETED SUCCESSFULLY!")
    print("="*60)
    print(f"Output directory: {OUTPUT_DIR}")
    print(f"  - Queries: {queries_dir}")
    print(f"  - VBA Code: {vba_dir}")
    print(f"  - Reports: {reports_dir}")
    
except Exception as e:
    print(f"\nERROR during extraction: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
    
finally:
    # Cleanup
    if access:
        try:
            access.Quit()
        except:
            pass
    
    # Release COM object
    try:
        del access
    except:
        pass

print("\n[OK] Done!")
