"""
Azure SQL Database Assessment Tool
Analyzes the TILLDBWEB_Prod database structure, objects, and statistics
"""

import os
import sys
from datetime import datetime
from pathlib import Path
from tabulate import tabulate
from db_connection import AzureSQLConnection

class DatabaseAssessment:
    """Comprehensive database assessment"""
    
    def __init__(self, output_dir="assessment_reports"):
        """Initialize assessment tool"""
        self.db = AzureSQLConnection()
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        self.report_lines = []
    
    def add_section(self, title, level=1):
        """Add a section header to the report"""
        if level == 1:
            self.report_lines.append("\n" + "=" * 80)
            self.report_lines.append(title)
            self.report_lines.append("=" * 80)
        elif level == 2:
            self.report_lines.append("\n" + "-" * 60)
            self.report_lines.append(title)
            self.report_lines.append("-" * 60)
        else:
            self.report_lines.append(f"\n### {title}")
    
    def add_text(self, text):
        """Add text to the report"""
        self.report_lines.append(text)
    
    def add_table(self, columns, data, title=None):
        """Add a formatted table to the report"""
        if title:
            self.add_text(f"\n{title}:")
        
        if data:
            table = tabulate(data, headers=columns, tablefmt="grid")
            self.report_lines.append(table)
        else:
            self.report_lines.append("  No data found.")
    
    def get_database_overview(self):
        """Get general database information"""
        self.add_section("DATABASE OVERVIEW", level=1)
        
        query = """
        SELECT 
            DB_NAME() AS DatabaseName,
            @@VERSION AS SQLServerVersion,
            (SELECT state_desc FROM sys.databases WHERE name = DB_NAME()) AS State,
            (SELECT recovery_model_desc FROM sys.databases WHERE name = DB_NAME()) AS RecoveryModel,
            (SELECT compatibility_level FROM sys.databases WHERE name = DB_NAME()) AS CompatibilityLevel,
            GETDATE() AS AssessmentDate
        """
        
        columns, results = self.db.execute_query(query)
        
        if results:
            row = results[0]
            self.add_text(f"\nDatabase Name: {row[0]}")
            self.add_text(f"SQL Server Version: {row[1][:100]}...")
            self.add_text(f"State: {row[2]}")
            self.add_text(f"Recovery Model: {row[3]}")
            self.add_text(f"Compatibility Level: {row[4]}")
            self.add_text(f"Assessment Date: {row[5]}")
    
    def get_table_statistics(self):
        """Get table count and statistics"""
        self.add_section("TABLE STATISTICS", level=1)
        
        query = """
        SELECT 
            t.name AS TableName,
            s.name AS SchemaName,
            p.rows AS [RowCount],
            SUM(a.total_pages) * 8 AS TotalSpaceKB,
            SUM(a.used_pages) * 8 AS UsedSpaceKB,
            (SUM(a.total_pages) - SUM(a.used_pages)) * 8 AS UnusedSpaceKB,
            COUNT(c.column_id) AS ColumnCount
        FROM sys.tables t
        INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
        INNER JOIN sys.indexes i ON t.object_id = i.object_id
        INNER JOIN sys.partitions p ON i.object_id = p.object_id AND i.index_id = p.index_id
        INNER JOIN sys.allocation_units a ON p.partition_id = a.container_id
        LEFT JOIN sys.columns c ON t.object_id = c.object_id
        WHERE t.is_ms_shipped = 0
        GROUP BY t.name, s.name, p.rows
        ORDER BY p.rows DESC
        """
        
        columns, results = self.db.execute_query(query)
        
        if results:
            self.add_text(f"\nTotal Tables: {len(results)}")
            self.add_table(columns, results[:20], "\nTop 20 Tables by Row Count")
            
            # Summary statistics
            total_rows = sum(row[2] for row in results)
            total_space_mb = sum(row[3] for row in results) / 1024
            self.add_text(f"\nTotal Rows Across All Tables: {total_rows:,}")
            self.add_text(f"Total Space Used: {total_space_mb:.2f} MB")
    
    def get_view_statistics(self):
        """Get view information"""
        self.add_section("VIEWS", level=1)
        
        query = """
        SELECT 
            s.name AS SchemaName,
            v.name AS ViewName,
            v.create_date AS CreatedDate,
            v.modify_date AS ModifiedDate
        FROM sys.views v
        INNER JOIN sys.schemas s ON v.schema_id = s.schema_id
        WHERE v.is_ms_shipped = 0
        ORDER BY s.name, v.name
        """
        
        columns, results = self.db.execute_query(query)
        
        if results:
            self.add_text(f"\nTotal Views: {len(results)}")
            self.add_table(columns, results, "\nAll Views")
        else:
            self.add_text("\nNo user-defined views found.")
    
    def get_stored_procedures(self):
        """Get stored procedure information"""
        self.add_section("STORED PROCEDURES", level=1)
        
        query = """
        SELECT 
            s.name AS SchemaName,
            p.name AS ProcedureName,
            p.create_date AS CreatedDate,
            p.modify_date AS ModifiedDate,
            COUNT(pr.parameter_id) AS ParameterCount
        FROM sys.procedures p
        INNER JOIN sys.schemas s ON p.schema_id = s.schema_id
        LEFT JOIN sys.parameters pr ON p.object_id = pr.object_id
        WHERE p.is_ms_shipped = 0
        GROUP BY s.name, p.name, p.create_date, p.modify_date
        ORDER BY s.name, p.name
        """
        
        columns, results = self.db.execute_query(query)
        
        if results:
            self.add_text(f"\nTotal Stored Procedures: {len(results)}")
            self.add_table(columns, results[:30], "\nStored Procedures (First 30)")
        else:
            self.add_text("\nNo user-defined stored procedures found.")
    
    def get_functions(self):
        """Get function information"""
        self.add_section("FUNCTIONS", level=1)
        
        query = """
        SELECT 
            s.name AS SchemaName,
            o.name AS FunctionName,
            o.type_desc AS FunctionType,
            o.create_date AS CreatedDate,
            o.modify_date AS ModifiedDate
        FROM sys.objects o
        INNER JOIN sys.schemas s ON o.schema_id = s.schema_id
        WHERE o.type IN ('FN', 'IF', 'TF', 'FS', 'FT')
        AND o.is_ms_shipped = 0
        ORDER BY s.name, o.name
        """
        
        columns, results = self.db.execute_query(query)
        
        if results:
            self.add_text(f"\nTotal Functions: {len(results)}")
            self.add_table(columns, results, "\nAll Functions")
        else:
            self.add_text("\nNo user-defined functions found.")
    
    def get_indexes(self):
        """Get index information"""
        self.add_section("INDEXES", level=1)
        
        query = """
        SELECT 
            s.name AS SchemaName,
            t.name AS TableName,
            i.name AS IndexName,
            i.type_desc AS IndexType,
            i.is_primary_key AS IsPrimaryKey,
            i.is_unique AS IsUnique,
            COUNT(ic.column_id) AS ColumnCount
        FROM sys.indexes i
        INNER JOIN sys.tables t ON i.object_id = t.object_id
        INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
        LEFT JOIN sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id
        WHERE t.is_ms_shipped = 0
        AND i.name IS NOT NULL
        GROUP BY s.name, t.name, i.name, i.type_desc, i.is_primary_key, i.is_unique
        ORDER BY s.name, t.name, i.name
        """
        
        columns, results = self.db.execute_query(query)
        
        if results:
            self.add_text(f"\nTotal Indexes: {len(results)}")
            
            # Count by type
            pk_count = sum(1 for row in results if row[4])
            unique_count = sum(1 for row in results if row[5] and not row[4])
            other_count = len(results) - pk_count - unique_count
            
            self.add_text(f"  Primary Keys: {pk_count}")
            self.add_text(f"  Unique Indexes: {unique_count}")
            self.add_text(f"  Other Indexes: {other_count}")
            
            self.add_table(columns, results[:30], "\nIndexes (First 30)")
    
    def get_foreign_keys(self):
        """Get foreign key relationships"""
        self.add_section("FOREIGN KEY RELATIONSHIPS", level=1)
        
        query = """
        SELECT 
            s1.name AS ParentSchema,
            t1.name AS ParentTable,
            c1.name AS ParentColumn,
            fk.name AS ForeignKeyName,
            s2.name AS ReferencedSchema,
            t2.name AS ReferencedTable,
            c2.name AS ReferencedColumn
        FROM sys.foreign_keys fk
        INNER JOIN sys.tables t1 ON fk.parent_object_id = t1.object_id
        INNER JOIN sys.schemas s1 ON t1.schema_id = s1.schema_id
        INNER JOIN sys.tables t2 ON fk.referenced_object_id = t2.object_id
        INNER JOIN sys.schemas s2 ON t2.schema_id = s2.schema_id
        INNER JOIN sys.foreign_key_columns fkc ON fk.object_id = fkc.constraint_object_id
        INNER JOIN sys.columns c1 ON fkc.parent_object_id = c1.object_id AND fkc.parent_column_id = c1.column_id
        INNER JOIN sys.columns c2 ON fkc.referenced_object_id = c2.object_id AND fkc.referenced_column_id = c2.column_id
        ORDER BY s1.name, t1.name, fk.name
        """
        
        columns, results = self.db.execute_query(query)
        
        if results:
            self.add_text(f"\nTotal Foreign Key Relationships: {len(results)}")
            self.add_table(columns, results[:30], "\nForeign Keys (First 30)")
        else:
            self.add_text("\nNo foreign key relationships found.")
    
    def get_triggers(self):
        """Get trigger information"""
        self.add_section("TRIGGERS", level=1)
        
        query = """
        SELECT 
            s.name AS SchemaName,
            OBJECT_NAME(tr.parent_id) AS TableName,
            tr.name AS TriggerName,
            tr.type_desc AS TriggerType,
            tr.create_date AS CreatedDate,
            tr.modify_date AS ModifiedDate,
            tr.is_disabled AS IsDisabled
        FROM sys.triggers tr
        INNER JOIN sys.objects o ON tr.parent_id = o.object_id
        INNER JOIN sys.schemas s ON o.schema_id = s.schema_id
        WHERE tr.parent_class = 1
        AND tr.is_ms_shipped = 0
        ORDER BY s.name, OBJECT_NAME(tr.parent_id), tr.name
        """
        
        columns, results = self.db.execute_query(query)
        
        if results:
            self.add_text(f"\nTotal Triggers: {len(results)}")
            self.add_table(columns, results, "\nAll Triggers")
        else:
            self.add_text("\nNo user-defined triggers found.")
    
    def get_linked_servers(self):
        """Get linked server information"""
        self.add_section("LINKED SERVERS", level=1)
        
        query = """
        SELECT 
            name AS ServerName,
            product AS Product,
            provider AS Provider,
            data_source AS DataSource,
            is_linked AS IsLinked,
            is_remote_login_enabled AS RemoteLoginEnabled
        FROM sys.servers
        WHERE is_linked = 1
        ORDER BY name
        """
        
        try:
            columns, results = self.db.execute_query(query)
            
            if results:
                self.add_text(f"\nTotal Linked Servers: {len(results)}")
                self.add_table(columns, results, "\nLinked Servers")
            else:
                self.add_text("\nNo linked servers found.")
        except:
            self.add_text("\nCould not retrieve linked server information (may require higher permissions).")
    
    def compare_with_access_extraction(self):
        """Compare Azure SQL tables with extracted Access queries"""
        self.add_section("COMPARISON WITH MS ACCESS EXTRACTION", level=1)
        
        # Get all SQL tables
        query = """
        SELECT DISTINCT
            s.name AS SchemaName,
            t.name AS TableName
        FROM sys.tables t
        INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
        WHERE t.is_ms_shipped = 0
        ORDER BY s.name, t.name
        """
        
        columns, results = self.db.execute_query(query)
        
        if results:
            sql_tables = set(f"{row[0]}.{row[1]}" for row in results)
            self.add_text(f"\nTables in Azure SQL: {len(sql_tables)}")
            
            # Check if extracted files exist
            extracted_tables_dir = Path("msaccess/extracted/tables")
            if extracted_tables_dir.exists():
                access_tables = set()
                for file in extracted_tables_dir.glob("*_schema.txt"):
                    table_name = file.stem.replace("_schema", "")
                    access_tables.add(table_name)
                
                self.add_text(f"Tables/Views in Access Extraction: {len(access_tables)}")
                
                # Find tables in SQL but not in Access extraction
                self.add_text("\nAnalysis:")
                self.add_text(f"  - Some Access 'tables' are actually views of SQL tables")
                self.add_text(f"  - The Access database is a frontend to this SQL database")
                self.add_text(f"  - Access queries may reference these SQL tables directly")
            else:
                self.add_text("\nNote: Extracted Access tables directory not found for comparison.")
    
    def generate_report(self):
        """Generate complete assessment report"""
        print("\n" + "=" * 80)
        print("AZURE SQL DATABASE ASSESSMENT")
        print("=" * 80)
        print(f"Database: {self.db.database}")
        print(f"Server: {self.db.server}")
        print(f"Assessment Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 80 + "\n")
        
        try:
            # Connect to database
            self.db.connect()
            
            # Run all assessments
            print("Gathering database overview...")
            self.get_database_overview()
            
            print("Analyzing tables...")
            self.get_table_statistics()
            
            print("Analyzing views...")
            self.get_view_statistics()
            
            print("Analyzing stored procedures...")
            self.get_stored_procedures()
            
            print("Analyzing functions...")
            self.get_functions()
            
            print("Analyzing indexes...")
            self.get_indexes()
            
            print("Analyzing foreign keys...")
            self.get_foreign_keys()
            
            print("Analyzing triggers...")
            self.get_triggers()
            
            print("Checking linked servers...")
            self.get_linked_servers()
            
            print("Comparing with Access extraction...")
            self.compare_with_access_extraction()
            
            # Save report
            report_file = self.output_dir / f"database_assessment_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            with open(report_file, 'w', encoding='utf-8') as f:
                f.write('\n'.join(self.report_lines))
            
            print(f"\n[+] Assessment complete!")
            print(f"[+] Report saved to: {report_file}")
            
            # Print report to console
            print("\n" + "=" * 80)
            print("ASSESSMENT REPORT")
            print("=" * 80)
            print('\n'.join(self.report_lines))
            
        except Exception as e:
            print(f"\n[x] Error during assessment: {e}")
            import traceback
            traceback.print_exc()
            return False
        finally:
            self.db.disconnect()
        
        return True


if __name__ == "__main__":
    print("Starting Azure SQL Database Assessment...")
    
    assessment = DatabaseAssessment()
    success = assessment.generate_report()
    
    sys.exit(0 if success else 1)
