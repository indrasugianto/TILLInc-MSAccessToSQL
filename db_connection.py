"""
Azure SQL Database Connection Utility
Provides secure connection management for the TILLDBWEB_Prod database
"""

import os
import pyodbc
from dotenv import load_dotenv
from typing import Optional
import sys

# Load environment variables from .env file
load_dotenv()

class AzureSQLConnection:
    """Manages connection to Azure SQL Database"""
    
    def __init__(self):
        """Initialize connection parameters from environment variables"""
        self.server = os.getenv('AZURE_SQL_SERVER')
        self.database = os.getenv('AZURE_SQL_DATABASE')
        self.username = os.getenv('AZURE_SQL_USER')
        self.password = os.getenv('AZURE_SQL_PASSWORD')
        self.port = os.getenv('AZURE_SQL_PORT', '1433')
        self.connection = None
        self.cursor = None
        
        # Validate required environment variables
        if not all([self.server, self.database, self.username, self.password]):
            raise ValueError(
                "Missing required environment variables. "
                "Please ensure .env file exists with all required fields."
            )
    
    def get_connection_string(self) -> str:
        """Build ODBC connection string for Azure SQL"""
        # Try to use the latest ODBC driver available
        drivers = [
            'ODBC Driver 18 for SQL Server',
            'ODBC Driver 17 for SQL Server',
            'ODBC Driver 13 for SQL Server',
            'SQL Server'
        ]
        
        available_driver = None
        for driver in drivers:
            if driver in [x for x in pyodbc.drivers()]:
                available_driver = driver
                break
        
        if not available_driver:
            raise Exception(
                "No suitable ODBC driver found. Please install "
                "'ODBC Driver 18 for SQL Server' or similar."
            )
        
        connection_string = (
            f"DRIVER={{{available_driver}}};"
            f"SERVER={self.server},{self.port};"
            f"DATABASE={self.database};"
            f"UID={self.username};"
            f"PWD={self.password};"
            f"Encrypt=yes;"
            f"TrustServerCertificate=no;"
            f"Connection Timeout=30;"
        )
        
        return connection_string
    
    def connect(self) -> pyodbc.Connection:
        """Establish connection to Azure SQL Database"""
        try:
            connection_string = self.get_connection_string()
            self.connection = pyodbc.connect(connection_string)
            self.cursor = self.connection.cursor()
            print(f"[+] Connected to Azure SQL Database: {self.database}")
            return self.connection
        except pyodbc.Error as e:
            print(f"[x] Error connecting to database: {e}")
            raise
    
    def disconnect(self):
        """Close database connection"""
        if self.cursor:
            self.cursor.close()
        if self.connection:
            self.connection.close()
            print("[+] Database connection closed")
    
    def execute_query(self, query: str, params: Optional[tuple] = None):
        """Execute a SQL query and return results"""
        if not self.connection:
            self.connect()
        
        try:
            if params:
                self.cursor.execute(query, params)
            else:
                self.cursor.execute(query)
            
            # Try to fetch results (SELECT queries)
            try:
                columns = [column[0] for column in self.cursor.description]
                results = self.cursor.fetchall()
                return columns, results
            except:
                # Non-SELECT query (INSERT, UPDATE, DELETE)
                self.connection.commit()
                return None, None
                
        except pyodbc.Error as e:
            print(f"[x] Error executing query: {e}")
            raise
    
    def __enter__(self):
        """Context manager entry"""
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit"""
        self.disconnect()


def test_connection():
    """Test database connection"""
    print("=" * 60)
    print("Testing Azure SQL Database Connection")
    print("=" * 60)
    
    try:
        with AzureSQLConnection() as db:
            print(f"\nServer: {db.server}")
            print(f"Database: {db.database}")
            print(f"User: {db.username}")
            
            # Test query
            print("\nExecuting test query...")
            columns, results = db.execute_query("SELECT @@VERSION AS Version")
            
            if results:
                print("\n[+] SQL Server Version:")
                print(f"  {results[0][0]}")
            
            # Get database info
            print("\nGetting database information...")
            columns, results = db.execute_query("""
                SELECT 
                    name AS DatabaseName,
                    state_desc AS State,
                    recovery_model_desc AS RecoveryModel,
                    compatibility_level AS CompatibilityLevel
                FROM sys.databases
                WHERE name = ?
            """, (db.database,))
            
            if results:
                print("\n[+] Database Information:")
                for i, col in enumerate(columns):
                    print(f"  {col}: {results[0][i]}")
        
        print("\n" + "=" * 60)
        print("[SUCCESS] Connection Test Successful!")
        print("=" * 60)
        return True
        
    except Exception as e:
        print("\n" + "=" * 60)
        print(f"[ERROR] Connection Test Failed: {e}")
        print("=" * 60)
        return False


if __name__ == "__main__":
    # Run connection test
    success = test_connection()
    sys.exit(0 if success else 1)
