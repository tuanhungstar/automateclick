import pyodbc
import mysql.connector
import cx_Oracle
import sqlite3
import pandas as pd
from typing import List, Dict, Any, Optional, Union, Tuple
from datetime import datetime
import logging
from contextlib import contextmanager

class DatabaseHandler:
    """
    A flexible class for handling multiple SQL database types including MSSQL, Oracle, MySQL, and SQLite.
    
    This class provides a unified interface for connecting to and working with different database systems,
    executing queries, managing transactions, and performing common database operations.
    
    Supported Databases:
        - Microsoft SQL Server (MSSQL)
        - Oracle Database
        - MySQL/MariaDB
        - SQLite
    
    Attributes:
        db_type (str): Type of database ('mssql', 'oracle', 'mysql', 'sqlite')
        connection: Active database connection object
        cursor: Database cursor for executing queries
        operation_log (List[str]): Log of all database operations
        
    Example:
        >>> # MSSQL connection
        >>> db = DatabaseHandler('mssql')
        >>> db.connect(server='localhost', database='mydb', username='user', password='pass')
        >>> results = db.execute_query("SELECT * FROM customers")
        >>> db.close_connection()
        
        >>> # MySQL connection
        >>> db = DatabaseHandler('mysql')
        >>> db.connect(host='localhost', database='mydb', username='user', password='pass')
        >>> db.execute_non_query("INSERT INTO products (name, price) VALUES (%s, %s)", ('Product1', 99.99))
        >>> db.close_connection()
    """
    
    def __init__(self, db_type: str):
        """
        Initialize the DatabaseHandler for a specific database type.
        
        Args:
            db_type (str): Database type ('mssql', 'oracle', 'mysql', 'sqlite')
            
        Raises:
            ValueError: If unsupported database type is provided
            
        Example:
            >>> db_mssql = DatabaseHandler('mssql')
            >>> db_oracle = DatabaseHandler('oracle')
            >>> db_mysql = DatabaseHandler('mysql')
            >>> db_sqlite = DatabaseHandler('sqlite')
        """
        supported_types = ['mssql', 'oracle', 'mysql', 'sqlite']
        if db_type.lower() not in supported_types:
            raise ValueError(f"Unsupported database type: {db_type}. Supported types: {supported_types}")
        
        self.db_type = db_type.lower()
        self.connection = None
        self.cursor = None
        self.operation_log = []
        self.autocommit = True
        
        # Set up logging
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)
    
    def connect(self, **kwargs) -> bool:
        """
        Connect to the database using the appropriate driver.
        
        Args:
            **kwargs: Connection parameters specific to each database type
            
        MSSQL Parameters:
            server (str): Server name or IP address
            database (str): Database name
            username (str, optional): Username (if not using Windows authentication)
            password (str, optional): Password (if not using Windows authentication)
            port (int, optional): Port number (default: 1433)
            driver (str, optional): ODBC driver name
            trusted_connection (bool, optional): Use Windows authentication (default: False)
            
        Oracle Parameters:
            host (str): Host name or IP address
            port (int, optional): Port number (default: 1521)
            service_name (str): Oracle service name
            username (str): Username
            password (str): Password
            
        MySQL Parameters:
            host (str): Host name or IP address
            port (int, optional): Port number (default: 3306)
            database (str): Database name
            username (str): Username
            password (str): Password
            charset (str, optional): Character set (default: 'utf8')
            
        SQLite Parameters:
            database (str): Path to SQLite database file
            
        Returns:
            bool: True if connection successful, False otherwise
            
        Example:
            >>> # MSSQL with SQL Server authentication
            >>> db.connect(server='localhost', database='mydb', username='sa', password='password')
            
            >>> # MSSQL with Windows authentication
            >>> db.connect(server='localhost', database='mydb', trusted_connection=True)
            
            >>> # Oracle connection
            >>> db.connect(host='localhost', service_name='ORCL', username='hr', password='password')
            
            >>> # MySQL connection
            >>> db.connect(host='localhost', database='mydb', username='root', password='password')
            
            >>> # SQLite connection
            >>> db.connect(database='mydb.sqlite')
        """
        try:
            if self.db_type == 'mssql':
                self._connect_mssql(**kwargs)
            elif self.db_type == 'oracle':
                self._connect_oracle(**kwargs)
            elif self.db_type == 'mysql':
                self._connect_mysql(**kwargs)
            elif self.db_type == 'sqlite':
                self._connect_sqlite(**kwargs)
            
            if self.connection:
                self.cursor = self.connection.cursor()
                self._log_operation(f"Connected to {self.db_type.upper()} database")
                return True
            
            return False
            
        except Exception as e:
            self.logger.error(f"Connection failed: {str(e)}")
            print(f"✗ Connection failed: {str(e)}")
            return False
    
    def _connect_mssql(self, **kwargs):
        """Connect to Microsoft SQL Server."""
        server = kwargs.get('server')
        database = kwargs.get('database')
        username = kwargs.get('username')
        password = kwargs.get('password')
        port = kwargs.get('port', 1433)
        driver = kwargs.get('driver', 'ODBC Driver 17 for SQL Server')
        trusted_connection = kwargs.get('trusted_connection', False)
        
        if not server or not database:
            raise ValueError("Server and database are required for MSSQL connection")
        
        if trusted_connection:
            conn_str = f"DRIVER={{{driver}}};SERVER={server},{port};DATABASE={database};Trusted_Connection=yes;"
        else:
            if not username or not password:
                raise ValueError("Username and password are required for SQL Server authentication")
            conn_str = f"DRIVER={{{driver}}};SERVER={server},{port};DATABASE={database};UID={username};PWD={password};"
        
        self.connection = pyodbc.connect(conn_str)
    
    def _connect_oracle(self, **kwargs):
        """Connect to Oracle Database."""
        host = kwargs.get('host')
        port = kwargs.get('port', 1521)
        service_name = kwargs.get('service_name')
        username = kwargs.get('username')
        password = kwargs.get('password')
        
        if not all([host, service_name, username, password]):
            raise ValueError("Host, service_name, username, and password are required for Oracle connection")
        
        dsn = cx_Oracle.makedsn(host, port, service_name=service_name)
        self.connection = cx_Oracle.connect(username, password, dsn)
    
    def _connect_mysql(self, **kwargs):
        """Connect to MySQL Database."""
        host = kwargs.get('host')
        port = kwargs.get('port', 3306)
        database = kwargs.get('database')
        username = kwargs.get('username')
        password = kwargs.get('password')
        charset = kwargs.get('charset', 'utf8')
        
        if not all([host, database, username, password]):
            raise ValueError("Host, database, username, and password are required for MySQL connection")
        
        self.connection = mysql.connector.connect(
            host=host,
            port=port,
            database=database,
            user=username,
            password=password,
            charset=charset
        )
    
    def _connect_sqlite(self, **kwargs):
        """Connect to SQLite Database."""
        database = kwargs.get('database')
        
        if not database:
            raise ValueError("Database file path is required for SQLite connection")
        
        self.connection = sqlite3.connect(database)
    
    def execute_query(self, query: str, params: Tuple = None, fetch_size: int = None) -> List[Dict[str, Any]]:
        """
        Execute a SELECT query and return results.
        
        Args:
            query (str): SQL SELECT query to execute
            params (Tuple, optional): Parameters for parameterized queries
            fetch_size (int, optional): Number of rows to fetch at once (for large results)
            
        Returns:
            List[Dict[str, Any]]: List of dictionaries representing query results
            
        Example:
            >>> # Simple query
            >>> results = db.execute_query("SELECT * FROM customers")
            >>> for row in results:
            ...     print(f"Customer: {row['name']}")
            
            >>> # Parameterized query
            >>> results = db.execute_query(
            ...     "SELECT * FROM customers WHERE city = ? AND age > ?", 
            ...     ('New York', 25)
            ... )
        """
        try:
            if not self.connection or not self.cursor:
                raise Exception("No database connection. Call connect() first.")
            
            if params:
                self.cursor.execute(query, params)
            else:
                self.cursor.execute(query)
            
            # Get column names
            if self.db_type == 'oracle':
                columns = [desc[0] for desc in self.cursor.description]
            else:
                columns = [desc[0] for desc in self.cursor.description]
            
            # Fetch results
            if fetch_size:
                rows = self.cursor.fetchmany(fetch_size)
            else:
                rows = self.cursor.fetchall()
            
            # Convert to list of dictionaries
            results = []
            for row in rows:
                row_dict = {}
                for i, value in enumerate(row):
                    row_dict[columns[i]] = value
                results.append(row_dict)
            
            self._log_operation(f"Executed query: {query[:100]}... (returned {len(results)} rows)")
            return results
            
        except Exception as e:
            self.logger.error(f"Query execution failed: {str(e)}")
            print(f"✗ Query execution failed: {str(e)}")
            return []
    
    def execute_non_query(self, query: str, params: Tuple = None) -> int:
        """
        Execute a non-SELECT query (INSERT, UPDATE, DELETE).
        
        Args:
            query (str): SQL query to execute
            params (Tuple, optional): Parameters for parameterized queries
            
        Returns:
            int: Number of affected rows, -1 if failed
            
        Example:
            >>> # Insert data
            >>> affected = db.execute_non_query(
            ...     "INSERT INTO customers (name, email) VALUES (?, ?)",
            ...     ('John Doe', 'john@example.com')
            ... )
            >>> print(f"Inserted {affected} rows")
            
            >>> # Update data
            >>> affected = db.execute_non_query(
            ...     "UPDATE customers SET email = ? WHERE name = ?",
            ...     ('newemail@example.com', 'John Doe')
            ... )
        """
        try:
            if not self.connection or not self.cursor:
                raise Exception("No database connection. Call connect() first.")
            
            if params:
                self.cursor.execute(query, params)
            else:
                self.cursor.execute(query)
            
            affected_rows = self.cursor.rowcount
            
            if self.autocommit:
                self.connection.commit()
            
            self._log_operation(f"Executed non-query: {query[:100]}... (affected {affected_rows} rows)")
            return affected_rows
            
        except Exception as e:
            self.logger.error(f"Non-query execution failed: {str(e)}")
            print(f"✗ Non-query execution failed: {str(e)}")
            if not self.autocommit:
                self.connection.rollback()
            return -1
    
    def execute_many(self, query: str, params_list: List[Tuple]) -> int:
        """
        Execute the same query multiple times with different parameters.
        
        Args:
            query (str): SQL query to execute
            params_list (List[Tuple]): List of parameter tuples
            
        Returns:
            int: Total number of affected rows, -1 if failed
            
        Example:
            >>> # Bulk insert
            >>> data = [
            ...     ('Alice', 'alice@example.com'),
            ...     ('Bob', 'bob@example.com'),
            ...     ('Charlie', 'charlie@example.com')
            ... ]
            >>> affected = db.execute_many(
            ...     "INSERT INTO customers (name, email) VALUES (?, ?)",
            ...     data
            ... )
            >>> print(f"Inserted {affected} rows")
        """
        try:
            if not self.connection or not self.cursor:
                raise Exception("No database connection. Call connect() first.")
            
            if self.db_type == 'mysql':
                self.cursor.executemany(query, params_list)
            else:
                for params in params_list:
                    self.cursor.execute(query, params)
            
            affected_rows = self.cursor.rowcount
            
            if self.autocommit:
                self.connection.commit()
            
            self._log_operation(f"Executed bulk operation: {len(params_list)} operations (affected {affected_rows} rows)")
            return affected_rows
            
        except Exception as e:
            self.logger.error(f"Bulk execution failed: {str(e)}")
            print(f"✗ Bulk execution failed: {str(e)}")
            if not self.autocommit:
                self.connection.rollback()
            return -1
    
    def query_to_dataframe(self, query: str, params: Tuple = None) -> pd.DataFrame:
        """
        Execute a query and return results as a pandas DataFrame.
        
        Args:
            query (str): SQL SELECT query to execute
            params (Tuple, optional): Parameters for parameterized queries
            
        Returns:
            pd.DataFrame: Query results as DataFrame
            
        Example:
            >>> df = db.query_to_dataframe("SELECT * FROM customers WHERE age > ?", (25,))
            >>> print(df.head())
            >>> print(f"Found {len(df)} customers")
        """
        try:
            if not self.connection:
                raise Exception("No database connection. Call connect() first.")
            
            if params:
                df = pd.read_sql_query(query, self.connection, params=params)
            else:
                df = pd.read_sql_query(query, self.connection)
            
            self._log_operation(f"Query to DataFrame: {query[:100]}... (returned {len(df)} rows)")
            return df
            
        except Exception as e:
            self.logger.error(f"DataFrame query failed: {str(e)}")
            print(f"✗ DataFrame query failed: {str(e)}")
            return pd.DataFrame()
    
    def dataframe_to_table(self, df: pd.DataFrame, table_name: str, 
                          if_exists: str = 'append', method: str = None) -> bool:
        """
        Insert a pandas DataFrame into a database table.
        
        Args:
            df (pd.DataFrame): DataFrame to insert
            table_name (str): Target table name
            if_exists (str): Action if table exists ('fail', 'replace', 'append') (default: 'append')
            method (str, optional): Insertion method for performance optimization
            
        Returns:
            bool: True if insertion successful, False otherwise
            
        Example:
            >>> df = pd.DataFrame({
            ...     'name': ['Alice', 'Bob', 'Charlie'],
            ...     'age': [25, 30, 35],
            ...     'city': ['NYC', 'LA', 'Chicago']
            ... })
            >>> success = db.dataframe_to_table(df, 'customers', if_exists='append')
            >>> if success:
            ...     print(f"Inserted {len(df)} rows into customers table")
        """
        try:
            if not self.connection:
                raise Exception("No database connection. Call connect() first.")
            
            df.to_sql(table_name, self.connection, if_exists=if_exists, 
                     index=False, method=method)
            
            self._log_operation(f"DataFrame to table '{table_name}': {len(df)} rows inserted")
            return True
            
        except Exception as e:
            self.logger.error(f"DataFrame insertion failed: {str(e)}")
            print(f"✗ DataFrame insertion failed: {str(e)}")
            return False
    
    def get_table_info(self, table_name: str) -> Dict[str, Any]:
        """
        Get information about a table structure.
        
        Args:
            table_name (str): Name of the table to analyze
            
        Returns:
            Dict[str, Any]: Dictionary containing table information
            
        Example:
            >>> info = db.get_table_info('customers')
            >>> print(f"Table has {info['column_count']} columns")
            >>> for col in info['columns']:
            ...     print(f"Column: {col['name']} ({col['type']})")
        """
        try:
            if not self.connection:
                raise Exception("No database connection. Call connect() first.")
            
            if self.db_type == 'mssql':
                query = """
                SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE, CHARACTER_MAXIMUM_LENGTH
                FROM INFORMATION_SCHEMA.COLUMNS 
                WHERE TABLE_NAME = ?
                ORDER BY ORDINAL_POSITION
                """
            elif self.db_type == 'mysql':
                query = """
                SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE, CHARACTER_MAXIMUM_LENGTH
                FROM INFORMATION_SCHEMA.COLUMNS 
                WHERE TABLE_NAME = %s AND TABLE_SCHEMA = DATABASE()
                ORDER BY ORDINAL_POSITION
                """
            elif self.db_type == 'oracle':
                query = """
                SELECT COLUMN_NAME, DATA_TYPE, NULLABLE, DATA_LENGTH
                FROM USER_TAB_COLUMNS 
                WHERE TABLE_NAME = UPPER(:1)
                ORDER BY COLUMN_ID
                """
            elif self.db_type == 'sqlite':
                query = f"PRAGMA table_info({table_name})"
            
            columns_info = self.execute_query(query, (table_name,))
            
            # Get row count
            count_query = f"SELECT COUNT(*) as row_count FROM {table_name}"
            count_result = self.execute_query(count_query)
            row_count = count_result[0]['row_count'] if count_result else 0
            
            table_info = {
                'table_name': table_name,
                'column_count': len(columns_info),
                'row_count': row_count,
                'columns': columns_info
            }
            
            self._log_operation(f"Retrieved table info for '{table_name}'")
            return table_info
            
        except Exception as e:
            self.logger.error(f"Failed to get table info: {str(e)}")
            print(f"✗ Failed to get table info: {str(e)}")
            return {}
    
    def get_table_list(self) -> List[str]:
        """
        Get a list of all tables in the database.
        
        Returns:
            List[str]: List of table names
            
        Example:
            >>> tables = db.get_table_list()
            >>> print(f"Database contains {len(tables)} tables:")
            >>> for table in tables:
            ...     print(f"  - {table}")
        """
        try:
            if not self.connection:
                raise Exception("No database connection. Call connect() first.")
            
            if self.db_type == 'mssql':
                query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'"
            elif self.db_type == 'mysql':
                query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = DATABASE()"
            elif self.db_type == 'oracle':
                query = "SELECT TABLE_NAME FROM USER_TABLES"
            elif self.db_type == 'sqlite':
                query = "SELECT name FROM sqlite_master WHERE type='table'"
            
            results = self.execute_query(query)
            
            if self.db_type == 'sqlite':
                tables = [row['name'] for row in results]
            else:
                tables = [row['TABLE_NAME'] for row in results]
            
            self._log_operation(f"Retrieved table list: {len(tables)} tables found")
            return tables
            
        except Exception as e:
            self.logger.error(f"Failed to get table list: {str(e)}")
            print(f"✗ Failed to get table list: {str(e)}")
            return []
    
    @contextmanager
    def transaction(self):
        """
        Context manager for database transactions.
        
        Automatically commits on success or rolls back on error.
        
        Example:
            >>> with db.transaction():
            ...     db.execute_non_query("INSERT INTO customers (name) VALUES (?)", ('Alice',))
            ...     db.execute_non_query("INSERT INTO orders (customer_id) VALUES (?)", (1,))
            ...     # Both operations will be committed together
        """
        if not self.connection:
            raise Exception("No database connection. Call connect() first.")
        
        old_autocommit = self.autocommit
        self.autocommit = False
        
        try:
            yield self
            self.connection.commit()
            self._log_operation("Transaction committed")
        except Exception as e:
            self.connection.rollback()
            self._log_operation(f"Transaction rolled back: {str(e)}")
            raise
        finally:
            self.autocommit = old_autocommit
    
    def backup_table(self, table_name: str, backup_file: str, format: str = 'csv') -> bool:
        """
        Backup a table to a file.
        
        Args:
            table_name (str): Name of the table to backup
            backup_file (str): Path to backup file
            format (str): Backup format ('csv', 'json', 'excel') (default: 'csv')
            
        Returns:
            bool: True if backup successful, False otherwise
            
        Example:
            >>> success = db.backup_table('customers', 'customers_backup.csv', 'csv')
            >>> if success:
            ...     print("Table backed up successfully")
        """
        try:
            df = self.query_to_dataframe(f"SELECT * FROM {table_name}")
            
            if df.empty:
                print(f"✗ Table '{table_name}' is empty or doesn't exist")
                return False
            
            if format.lower() == 'csv':
                df.to_csv(backup_file, index=False)
            elif format.lower() == 'json':
                df.to_json(backup_file, orient='records', indent=2)
            elif format.lower() == 'excel':
                df.to_excel(backup_file, index=False)
            else:
                print(f"✗ Unsupported format: {format}")
                return False
            
            self._log_operation(f"Backed up table '{table_name}' to '{backup_file}' ({format})")
            return True
            
        except Exception as e:
            self.logger.error(f"Backup failed: {str(e)}")
            print(f"✗ Backup failed: {str(e)}")
            return False
    
    def test_connection(self) -> Dict[str, Any]:
        """
        Test the database connection and return connection details.
        
        Returns:
            Dict[str, Any]: Connection test results and database information
            
        Example:
            >>> test_result = db.test_connection()
            >>> if test_result['connected']:
            ...     print(f"Connected to {test_result['database_type']} database")
            ...     print(f"Database version: {test_result['version']}")
        """
        try:
            if not self.connection:
                return {'connected': False, 'error': 'No connection established'}
            
            # Test with a simple query
            if self.db_type == 'mssql':
                version_query = "SELECT @@VERSION as version"
            elif self.db_type == 'mysql':
                version_query = "SELECT VERSION() as version"
            elif self.db_type == 'oracle':
                version_query = "SELECT * FROM V$VERSION WHERE ROWNUM = 1"
            elif self.db_type == 'sqlite':
                version_query = "SELECT sqlite_version() as version"
            
            result = self.execute_query(version_query)
            version = result[0]['version'] if result else 'Unknown'
            
            test_result = {
                'connected': True,
                'database_type': self.db_type.upper(),
                'version': version,
                'tables_count': len(self.get_table_list())
            }
            
            self._log_operation("Connection test completed successfully")
            return test_result
            
        except Exception as e:
            self.logger.error(f"Connection test failed: {str(e)}")
            return {
                'connected': False,
                'error': str(e)
            }
    
    def close_connection(self) -> bool:
        """
        Close the database connection.
        
        Returns:
            bool: True if connection closed successfully, False otherwise
            
        Example:
            >>> db.close_connection()
        """
        try:
            if self.cursor:
                self.cursor.close()
                self.cursor = None
            
            if self.connection:
                self.connection.close()
                self.connection = None
            
            self._log_operation("Database connection closed")
            return True
            
        except Exception as e:
            self.logger.error(f"Error closing connection: {str(e)}")
            print(f"✗ Error closing connection: {str(e)}")
            return False
    
    def get_operation_log(self) -> List[str]:
        """
        Get the complete log of all database operations.
        
        Returns:
            List[str]: List of all operations with timestamps
            
        Example:
            >>> log = db.get_operation_log()
            >>> for entry in log:
            ...     print(entry)
        """
        return self.operation_log.copy()
    
    def clear_log(self):
        """
        Clear the operation log.
        
        Example:
            >>> db.clear_log()
        """
        self.operation_log = []
        print("✓ Operation log cleared")
    
    def _log_operation(self, operation: str):
        """
        Internal method to log operations with timestamps.
        
        Args:
            operation (str): Description of the operation performed
        """
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.operation_log.append(f"{timestamp} - {operation}")

# Example usage and demonstrations
if __name__ == "__main__":
    """
    Comprehensive examples demonstrating database operations across different database types.
    """
    
    print("=" * 60)
    print("DATABASE HANDLER DEMONSTRATIONS")
    print("=" * 60)
    
    # Example 1: SQLite (easiest to demonstrate)
    print("\n1. SQLite Example:")
    db_sqlite = DatabaseHandler('sqlite')
    
    if db_sqlite.connect(database='example.db'):
        print("✓ Connected to SQLite database")
        
        # Create a sample table
        create_table_sql = """
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            department TEXT,
            salary REAL,
            hire_date DATE
        )
        """
        
        db_sqlite.execute_non_query(create_table_sql)
        print("✓ Created employees table")
        
        # Insert sample data
        sample_data = [
            ('Alice Johnson', 'IT', 75000, '2023-01-15'),
            ('Bob Smith', 'HR', 65000, '2023-02-20'),
            ('Charlie Brown', 'Finance', 80000, '2023-03-10')
        ]
        
        for emp_data in sample_data:
            db_sqlite.execute_non_query(
                "INSERT INTO employees (name, department, salary, hire_date) VALUES (?, ?, ?, ?)",
                emp_data
            )
        
        print(f"✓ Inserted {len(sample_data)} employees")
        
        # Query data
        results = db_sqlite.execute_query("SELECT * FROM employees WHERE salary > ?", (70000,))
        print(f"✓ Found {len(results)} high-salary employees:")
        for emp in results:
            print(f"  - {emp['name']}: ${emp['salary']}")
        
        # Get table info
        table_info = db_sqlite.get_table_info('employees')
        print(f"✓ Table info: {table_info['column_count']} columns, {table_info['row_count']} rows")
        
        # Query to DataFrame
        df = db_sqlite.query_to_dataframe("SELECT department, AVG(salary) as avg_salary FROM employees GROUP BY department")
        print("✓ Average salary by department:")
        print(df)
        
        # Test connection
        test_result = db_sqlite.test_connection()
        print(f"✓ Connection test: {test_result}")
        
        # Backup table
        backup_success = db_sqlite.backup_table('employees', 'employees_backup.csv')
        if backup_success:
            print("✓ Table backed up successfully")
        
        db_sqlite.close_connection()
        print("✓ SQLite connection closed")
    
    # Example 2: MySQL (commented out as it requires MySQL server)
    print("\n2. MySQL Example (commented - requires MySQL server):")
    print("""
    # MySQL connection example:
    db_mysql = DatabaseHandler('mysql')
    
    if db_mysql.connect(
        host='localhost',
        database='mycompany',
        username='root',
        password='password',
        port=3306
    ):
        print("✓ Connected to MySQL database")
        
        # Execute queries
        results = db_mysql.execute_query("SELECT * FROM customers LIMIT 10")
        
        # Bulk insert
        data = [('John', 'john@email.com'), ('Jane', 'jane@email.com')]
        db_mysql.execute_many(
            "INSERT INTO customers (name, email) VALUES (%s, %s)",
            data
        )
        
        db_mysql.close_connection()
    """)
    
    # Example 3: MSSQL (commented out as it requires SQL Server)
    print("\n3. MSSQL Example (commented - requires SQL Server):")
    print("""
    # MSSQL connection example:
    db_mssql = DatabaseHandler('mssql')
    
    # SQL Server Authentication
    if db_mssql.connect(
        server='localhost',
        database='CompanyDB',
        username='sa',
        password='password'
    ):
        print("✓ Connected to SQL Server")
        
        # Or Windows Authentication
        # db_mssql.connect(
        #     server='localhost',
        #     database='CompanyDB',
        #     trusted_connection=True
        # )
        
        # Transaction example
        with db_mssql.transaction():
            db_mssql.execute_non_query(
                "INSERT INTO orders (customer_id, total) VALUES (?, ?)",
                (123, 999.99)
            )
            db_mssql.execute_non_query(
                "UPDATE customers SET last_order_date = GETDATE() WHERE id = ?",
                (123,)
            )
        
        db_mssql.close_connection()
    """)
    
    # Example 4: Oracle (commented out as it requires Oracle database)
    print("\n4. Oracle Example (commented - requires Oracle database):")
    print("""
    # Oracle connection example:
    db_oracle = DatabaseHandler('oracle')
    
    if db_oracle.connect(
        host='localhost',
        port=1521,
        service_name='ORCL',
        username='hr',
        password='password'
    ):
        print("✓ Connected to Oracle database")
        
        # Query with Oracle-specific syntax
        results = db_oracle.execute_query(
            "SELECT * FROM employees WHERE ROWNUM <= :1",
            (10,)
        )
        
        # Get table list
        tables = db_oracle.get_table_list()
        print(f"Found {len(tables)} tables in Oracle database")
        
        db_oracle.close_connection()
    """)
    
    # Show operation log for SQLite example
    print("\n" + "="*60)
    print("OPERATION LOG (SQLite Example)")
    print("="*60)
    
    log = db_sqlite.get_operation_log()
    for entry in log:
        print(entry)