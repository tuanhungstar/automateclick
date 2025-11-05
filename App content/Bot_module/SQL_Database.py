#!/usr/bin/env python3

import mysql.connector
import pyodbc  # For MS SQL Server
import cx_Oracle  # For Oracle (Replaced oracledb)
import os
import sys
script_dir = os.path.dirname(os.path.abspath(__file__))
my_lib_dir = os.path.join(script_dir, "my_lib")
if my_lib_dir not in sys.path:
    sys.path.insert(0, my_lib_dir)
    
# --- IMPORT FROM YOUR SHARED LIBRARY ---
# This assumes 'my_lib' is in your Python path
from my_lib.shared_context import ExecutionContext as Context
class SQL_Tool:
# --- Connection Methods ---
    def __init__(self, context: Context):
        """Initializes the Bot_utility class.

        Args:
            context (Context): A shared context object for logging and state management
                               across different bot components.
        """
        self.context = context # Store the shared context
        self.log_prefix = f"[{self.__class__.__name__}]"
        pass
    def connect_mysql(self,host, user, password, database, port=3306):
        """
        Builds a connection to a MySQL database.
        
        Args:
            context (ExecutionContext): The shared execution context for logging.
            host (str): The database server hostname or IP.
            user (str): The username.
            password (str): The password.
            database (str): The name of the database to connect to.
            port (int): The port number (default is 3306).
            
        Returns:
            mysql.connector.connection.MySQLConnection or None: The connection object or None if failed.
        """
        self.self.context.add_log(f"Attempting to connect to MySQL (Host: {host}, DB: {database})...")
        try:
            connection = mysql.connector.connect(
                host=host,
                user=user,
                password=password,
                database=database,
                port=port
            )
            if connection.is_connected():
                self.context.add_log("✅ MySQL Connection Successful.")
                return connection
        except mysql.connector.Error as e:
            self.self.context.add_log(f"❌ MySQL Connection Error: {e}")
            return None
        return None

    def connect_mssql(self,server, database, username, password, driver='{ODBC Driver 17 for SQL Server}'):
        """
        Builds a connection to a Microsoft SQL Server database using pyodbc.
        
        Args:
            context (ExecutionContext): The shared execution context for logging.
            server (str): The server hostname or IP.
            database (str): The name of the database.
            username (str): The username.
            password (str): The password.
            driver (str): The ODBC driver string.
                          
        Returns:
            pyodbc.Connection or None: The connection object or None if failed.
        """
        self.context.add_log(f"Attempting to connect to MS SQL Server (Server: {server}, DB: {database})...")
        connection_string = (
            f"DRIVER={driver};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"UID={username};"
            f"PWD={password}"
        )
        try:
            connection = pyodbc.connect(connection_string)
            self.context.add_log("✅ MS SQL Server Connection Successful.")
            return connection
        except pyodbc.Error as e:
            self.context.add_log(f"❌ MS SQL Server Connection Error: {e}")
            return None
        return None

    def connect_oracle(self,username, password, dsn):
        """
        Builds a connection to an Oracle database using cx_Oracle.
        
        NOTE: Requires the Oracle Instant Client to be installed and
        configured on your system. You may need to call
        cx_Oracle.init_oracle_client(lib_dir="/path/to/instantclient")
        at the start of your application.
        
        Args:
            context (ExecutionContext): The shared execution context for logging.
            username (str): The database username.
            password (str): The password.
            dsn (str): The Data Source Name or "Easy Connect" string.
                       Example: 'localhost:1521/XEPDB1' (host:port/service_name)
                       
        Returns:
            cx_Oracle.Connection or None: The connection object or None if failed.
        """
        self.context.add_log(f"Attempting to connect to Oracle (DSN: {dsn})...")
        try:
            connection = cx_Oracle.connect(
                user=username,
                password=password,
                dsn=dsn
            )
            self.context.add_log("✅ Oracle Connection Successful.")
            return connection
        except cx_Oracle.Error as e:
            self.context.add_log(f"❌ Oracle Connection Error: {e}")
            return None
        return None

    # --- Data Operation Methods ---

    def read_data(self,connection, query):
        """
        Reads data from the database using a SELECT query.
        (This function is generic and works for cx_Oracle too)
        
        Args:
            context (ExecutionContext): The shared execution context for logging.
            connection: A valid database connection object (from any connect_* method).
            query (str): The SQL SELECT statement to execute.
            
        Returns:
            list: A list of tuples, where each tuple is a row. Returns empty list on failure.
        """
        if not connection:
            self.context.add_log("❌ Read Error: Invalid connection.")
            return []
            
        cursor = None
        try:
            cursor = connection.cursor()
            cursor.execute(query)
            results = cursor.fetchall()
            self.context.add_log(f"Read successful, {len(results)} rows fetched.")
            return results
            
        except Exception as e:
            self.context.add_log(f"❌ Read Error: {e}")
            return []
            
        finally:
            if cursor:
                cursor.close()

    def write_data(self,connection, query, data=None):
        """
        Writes data to the database (INSERT, UPDATE, DELETE).
        (This function is generic and works for cx_Oracle too)
        
        Args:
            context (ExecutionContext): The shared execution context for logging.
            connection: A valid database connection object.
            query (str): The SQL statement (e.g., INSERT, UPDATE).
            data (tuple, optional): A tuple of data to be used with
                                    parameterized queries.
                                    
        Returns:
            bool: True if successful, False otherwise.
        """
        if not connection:
            self.context.add_log("❌ Write Error: Invalid connection.")
            return False
            
        cursor = None
        try:
            cursor = connection.cursor()
            
            if data:
                cursor.execute(query, data)
            else:
                cursor.execute(query)
                
            connection.commit()
            self.context.add_log(f"Write successful. {cursor.rowcount} rows affected.")
            return True
            
        except Exception as e:
            self.context.add_log(f"❌ Write Error: {e}")
            try:
                connection.rollback()
                self.context.add_log("Transaction rolled back.")
            except Exception as rb_e:
                self.context.add_log(f"Error during rollback: {rb_e}")
            return False
            
        finally:
            if cursor:
                cursor.close()

