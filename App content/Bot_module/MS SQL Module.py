# File: Bot_module/mssql_module.py
from datetime import datetime, timedelta 
import sys
from typing import Optional, List, Dict, Any
from urllib.parse import quote_plus
import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy import event

# --- PyQt6 Imports ---
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QDialogButtonBox,
    QComboBox, QWidget, QGroupBox, QRadioButton, QMessageBox, QLabel,
    QCheckBox, QApplication, QTextEdit, QTreeWidget, QTreeWidgetItem,
    QHBoxLayout, QHeaderView, QTreeWidgetItemIterator,
    QListWidget, QSpinBox, QTabWidget  # <- Add these three new widgets
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer

# --- Database Imports ---
try:
    import pyodbc
except ImportError:
    print("Warning: 'pyodbc' library not found. Please install it using: pip install pyodbc")
    pyodbc = None

# --- Main App Imports (Fallback for standalone testing) ---
try:
    from my_lib.shared_context import ExecutionContext
except ImportError:
    print("Warning: Could not import main app libraries. Using fallbacks.")
    class ExecutionContext:
        def add_log(self, message: str): print(message)
        def get_variable(self, name: str):
            if "conn" in name: return create_engine("sqlite:///:memory:") # Return a dummy engine
            if "df" in name: return pd.DataFrame({'id': [1], 'value': ['a'], 'type': ['t']})
            return None

#
# --- [CLASS 1] HELPER: The GUI Dialog for MS SQL Connection ---
#
class _MssqlConfigDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None,
                 initial_variable: Optional[str] = None):
        super().__init__(parent)
        self.setWindowTitle("MS SQL Database Configuration")
        self.setMinimumWidth(450)
        self.global_variables = global_variables
        self.initial_config = initial_config
        main_layout = QVBoxLayout(self)
        connection_group = QGroupBox("Connection Details")
        form_layout = QFormLayout(connection_group)
        self.host_edit = QLineEdit()
        self.host_edit.setPlaceholderText("e.g., SERVER_NAME\\SQLEXPRESS or 192.168.1.100")
        form_layout.addRow("Host / Server:", self.host_edit)
        self.db_name_edit = QLineEdit()
        self.db_name_edit.setPlaceholderText("e.g., MyDatabase")
        form_layout.addRow("Database Name:", self.db_name_edit)
        main_layout.addWidget(connection_group)
        auth_group = QGroupBox("Authentication")
        auth_layout = QVBoxLayout(auth_group)
        self.win_auth_radio = QRadioButton("Windows Authentication")
        self.sql_auth_radio = QRadioButton("SQL Server Authentication")
        auth_layout.addWidget(self.win_auth_radio)
        auth_layout.addWidget(self.sql_auth_radio)
        self.sql_auth_widget = QWidget()
        sql_auth_form = QFormLayout(self.sql_auth_widget)
        sql_auth_form.setContentsMargins(20, 5, 5, 5)
        self.username_edit = QLineEdit()
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        sql_auth_form.addRow("Username:", self.username_edit)
        sql_auth_form.addRow("Password:", self.password_edit)
        auth_layout.addWidget(self.sql_auth_widget)
        main_layout.addWidget(auth_group)
        self.test_button = QPushButton("Test Connection")
        main_layout.addWidget(self.test_button)
        assignment_group = QGroupBox("Assign Connection to Variable")
        assign_layout = QVBoxLayout(assignment_group)
        self.assign_checkbox = QCheckBox("Assign successful connection object to a variable")
        assign_layout.addWidget(self.assign_checkbox)
        self.new_var_radio = QRadioButton("New Variable Name:")
        self.new_var_input = QLineEdit("db_connection")
        self.existing_var_radio = QRadioButton("Existing Variable:")
        self.existing_var_combo = QComboBox()
        self.existing_var_combo.addItem("-- Select Variable --")
        self.existing_var_combo.addItems(self.global_variables)
        assign_form = QFormLayout()
        assign_form.addRow(self.new_var_radio, self.new_var_input)
        assign_form.addRow(self.existing_var_radio, self.existing_var_combo)
        assign_layout.addLayout(assign_form)
        main_layout.addWidget(assignment_group)
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)
        self.win_auth_radio.toggled.connect(self._toggle_auth_widgets)
        self.assign_checkbox.toggled.connect(self._toggle_assignment_widgets)
        self.new_var_radio.toggled.connect(self._toggle_assignment_widgets)
        self.test_button.clicked.connect(self._test_connection)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.win_auth_radio.setChecked(True)
        self.assign_checkbox.setChecked(True)
        self.new_var_radio.setChecked(True)
        self._toggle_auth_widgets()
        self._toggle_assignment_widgets()
        self._populate_from_initial_config(initial_config, initial_variable)

    def _toggle_auth_widgets(self):
        is_sql_auth = self.sql_auth_radio.isChecked()
        self.sql_auth_widget.setVisible(is_sql_auth)

    def _toggle_assignment_widgets(self):
        is_assign_enabled = self.assign_checkbox.isChecked()
        self.new_var_radio.setVisible(is_assign_enabled)
        self.new_var_input.setVisible(is_assign_enabled and self.new_var_radio.isChecked())
        self.existing_var_radio.setVisible(is_assign_enabled)
        self.existing_var_combo.setVisible(is_assign_enabled and self.existing_var_radio.isChecked())

    def _get_connection_string(self) -> Optional[str]:
        host = self.host_edit.text().strip()
        database = self.db_name_edit.text().strip()
        if not host or not database:
            QMessageBox.warning(self, "Input Error", "Host and Database Name cannot be empty.")
            return None
        driver = ""
        if pyodbc:
            drivers = [d for d in pyodbc.drivers() if "sql server" in d.lower()]
            if any("18" in d for d in drivers): driver = [d for d in drivers if "18" in d][0]
            elif any("17" in d for d in drivers): driver = [d for d in drivers if "17" in d][0]
            elif drivers: driver = drivers[-1]
        if not driver:
             QMessageBox.critical(self, "Driver Error", "No MS SQL Server ODBC driver found. Please install 'ODBC Driver 17 for SQL Server' or newer.")
             return None
        
        driver_for_url = driver.replace(' ', '+')
        conn_url = f"mssql+pyodbc:///?odbc_connect="
        
        params = {"DRIVER": driver, "SERVER": host, "DATABASE": database}
        if self.win_auth_radio.isChecked():
            params["Trusted_Connection"] = "yes"
        else:
            username = self.username_edit.text().strip()
            password = self.password_edit.text()
            if not username:
                QMessageBox.warning(self, "Input Error", "Username cannot be empty for SQL Server Authentication.")
                return None
            params["UID"] = username
            params["PWD"] = password
            
        conn_str = ";".join(f"{k}={v}" for k, v in params.items())
        return conn_url + quote_plus(conn_str)

    def _test_connection(self):
        conn_url = self._get_connection_string()
        if not conn_url: return
        try:
            self.test_button.setText("Testing...")
            self.test_button.setEnabled(False)
            QApplication.processEvents()
            engine = create_engine(conn_url)
            with engine.connect() as connection:
                pass # Connection is successful if this doesn't raise an error
            QMessageBox.information(self, "Success", "Connection successful!")
        except Exception as e:
            QMessageBox.critical(self, "Connection Failed", f"Could not connect to the database.\n\nError: {e}")
        finally:
            self.test_button.setText("Test Connection")
            self.test_button.setEnabled(True)

    def _populate_from_initial_config(self, config: Optional[Dict[str, Any]], variable: Optional[str]):
        if config is None: return
        self.host_edit.setText(config.get("host", ""))
        self.db_name_edit.setText(config.get("database", ""))
        auth_mode = config.get("auth_mode", "windows")
        if auth_mode == "sql":
            self.sql_auth_radio.setChecked(True)
            self.username_edit.setText(config.get("username", ""))
        else:
            self.win_auth_radio.setChecked(True)
        if variable:
            self.assign_checkbox.setChecked(True)
            if variable in self.global_variables:
                self.existing_var_radio.setChecked(True)
                self.existing_var_combo.setCurrentText(variable)
            else:
                self.new_var_radio.setChecked(True)
                self.new_var_input.setText(variable)
        else:
            self.assign_checkbox.setChecked(False)
        self._toggle_assignment_widgets()

    def get_executor_method_name(self) -> str:
        return "_execute_mssql_connection"

    def get_config_data(self) -> Optional[Dict[str, Any]]:
        host = self.host_edit.text().strip()
        database = self.db_name_edit.text().strip()
        if not host or not database:
            QMessageBox.warning(self, "Input Error", "Host and Database Name are required.")
            return None
        config = {"host": host, "database": database}
        if self.win_auth_radio.isChecked():
            config["auth_mode"] = "windows"
        else:
            username = self.username_edit.text().strip()
            if not username:
                QMessageBox.warning(self, "Input Error", "Username is required for SQL Server Authentication.")
                return None
            config["auth_mode"] = "sql"
            config["username"] = username
            config["password"] = self.password_edit.text()
        return config

    def get_assignment_variable(self) -> Optional[str]:
        if not self.assign_checkbox.isChecked(): return None
        if self.new_var_radio.isChecked():
            var_name = self.new_var_input.text().strip()
            if not var_name:
                QMessageBox.warning(self, "Input Error", "New variable name cannot be empty.")
                return None
            return var_name
        else:
            var_name = self.existing_var_combo.currentText()
            if var_name == "-- Select Variable --":
                QMessageBox.warning(self, "Input Error", "Please select an existing variable.")
                return None
            return var_name

#
# --- [CLASS 1] The Public-Facing Module Class for Connection ---
#
class MssqlDatabase:
    """A module to create and configure a reusable MS SQL connection engine."""
    def __init__(self, context: Optional[ExecutionContext] = None):
        self.context = context

    def _log(self, message: str):
        if self.context: self.context.add_log(message)
        else: print(message)

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str],
                           initial_config: Optional[Dict[str, Any]] = None,
                           initial_variable: Optional[str] = None) -> QDialog:
        self._log("Opening MS SQL Database configuration...")
        return _MssqlConfigDialog(
            global_variables=global_variables, parent=parent_window,
            initial_config=initial_config, initial_variable=initial_variable)

    def _execute_mssql_connection(self, context: ExecutionContext, config_data: dict):
        self.context = context
        if not pyodbc:
            msg = "'pyodbc' library is not installed."
            self._log(f"FATAL ERROR: {msg}"); raise ImportError(msg)

        host, database, auth_mode = config_data.get("host"), config_data.get("database"), config_data.get("auth_mode")
        driver = ""
        drivers = [d for d in pyodbc.drivers() if "sql server" in d.lower()]
        if any("18" in d for d in drivers): driver = [d for d in drivers if "18" in d][0]
        elif any("17" in d for d in drivers): driver = [d for d in drivers if "17" in d][0]
        elif drivers: driver = drivers[-1]
        if not driver:
            msg = "No MS SQL Server ODBC driver found. Install 'ODBC Driver 17 for SQL Server' or newer."
            self._log(f"FATAL ERROR: {msg}"); raise ConnectionError(msg)
        
        conn_url = f"mssql+pyodbc:///?odbc_connect="
        params = {"DRIVER": driver, "SERVER": host, "DATABASE": database}
        if auth_mode == "windows":
            params["Trusted_Connection"] = "yes"
        elif auth_mode == "sql":
            params["UID"], params["PWD"] = config_data.get("username"), config_data.get("password")
        else:
            raise ValueError(f"Invalid authentication mode: {auth_mode}")
        conn_str = ";".join(f"{k}={v}" for k, v in params.items())
        conn_url += quote_plus(conn_str)

        try:
            self._log("Creating SQLAlchemy engine...")
            # **THIS IS THE FIX**: Do NOT set fast_executemany=True here.
            engine = create_engine(conn_url)
            with engine.connect() as connection:
                self._log("Connection successful. Engine created and ready to use.")
            return engine
        except Exception as e:
            self._log(f"FATAL ERROR during engine creation: {e}"); raise

#
# --- [CLASS 2] HELPER: The GUI Dialog for MS SQL Query ---
#
class _MssqlQueryDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None,
                 initial_variable: Optional[str] = None):
        super().__init__(parent)
        self.setWindowTitle("MS SQL Query Executor")
        self.setMinimumSize(600, 500)
        self.global_variables = global_variables
        main_layout = QVBoxLayout(self)
        conn_group = QGroupBox("Database Connection")
        conn_layout = QFormLayout(conn_group)
        self.conn_var_combo = QComboBox()
        self.conn_var_combo.addItem("-- Select Connection Variable --")
        self.conn_var_combo.addItems(self.global_variables)
        conn_layout.addRow("Connection Engine from Variable:", self.conn_var_combo)
        main_layout.addWidget(conn_group)
        sql_group = QGroupBox("SQL Statement")
        sql_layout = QVBoxLayout(sql_group)
        self.sql_from_text_radio = QRadioButton("Enter SQL Statement directly:")
        self.sql_statement_edit = QTextEdit()
        self.sql_statement_edit.setPlaceholderText("SELECT * FROM MyTable WHERE Condition = 'value'")
        self.sql_statement_edit.setFontFamily("Courier New")
        self.sql_from_var_radio = QRadioButton("Get SQL Statement from Variable:")
        self.sql_var_combo = QComboBox()
        self.sql_var_combo.addItem("-- Select Variable --")
        self.sql_var_combo.addItems(self.global_variables)
        sql_layout.addWidget(self.sql_from_text_radio)
        sql_layout.addWidget(self.sql_statement_edit)
        sql_layout.addWidget(self.sql_from_var_radio)
        sql_layout.addWidget(self.sql_var_combo)
        main_layout.addWidget(sql_group)
        assignment_group = QGroupBox("Assign Query Result to Variable")
        assign_layout = QVBoxLayout(assignment_group)
        self.assign_checkbox = QCheckBox("Assign results (DataFrame) to a variable")
        assign_layout.addWidget(self.assign_checkbox)
        self.new_var_radio = QRadioButton("New Variable Name:")
        self.new_var_input = QLineEdit("sql_results")
        self.existing_var_radio = QRadioButton("Existing Variable:")
        self.existing_var_combo = QComboBox()
        self.existing_var_combo.addItem("-- Select Variable --")
        self.existing_var_combo.addItems(self.global_variables)
        assign_form = QFormLayout()
        assign_form.addRow(self.new_var_radio, self.new_var_input)
        assign_form.addRow(self.existing_var_radio, self.existing_var_combo)
        assign_layout.addLayout(assign_form)
        main_layout.addWidget(assignment_group)
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)
        self.sql_from_text_radio.toggled.connect(self._toggle_sql_input_widgets)
        self.assign_checkbox.toggled.connect(self._toggle_assignment_widgets)
        self.new_var_radio.toggled.connect(self._toggle_assignment_widgets)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.sql_from_text_radio.setChecked(True)
        self.assign_checkbox.setChecked(True)
        self.new_var_radio.setChecked(True)
        self._toggle_sql_input_widgets()
        self._toggle_assignment_widgets()
        if initial_config: self._populate_from_initial_config(initial_config, initial_variable)

    def _toggle_sql_input_widgets(self):
        self.sql_statement_edit.setEnabled(self.sql_from_text_radio.isChecked())
        self.sql_var_combo.setEnabled(not self.sql_from_text_radio.isChecked())
    def _toggle_assignment_widgets(self):
        is_assign_enabled = self.assign_checkbox.isChecked()
        self.new_var_radio.setVisible(is_assign_enabled)
        self.new_var_input.setVisible(is_assign_enabled and self.new_var_radio.isChecked())
        self.existing_var_radio.setVisible(is_assign_enabled)
        self.existing_var_combo.setVisible(is_assign_enabled and self.existing_var_radio.isChecked())
    def _populate_from_initial_config(self, config, variable):
        self.conn_var_combo.setCurrentText(config.get("connection_var", "-- Select Connection Variable --"))
        if config.get("sql_source", "direct") == "variable":
            self.sql_from_var_radio.setChecked(True)
            self.sql_var_combo.setCurrentText(config.get("sql_statement_or_var", "-- Select Variable --"))
        else:
            self.sql_from_text_radio.setChecked(True)
            self.sql_statement_edit.setText(config.get("sql_statement_or_var", ""))
        if variable:
            self.assign_checkbox.setChecked(True)
            if variable in self.global_variables:
                self.existing_var_radio.setChecked(True)
                self.existing_var_combo.setCurrentText(variable)
            else:
                self.new_var_radio.setChecked(True)
                self.new_var_input.setText(variable)
        else: self.assign_checkbox.setChecked(False)
        self._toggle_assignment_widgets()
    def get_executor_method_name(self): return "_execute_mssql_query"
    def get_config_data(self):
        conn_var = self.conn_var_combo.currentText()
        if conn_var == "-- Select Connection Variable --":
            QMessageBox.warning(self, "Input Error", "Please select a global variable containing the database connection engine."); return None
        config = {"connection_var": conn_var}
        if self.sql_from_text_radio.isChecked():
            sql_statement = self.sql_statement_edit.toPlainText().strip()
            if not sql_statement:
                QMessageBox.warning(self, "Input Error", "The SQL statement cannot be empty."); return None
            config["sql_source"] = "direct"
            config["sql_statement_or_var"] = sql_statement
        else:
            sql_var = self.sql_var_combo.currentText()
            if sql_var == "-- Select Variable --":
                QMessageBox.warning(self, "Input Error", "Please select a variable containing the SQL statement."); return None
            config["sql_source"] = "variable"
            config["sql_statement_or_var"] = sql_var
        return config
    def get_assignment_variable(self):
        if not self.assign_checkbox.isChecked(): return None
        if self.new_var_radio.isChecked():
            var_name = self.new_var_input.text().strip()
            if not var_name:
                QMessageBox.warning(self, "Input Error", "New variable name cannot be empty."); return None
            return var_name
        else:
            var_name = self.existing_var_combo.currentText()
            if var_name == "-- Select Variable --":
                QMessageBox.warning(self, "Input Error", "Please select an existing variable."); return None
            return var_name
#
# --- [CLASS 2] The Public-Facing Module Class for Querying ---
#
class Mssql_Query:
    """A module to execute a SQL query using an existing connection engine."""
    def __init__(self, context: Optional[ExecutionContext] = None): self.context = context
    def _log(self, message: str):
        if self.context: self.context.add_log(message)
        else: print(message)

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str],
                           initial_config: Optional[Dict[str, Any]] = None,
                           initial_variable: Optional[str] = None) -> QDialog:
        self._log("Opening MS SQL Query configuration...")
        return _MssqlQueryDialog(
            global_variables=global_variables, parent=parent_window,
            initial_config=initial_config, initial_variable=initial_variable)

    def _execute_mssql_query(self, context: ExecutionContext, config_data: dict) -> pd.DataFrame:
        self.context = context
        conn_var_name = config_data["connection_var"]
        db_engine = self.context.get_variable(conn_var_name)
        if not hasattr(db_engine, 'connect'):
            raise TypeError(f"Variable '@{conn_var_name}' does not contain a valid database engine.")
        sql_statement = ""
        sql_source = config_data.get("sql_source")
        if sql_source == "direct": sql_statement = config_data.get("sql_statement_or_var")
        elif sql_source == "variable": sql_statement = self.context.get_variable(config_data.get("sql_statement_or_var"))
        if not isinstance(sql_statement, str) or not sql_statement.strip(): raise ValueError("SQL statement is empty or not a string.")
        try:
            self._log(f"Query: {sql_statement[:200]}...")
            df = pd.read_sql(sql_statement, db_engine)
            self._log(f"Query successful. Returned DataFrame with {len(df)} rows.")
            return df
        except Exception as e:
            self._log(f"FATAL ERROR during SQL query execution: {e}"); raise

#
# --- [CLASS 3] HELPER: Worker thread for fetching table schema ---
#
class _SchemaLoaderThread(QThread):
    """Worker thread to fetch database schema using a SQLAlchemy Engine."""
    schema_ready = pyqtSignal(dict)
    error_occurred = pyqtSignal(str)

    def __init__(self, engine_obj):
        super().__init__()
        self.engine_obj = engine_obj

    # --- THIS IS THE CORRECTED run METHOD ---
    def run(self):
        """Connects to the engine to get a raw DBAPI connection and cursor."""
        try:
            # Use the engine to get a connection
            with self.engine_obj.connect() as connection:
                # Get the underlying raw DBAPI connection to access the cursor
                dbapi_connection = connection.connection
                cursor = dbapi_connection.cursor()
                
                query = "SELECT TABLE_SCHEMA, TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' ORDER BY TABLE_SCHEMA, TABLE_NAME;"
                cursor.execute(query)
                tables = cursor.fetchall()
                
                schema_dict = {}
                for row in tables:
                    schema, table = row.TABLE_SCHEMA, row.TABLE_NAME
                    if schema not in schema_dict:
                        schema_dict[schema] = []
                    schema_dict[schema].append(table)
                
                self.schema_ready.emit(schema_dict)
        except Exception as e:
            self.error_occurred.emit(str(e))

#
# --- [CLASS 3] HELPER: The GUI Dialog for MS SQL Write ---
#
class _MssqlWriteDialog(QDialog):
    def __init__(self, global_variables: List[str], df_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None):
        super().__init__(parent)
        self.setWindowTitle("MS SQL Write")
        self.setMinimumSize(800, 700)
        self.global_variables = global_variables
        self.df_variables = df_variables
        self.initial_config = initial_config or {}
        self.schema_data = {}
        self.source_df_columns = []
        self.exclude_columns = []
        self.include_columns = []
        
        main_layout = QVBoxLayout(self)
        
        # Source Data Group
        source_group = QGroupBox("Source Data")
        source_form = QFormLayout(source_group)
        self.conn_var_combo = QComboBox()
        self.conn_var_combo.addItems(["-- Select Connection --"] + self.global_variables)
        self.df_var_combo = QComboBox()
        self.df_var_combo.addItems(["-- Select DataFrame --"] + self.df_variables)
        source_form.addRow("Connection Engine Variable:", self.conn_var_combo)
        source_form.addRow("DataFrame Source Variable:", self.df_var_combo)
        main_layout.addWidget(source_group)
        
        # Table Selection Group
        table_group = QGroupBox("Table to write")
        table_layout = QVBoxLayout(table_group)
        
        table_selection_layout = QHBoxLayout()
        table_selection_layout.addWidget(QLabel("Schema:"))
        self.schema_input = QLineEdit()
        self.schema_input.setPlaceholderText("dbo")
        table_selection_layout.addWidget(self.schema_input)
        
        schema_browse_button = QPushButton("📁")
        schema_browse_button.setFixedSize(30, 25)
        table_selection_layout.addWidget(schema_browse_button)
        
        table_selection_layout.addWidget(QLabel("Table:"))
        self.table_input = QLineEdit()
        table_selection_layout.addWidget(self.table_input)
        
        table_browse_button = QPushButton("📁")
        table_browse_button.setFixedSize(30, 25)
        table_selection_layout.addWidget(table_browse_button)
        
        self.select_table_button = QPushButton("Select a table")
        table_selection_layout.addWidget(self.select_table_button)
        
        table_layout.addLayout(table_selection_layout)
        
        # Options row - simplified to only include batch size and remove existing table
        options_layout = QHBoxLayout()
        options_layout.addWidget(QLabel("Batch Size:"))
        self.batch_size_spin = QSpinBox()
        self.batch_size_spin.setRange(100, 100000)
        self.batch_size_spin.setValue(1000)
        options_layout.addWidget(self.batch_size_spin)
        
        self.remove_table_check = QCheckBox("Remove existing table")
        options_layout.addWidget(self.remove_table_check)
        options_layout.addStretch()  # Push everything to the left
        
        table_layout.addLayout(options_layout)
        main_layout.addWidget(table_group)
        
        # Column Selection Group
        col_group = QGroupBox("Select the columns to write (SET in SQL)")
        col_layout = QVBoxLayout(col_group)
        
        # Selection method radio buttons
        method_layout = QHBoxLayout()
        self.manual_selection_radio = QRadioButton("Manual Selection")
        self.wildcard_selection_radio = QRadioButton("Wildcard/Regex Selection")
        self.type_selection_radio = QRadioButton("Type Selection")
        self.manual_selection_radio.setChecked(True)
        
        method_layout.addWidget(self.manual_selection_radio)
        method_layout.addWidget(self.wildcard_selection_radio)
        method_layout.addWidget(self.type_selection_radio)
        method_layout.addStretch()
        col_layout.addLayout(method_layout)
        
        # Column selection panels
        panels_layout = QHBoxLayout()
        
        # Exclude panel (red border)
        self.exclude_group = QGroupBox("Exclude")
        self.exclude_group.setStyleSheet("QGroupBox { border: 2px solid red; }")
        exclude_layout = QVBoxLayout(self.exclude_group)
        
        exclude_filter_layout = QHBoxLayout()
        exclude_filter_layout.addWidget(QLabel("🔍"))
        self.exclude_filter = QLineEdit()
        self.exclude_filter.setPlaceholderText("Filter")
        exclude_filter_layout.addWidget(self.exclude_filter)
        exclude_layout.addLayout(exclude_filter_layout)
        
        self.exclude_list = QListWidget()
        self.exclude_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)  # Allow multiple selection
        exclude_layout.addWidget(self.exclude_list)
        
        self.enforce_exclusion_radio = QRadioButton("Enforce exclusion")
        exclude_layout.addWidget(self.enforce_exclusion_radio)
        
        panels_layout.addWidget(self.exclude_group)
        
        # Arrow buttons
        arrows_layout = QVBoxLayout()
        arrows_layout.addStretch()
        
        self.move_right_button = QPushButton(">")
        self.move_right_button.setFixedSize(40, 30)
        self.move_right_button.setToolTip("Move selected columns to Include")
        arrows_layout.addWidget(self.move_right_button)
        
        self.move_all_right_button = QPushButton(">>")
        self.move_all_right_button.setFixedSize(40, 30)
        self.move_all_right_button.setToolTip("Move all columns to Include")
        arrows_layout.addWidget(self.move_all_right_button)
        
        self.move_left_button = QPushButton("<")
        self.move_left_button.setFixedSize(40, 30)
        self.move_left_button.setToolTip("Move selected columns to Exclude")
        arrows_layout.addWidget(self.move_left_button)
        
        self.move_all_left_button = QPushButton("<<")
        self.move_all_left_button.setFixedSize(40, 30)
        self.move_all_left_button.setToolTip("Move all columns to Exclude")
        arrows_layout.addWidget(self.move_all_left_button)
        
        arrows_layout.addStretch()
        panels_layout.addLayout(arrows_layout)
        
        # Include panel (green border)
        self.include_group = QGroupBox("Include")
        self.include_group.setStyleSheet("QGroupBox { border: 2px solid green; }")
        include_layout = QVBoxLayout(self.include_group)
        
        include_filter_layout = QHBoxLayout()
        include_filter_layout.addWidget(QLabel("🔍"))
        self.include_filter = QLineEdit()
        self.include_filter.setPlaceholderText("Filter")
        include_filter_layout.addWidget(self.include_filter)
        include_layout.addLayout(include_filter_layout)
        
        self.include_list = QListWidget()
        self.include_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)  # Allow multiple selection
        include_layout.addWidget(self.include_list)
        
        self.enforce_inclusion_radio = QRadioButton("Enforce inclusion")
        self.enforce_inclusion_radio.setChecked(True)
        include_layout.addWidget(self.enforce_inclusion_radio)
        
        panels_layout.addWidget(self.include_group)
        
        col_layout.addLayout(panels_layout)
        main_layout.addWidget(col_group)
        
        # Buttons
        button_layout = QHBoxLayout()
        self.ok_button = QPushButton("OK")
        #self.apply_button = QPushButton("Apply")
        self.cancel_button = QPushButton("Cancel")
        self.help_button = QPushButton("❓")
        self.help_button.setFixedSize(30, 30)
        
        button_layout.addStretch()
        button_layout.addWidget(self.ok_button)
        #button_layout.addWidget(self.apply_button)
        button_layout.addWidget(self.cancel_button)
        button_layout.addWidget(self.help_button)
        main_layout.addLayout(button_layout)
        
        # Connect signals
        self.df_var_combo.currentTextChanged.connect(self._on_dataframe_changed)
        self.select_table_button.clicked.connect(self._show_table_selector)
        self.exclude_filter.textChanged.connect(self._filter_exclude_list)
        self.include_filter.textChanged.connect(self._filter_include_list)
        
        # Arrow button connections
        self.move_right_button.clicked.connect(self._move_selected_to_include)
        self.move_all_right_button.clicked.connect(self._move_all_to_include)
        self.move_left_button.clicked.connect(self._move_selected_to_exclude)
        self.move_all_left_button.clicked.connect(self._move_all_to_exclude)
        
        self.ok_button.clicked.connect(self.accept)
        #self.apply_button.clicked.connect(self._apply_changes)
        self.cancel_button.clicked.connect(self.reject)
        
        # Initialize
        self._populate_from_initial_config()
        
    def _on_dataframe_changed(self, df_var_name: str):
        """When DataFrame changes, load all columns into the Exclude list"""
        if df_var_name == "-- Select DataFrame --":
            self.source_df_columns = []
            self.exclude_columns = []
            self.include_columns = []
        else:
            df = self._get_context_variable(df_var_name)
            if isinstance(df, pd.DataFrame):
                self.source_df_columns = list(df.columns)
                # IMPORTANT: Initialize ALL columns in EXCLUDE list (user will move desired ones to include)
                self.exclude_columns = self.source_df_columns.copy()
                self.include_columns = []
                self._log(f"Loaded {len(self.source_df_columns)} columns from DataFrame '{df_var_name}' into Exclude list")
            else:
                self.source_df_columns = []
                self.exclude_columns = []
                self.include_columns = []
                if df_var_name != "-- Select DataFrame --":
                    QMessageBox.warning(self, "Invalid DataFrame", 
                                      f"Variable '{df_var_name}' does not contain a valid pandas DataFrame.")
        
        # Clear any existing filters and refresh the lists
        self.exclude_filter.clear()
        self.include_filter.clear()
        self._refresh_column_lists()
        
        # Log the action for user feedback
        if self.exclude_columns:
            self._log(f"All {len(self.exclude_columns)} DataFrame columns loaded into Exclude list. Move desired columns to Include list using arrow buttons.")

    def _log(self, message: str):
        """Log message if context is available"""
        if hasattr(self.parent(), 'add_log'):
            self.parent().add_log(message)
        else:
            print(message)  # Fallback for testing
        
    def _get_context_variable(self, var_name: str):
        if hasattr(self.parent(), 'get_variable_for_dialog'):
            return self.parent().get_variable_for_dialog(var_name)
        return None
        
    def _refresh_column_lists(self):
        """Refresh both exclude and include list widgets based on current column assignments and filters"""
        # Clear the list widgets
        self.exclude_list.clear()
        self.include_list.clear()
        
        # Get current filter text (case-insensitive)
        exclude_filter_text = self.exclude_filter.text().lower()
        include_filter_text = self.include_filter.text().lower()
        
        # Populate exclude list with filtered columns
        for col in self.exclude_columns:
            if exclude_filter_text in col.lower():
                self.exclude_list.addItem(col)
        
        # Populate include list with filtered columns
        for col in self.include_columns:
            if include_filter_text in col.lower():
                self.include_list.addItem(col)
        
        # Update the group box titles to show counts
        self._update_group_titles()

    def _update_group_titles(self):
        """Update group box titles to show column counts"""
        exclude_count = len([item for item in self.exclude_columns if self.exclude_filter.text().lower() in item.lower()])
        include_count = len([item for item in self.include_columns if self.include_filter.text().lower() in item.lower()])
        
        # Update group box titles
        self.exclude_group.setTitle(f"Exclude ({exclude_count})")
        self.include_group.setTitle(f"Include ({include_count})")
        
    def _filter_exclude_list(self):
        self._refresh_column_lists()
        
    def _filter_include_list(self):
        self._refresh_column_lists()
        
    def _move_selected_to_include(self):
        selected_items = self.exclude_list.selectedItems()
        for item in selected_items:
            col_name = item.text()
            if col_name in self.exclude_columns:
                self.exclude_columns.remove(col_name)
                self.include_columns.append(col_name)
        self._refresh_column_lists()
        
    def _move_all_to_include(self):
        # Move all currently visible (filtered) exclude columns to include
        visible_exclude_cols = [self.exclude_list.item(i).text() for i in range(self.exclude_list.count())]
        for col in visible_exclude_cols:
            if col in self.exclude_columns:
                self.exclude_columns.remove(col)
                self.include_columns.append(col)
        self._refresh_column_lists()
        
    def _move_selected_to_exclude(self):
        selected_items = self.include_list.selectedItems()
        for item in selected_items:
            col_name = item.text()
            if col_name in self.include_columns:
                self.include_columns.remove(col_name)
                self.exclude_columns.append(col_name)
        self._refresh_column_lists()
        
    def _move_all_to_exclude(self):
        # Move all currently visible (filtered) include columns to exclude
        visible_include_cols = [self.include_list.item(i).text() for i in range(self.include_list.count())]
        for col in visible_include_cols:
            if col in self.include_columns:
                self.include_columns.remove(col)
                self.exclude_columns.append(col)
        self._refresh_column_lists()
        
    def _show_table_selector(self):
        """Shows a dialog with database schema and tables in a tree structure"""
        engine = self._get_context_variable(self.conn_var_combo.currentText())
        if not engine:
            QMessageBox.warning(self, "No Connection", "Please select a connection first.")
            return
            
        # Create table selector dialog
        table_dialog = QDialog(self)
        table_dialog.setWindowTitle("Select Table")
        table_dialog.setMinimumSize(400, 500)
        
        layout = QVBoxLayout(table_dialog)
        
        # Loading label
        loading_label = QLabel("Loading tables...")
        layout.addWidget(loading_label)
        
        # Tree widget for schema/tables
        tree_widget = QTreeWidget()
        tree_widget.setHeaderLabels(["Schemas & Tables"])
        tree_widget.hide()  # Hide initially while loading
        layout.addWidget(tree_widget)
        
        # Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        layout.addWidget(button_box)
        
        # Load schema in background
        def load_schema_async():
            try:
                with engine.connect() as connection:
                    dbapi_connection = connection.connection
                    cursor = dbapi_connection.cursor()
                    
                    query = "SELECT TABLE_SCHEMA, TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' ORDER BY TABLE_SCHEMA, TABLE_NAME;"
                    cursor.execute(query)
                    tables = cursor.fetchall()
                    
                    schema_dict = {}
                    for row in tables:
                        schema, table = row.TABLE_SCHEMA, row.TABLE_NAME
                        if schema not in schema_dict:
                            schema_dict[schema] = []
                        schema_dict[schema].append(table)
                    
                    # Populate tree widget
                    tree_widget.clear()
                    for schema, tables in sorted(schema_dict.items()):
                        schema_item = QTreeWidgetItem(tree_widget, [schema])
                        for table in sorted(tables):
                            table_item = QTreeWidgetItem(schema_item, [table])
                    
                    # Show tree and hide loading label
                    loading_label.hide()
                    tree_widget.show()
                    tree_widget.expandAll()
                    
            except Exception as e:
                loading_label.setText(f"Error loading tables: {str(e)}")
        
        # Load schema asynchronously
        QTimer.singleShot(100, load_schema_async)
        
        button_box.accepted.connect(table_dialog.accept)
        button_box.rejected.connect(table_dialog.reject)
        
        if table_dialog.exec() == QDialog.DialogCode.Accepted:
            selected_item = tree_widget.currentItem()
            if selected_item and selected_item.parent():
                # User selected a table (not schema)
                schema_name = selected_item.parent().text(0)
                table_name = selected_item.text(0)
                
                self.schema_input.setText(schema_name)
                self.table_input.setText(table_name)
                
    def _apply_changes(self):
        # Apply current settings without closing dialog
        QMessageBox.information(self, "Applied", "Settings applied successfully.")
        
    def _populate_from_initial_config(self):
        if not self.initial_config:
            return
            
        # Populate basic fields
        self.conn_var_combo.setCurrentText(self.initial_config.get("connection_var", "-- Select Connection --"))
        self.df_var_combo.setCurrentText(self.initial_config.get("dataframe_var", "-- Select DataFrame --"))
        self.schema_input.setText(self.initial_config.get("schema", "dbo"))
        self.table_input.setText(self.initial_config.get("table", ""))
        
        # Populate options
        self.batch_size_spin.setValue(self.initial_config.get("batch_size", 1000))
        self.remove_table_check.setChecked(self.initial_config.get("remove_table", False))
        
        # Load dataframe columns if available
        self._on_dataframe_changed(self.df_var_combo.currentText())
        
        # Restore column selections if available
        if "include_columns" in self.initial_config:
            saved_include = self.initial_config["include_columns"]
            saved_exclude = self.initial_config.get("exclude_columns", [])
            
            self.include_columns = [col for col in saved_include if col in self.source_df_columns]
            self.exclude_columns = [col for col in self.source_df_columns if col not in self.include_columns]
            self._refresh_column_lists()
        
    def get_executor_method_name(self):
        return "_execute_mssql_write"
        
    def get_assignment_variable(self):
        return None
        
    def get_config_data(self):
        config = {
            "connection_var": self.conn_var_combo.currentText(),
            "dataframe_var": self.df_var_combo.currentText(),
            "schema": self.schema_input.text().strip() or "dbo",
            "table": self.table_input.text().strip(),
            "batch_size": self.batch_size_spin.value(),
            "remove_table": self.remove_table_check.isChecked(),
            "include_columns": self.include_columns.copy(),
            "exclude_columns": self.exclude_columns.copy(),
            "enforce_inclusion": self.enforce_inclusion_radio.isChecked(),
            "selection_method": "manual" if self.manual_selection_radio.isChecked() else 
                             "wildcard" if self.wildcard_selection_radio.isChecked() else "type"
        }
        
        # Validation
        if config["connection_var"] == "-- Select Connection --":
            QMessageBox.warning(self, "Input Error", "Please select a connection.")
            return None
        if config["dataframe_var"] == "-- Select DataFrame --":
            QMessageBox.warning(self, "Input Error", "Please select a DataFrame.")
            return None
        if not config["table"]:
            QMessageBox.warning(self, "Input Error", "Please enter a table name.")
            return None
        if not config["include_columns"]:
            QMessageBox.warning(self, "Input Error", "Please select at least one column to write.")
            return None
            
        return config

#
# --- [CLASS 3] The Public-Facing Module Class for Writing ---
#
class Mssql_Write:
    """A module to write a pandas DataFrame to an MS SQL database."""
    def __init__(self, context: Optional[ExecutionContext] = None): 
        self.context = context
        
    def _log(self, message: str):
        if self.context: 
            self.context.add_log(message)
        else: 
            print(message)

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str],
                           initial_config: Optional[Dict[str, Any]] = None, **kwargs) -> QDialog:
        self._log("Opening MS SQL Write configuration...")
        df_variables = []
        if hasattr(parent_window, 'get_dataframe_variables'): 
            df_variables = parent_window.get_dataframe_variables()
        else: 
            df_variables = global_variables
        return _MssqlWriteDialog(global_variables, df_variables, parent_window, initial_config)

    def _execute_mssql_write(self, context: ExecutionContext, config_data: dict):
        self.context = context
        db_engine = context.get_variable(config_data["connection_var"])
        df_to_write = context.get_variable(config_data["dataframe_var"])
        
        if not hasattr(db_engine, 'connect'): 
            raise TypeError(f"Variable '@{config_data['connection_var']}' is not a valid database engine.")
        if not isinstance(df_to_write, pd.DataFrame): 
            raise TypeError(f"Variable '@{config_data['dataframe_var']}' is not a pandas DataFrame.")
        
        table_name = config_data["table"]
        schema = config_data.get("schema", "dbo")
        remove_table = config_data.get("remove_table", False)
        include_columns = config_data.get("include_columns", [])
        batch_size = config_data.get("batch_size", 1000)
        
        # Filter dataframe to only include selected columns
        if include_columns:
            # Ensure all specified columns exist in the dataframe
            missing_cols = [col for col in include_columns if col not in df_to_write.columns]
            if missing_cols:
                raise ValueError(f"Columns not found in DataFrame: {missing_cols}")
            
            # Filter the dataframe
            df_filtered = df_to_write[include_columns].copy()
            self._log(f"Filtered DataFrame from {len(df_to_write.columns)} to {len(df_filtered.columns)} columns: {include_columns}")
        else:
            df_filtered = df_to_write.copy()
        
        self._log(f"Preparing to write {len(df_filtered)} rows to [{schema}].[{table_name}].")
        
        try:
            if remove_table:
                # Drop and recreate table
                self._log("'Remove existing table' is checked. Dropping table if it exists...")
                with db_engine.begin() as conn:
                    try:
                        conn.execute(text(f"DROP TABLE [{schema}].[{table_name}]"))
                        self._log(f"Table [{schema}].[{table_name}] dropped successfully.")
                    except Exception:
                        self._log(f"Table [{schema}].[{table_name}] does not exist or could not be dropped. Continuing...")
                
                # Create new table and insert data
                df_filtered.to_sql(
                    name=table_name, 
                    con=db_engine, 
                    schema=schema, 
                    if_exists='replace', 
                    index=False,
                    chunksize=batch_size
                )
                self._log(f"New table created and {len(df_filtered)} rows inserted.")
            else:
                # Append to existing table
                df_filtered.to_sql(
                    name=table_name, 
                    con=db_engine, 
                    schema=schema, 
                    if_exists='append', 
                    index=False,
                    chunksize=batch_size
                )
                self._log(f"Successfully appended {len(df_filtered)} rows to existing table.")
                
        except Exception as e:
            self._log(f"FATAL ERROR during SQL write: {e}")
            raise

#
# --- [CLASS 4] HELPER: The GUI Dialog for MS SQL Merge ---
#
class _MssqlMergeDialog(QDialog):
    """A custom GUI to configure a MERGE (UPSERT) operation."""
    def __init__(self, global_variables: List[str], df_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None):
        super().__init__(parent)
        self.setWindowTitle("MS SQL Merge (Update/Insert)")
        self.setMinimumSize(700, 800) # Increased height slightly for the new group
        self.global_variables = global_variables
        self.df_variables = df_variables
        self.initial_config = initial_config or {}
        self.source_df_columns, self.target_table_columns, self.join_condition_widgets, self.update_column_widgets = [], [], [], []

        # --- UI Layout ---
        main_layout = QVBoxLayout(self)
        
        # 1. Source Data
        source_group = QGroupBox("1. Source Data")
        source_form = QFormLayout(source_group)
        self.conn_var_combo = QComboBox(); self.conn_var_combo.addItems(["-- Select Connection --"] + self.global_variables)
        self.df_var_combo = QComboBox(); self.df_var_combo.addItems(["-- Select DataFrame --"] + self.df_variables)
        source_form.addRow("Connection Engine:", self.conn_var_combo); source_form.addRow("Source DataFrame:", self.df_var_combo)
        main_layout.addWidget(source_group)
        
        # 2. Target Table Browser
        target_group = QGroupBox("2. Target Table Browser")
        target_layout = QVBoxLayout(target_group)
        self.load_schema_button = QPushButton("Load Database Schema")
        target_layout.addWidget(self.load_schema_button)
        self.table_tree = QTreeWidget()
        self.table_tree.setHeaderLabels(["Schemas & Tables"])
        self.table_tree.header().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        target_layout.addWidget(self.table_tree)
        main_layout.addWidget(target_group)

        # 2a. Target Table Selected (NEW)
        selected_table_group = QGroupBox("2a. Target Table Selected")
        selected_table_layout = QFormLayout(selected_table_group)
        self.selected_schema_text = QLineEdit()
        self.selected_schema_text.setReadOnly(True)
        self.selected_schema_text.setPlaceholderText("Schema will appear here...")
        self.selected_table_text = QLineEdit()
        self.selected_table_text.setReadOnly(True)
        self.selected_table_text.setPlaceholderText("Table will appear here...")
        selected_table_layout.addRow("Schema:", self.selected_schema_text)
        selected_table_layout.addRow("Table:", self.selected_table_text)
        main_layout.addWidget(selected_table_group)

        # 3. Merge Configuration
        config_group = QGroupBox("3. Merge Configuration")
        config_layout = QVBoxLayout(config_group)
        join_group = QGroupBox("ON (Join Condition)")
        self.join_layout = QVBoxLayout(join_group)
        self.add_join_condition_button = QPushButton("➕ Add Condition")
        self.join_layout.addWidget(self.add_join_condition_button)
        config_layout.addWidget(join_group)
        update_group = QGroupBox("WHEN MATCHED THEN UPDATE")
        self.update_layout = QVBoxLayout(update_group)
        self.add_update_column_button = QPushButton("➕ Add Column to Update")
        self.update_layout.addWidget(self.add_update_column_button)
        config_layout.addWidget(update_group)
        insert_group = QGroupBox("WHEN NOT MATCHED BY TARGET THEN INSERT")
        insert_layout = QVBoxLayout(insert_group)
        insert_layout.addWidget(QLabel("All columns from the source DataFrame will be inserted."))
        config_layout.addWidget(insert_group)
        main_layout.addWidget(config_group)

        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # --- Connections ---
        self.load_schema_button.clicked.connect(self._load_schema)
        self.df_var_combo.currentTextChanged.connect(self._on_source_df_changed)
        self.table_tree.currentItemChanged.connect(self._on_target_table_changed)
        self.add_join_condition_button.clicked.connect(lambda: self._add_row_to_layout(self.join_layout, self.join_condition_widgets, self._create_join_condition_row))
        self.add_update_column_button.clicked.connect(lambda: self._add_row_to_layout(self.update_layout, self.update_column_widgets, self._create_update_column_row))
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        self._populate_static_fields()

    def showEvent(self, event):
        super().showEvent(event)
        if not event.spontaneous():
            QTimer.singleShot(50, self._initial_load_sequence)

    def _initial_load_sequence(self):
        self._on_source_df_changed(self.df_var_combo.currentText())
        if self.initial_config and self.initial_config.get("target_table"):
            self._load_schema()

    def _populate_static_fields(self):
        if not self.initial_config: return
        self.conn_var_combo.setCurrentText(self.initial_config.get("connection_var", "-- Select Connection --"))
        self.df_var_combo.setCurrentText(self.initial_config.get("dataframe_var", "-- Select DataFrame --"))

    def _on_source_df_changed(self, df_var_name: str):
        if df_var_name == "-- Select DataFrame --":
            self.source_df_columns = []
        else:
            df = self._get_context_variable(df_var_name)
            self.source_df_columns = sorted(list(df.columns)) if isinstance(df, pd.DataFrame) else []
        self._rebuild_dynamic_combos()
        self._repopulate_dynamic_rows_if_ready()

    def _on_target_table_changed(self, current: QTreeWidgetItem, previous: Optional[QTreeWidgetItem]):
        if not (current and current.parent()):
            self.target_table_columns = []
            self.selected_schema_text.clear()
            self.selected_table_text.clear()
            self._rebuild_dynamic_combos()
            return
        
        schema, table = current.parent().text(0), current.text(0)
        self.selected_schema_text.setText(schema)
        self.selected_table_text.setText(table)
        
        engine_obj = self._get_context_variable(self.conn_var_combo.currentText())
        if not engine_obj:
            self.target_table_columns = []
            self._rebuild_dynamic_combos()
            return
            
        try:
            with engine_obj.connect() as connection:
                cursor = connection.connection.cursor()
                self.target_table_columns = sorted([row.column_name for row in cursor.columns(schema=schema, table=table)])
        except Exception as e:
            QMessageBox.warning(self, "Column Load Error", f"Could not fetch columns for {schema}.{table}:\n{e}")
            self.target_table_columns = []
        self._rebuild_dynamic_combos()
        self._repopulate_dynamic_rows_if_ready()

    def _repopulate_dynamic_rows_if_ready(self):
        if not self.initial_config or not self.source_df_columns or not self.target_table_columns:
            return

        for widget, _ in self.join_condition_widgets: widget.deleteLater()
        self.join_condition_widgets.clear()
        for widget, _ in self.update_column_widgets: widget.deleteLater()
        self.update_column_widgets.clear()

        for condition in self.initial_config.get("on_conditions", []):
            self._add_row_to_layout(self.join_layout, self.join_condition_widgets, self._create_join_condition_row)
            _, combos = self.join_condition_widgets[-1]
            combos['source'].setCurrentText(condition.get('source', '-- Source --'))
            combos['target'].setCurrentText(condition.get('target', '-- Target --'))

        for update_col in self.initial_config.get("update_columns", []):
            self._add_row_to_layout(self.update_layout, self.update_column_widgets, self._create_update_column_row)
            _, combos = self.update_column_widgets[-1]
            combos['target'].setCurrentText(update_col.get('target', '-- Target --'))
            combos['source'].setCurrentText(update_col.get('source', '-- Source --'))
        
        self.initial_config = None

    def _get_context_variable(self, var_name: str):
        if not hasattr(self.parent(), 'get_variable_for_dialog'):
            QMessageBox.critical(self, "Error", "Main application support function 'get_variable_for_dialog' is missing."); return None
        return self.parent().get_variable_for_dialog(var_name)

    def _load_schema(self):
        engine_obj = self._get_context_variable(self.conn_var_combo.currentText())
        if not engine_obj:
            QMessageBox.warning(self, "Connection Missing", "Please select a valid connection engine first.")
            return

        self.load_schema_button.setText("Loading...")
        self.load_schema_button.setEnabled(False)
        self.table_tree.clear()

        self.schema_loader_thread = _SchemaLoaderThread(engine_obj)
        self.schema_loader_thread.schema_ready.connect(self._on_schema_loaded)
        self.schema_loader_thread.error_occurred.connect(lambda e: (
            QMessageBox.critical(self, "Schema Load Failed", str(e)),
            self.load_schema_button.setText("Reload Schema"),
            self.load_schema_button.setEnabled(True)
        ))
        self.schema_loader_thread.start()

    def _on_schema_loaded(self, schema_dict: dict):
        self.table_tree.clear()
        for schema, tables in sorted(schema_dict.items()):
            schema_item = QTreeWidgetItem(self.table_tree, [schema])
            for table in sorted(tables):
                QTreeWidgetItem(schema_item, [table])
        self.load_schema_button.setText("Reload Schema")
        self.load_schema_button.setEnabled(True)
        self._restore_tree_selection()

    def _restore_tree_selection(self):
        config = self.initial_config if self.initial_config else {}
        saved_schema, saved_table = config.get("target_schema"), config.get("target_table")
        if not saved_schema or not saved_table: return

        iterator = QTreeWidgetItemIterator(self.table_tree)
        while iterator.value():
            item = iterator.value()
            if item.parent() and item.parent().text(0) == saved_schema and item.text(0) == saved_table:
                self.table_tree.setCurrentItem(item)
                self.table_tree.scrollToItem(item)
                self.table_tree.expandItem(item.parent())
                break
            iterator += 1

    def _create_join_condition_row(self):
        widget = QWidget(); layout = QHBoxLayout(widget); layout.setContentsMargins(0, 0, 0, 0)
        source_combo = QComboBox(); source_combo.addItems(["-- Source --"] + self.source_df_columns)
        target_combo = QComboBox(); target_combo.addItems(["-- Target --"] + self.target_table_columns)
        remove_button = QPushButton("➖"); remove_button.setFixedWidth(30)
        remove_button.clicked.connect(lambda: self._remove_row_from_layout(widget, self.join_layout, self.join_condition_widgets))
        layout.addWidget(QLabel("Source:")); layout.addWidget(source_combo)
        layout.addWidget(QLabel(" = Target:")); layout.addWidget(target_combo)
        layout.addWidget(remove_button)
        return widget, {'source': source_combo, 'target': target_combo}

    def _create_update_column_row(self):
        widget = QWidget(); layout = QHBoxLayout(widget); layout.setContentsMargins(0, 0, 0, 0)
        target_combo = QComboBox(); target_combo.addItems(["-- Target --"] + self.target_table_columns)
        source_combo = QComboBox(); source_combo.addItems(["-- Source --"] + self.source_df_columns)
        remove_button = QPushButton("➖"); remove_button.setFixedWidth(30)
        remove_button.clicked.connect(lambda: self._remove_row_from_layout(widget, self.update_layout, self.update_column_widgets))
        layout.addWidget(QLabel("SET Target:")); layout.addWidget(target_combo)
        layout.addWidget(QLabel(" = Source:")); layout.addWidget(source_combo)
        layout.addWidget(remove_button)
        return widget, {'target': target_combo, 'source': source_combo}

    def _add_row_to_layout(self, layout, widget_list, create_row_func):
        row_widget, combo_dict = create_row_func()
        layout.addWidget(row_widget)
        widget_list.append((row_widget, combo_dict))

    def _remove_row_from_layout(self, widget_to_remove, layout, widget_list):
        for item in widget_list:
            if item[0] is widget_to_remove:
                widget_list.remove(item)
                break
        widget_to_remove.deleteLater()

    def _rebuild_dynamic_combos(self):
        for _, combos in self.join_condition_widgets:
            s_val = combos['source'].currentText(); t_val = combos['target'].currentText()
            combos['source'].clear(); combos['source'].addItems(["-- Source --"] + self.source_df_columns)
            combos['target'].clear(); combos['target'].addItems(["-- Target --"] + self.target_table_columns)
            combos['source'].setCurrentText(s_val); combos['target'].setCurrentText(t_val)
        for _, combos in self.update_column_widgets:
            t_val = combos['target'].currentText(); s_val = combos['source'].currentText()
            combos['target'].clear(); combos['target'].addItems(["-- Target --"] + self.target_table_columns)
            combos['source'].clear(); combos['source'].addItems(["-- Source --"] + self.source_df_columns)
            combos['target'].setCurrentText(t_val); combos['source'].setCurrentText(s_val)

    def get_executor_method_name(self): return "_execute_mssql_merge"
    def get_assignment_variable(self): return None

    def get_config_data(self):
        config = {"connection_var": self.conn_var_combo.currentText(), "dataframe_var": self.df_var_combo.currentText()}
        if config["connection_var"] == "-- Select Connection --" or config["dataframe_var"] == "-- Select DataFrame --":
            QMessageBox.warning(self, "Input Error", "Please select a Connection and a Source DataFrame."); return None
        
        target_schema = self.selected_schema_text.text()
        target_table = self.selected_table_text.text()
        if not target_schema or not target_table:
            QMessageBox.warning(self, "Input Error", "Please select a target table from the schema browser."); return None
        config["target_schema"], config["target_table"] = target_schema, target_table
        
        config["on_conditions"] = [{'source': c['source'].currentText(), 'target': c['target'].currentText()} for _, c in self.join_condition_widgets]
        if not config["on_conditions"] or any(c['source'] == "-- Source --" or c['target'] == "-- Target --" for c in config["on_conditions"]):
            QMessageBox.warning(self, "Input Error", "All 'ON' conditions must be fully specified."); return None
        config["update_columns"] = [{'target': c['target'].currentText(), 'source': c['source'].currentText()} for _, c in self.update_column_widgets]
        if not config["update_columns"] or any(c['target'] == "-- Target --" or c['source'] == "-- Source --" for c in config["update_columns"]):
            QMessageBox.warning(self, "Input Error", "All 'UPDATE' rules must be fully specified."); return None
        return config

#
# --- [CLASS 4] The Public-Facing Module Class for Merging ---
    # --- THIS IS THE CORRECTED EXECUTOR METHOD ---
class Mssql_Merge:
    """A module to MERGE (update/insert) a DataFrame into an MS SQL table."""
    def __init__(self, context: Optional[ExecutionContext] = None): self.context = context
    def _log(self, message: str):
        if self.context: self.context.add_log(message)
        else: print(message)

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str],
                           initial_config: Optional[Dict[str, Any]] = None, **kwargs) -> QDialog:
        self._log("Opening MS SQL Merge configuration...")
        df_variables = []
        if hasattr(parent_window, 'get_dataframe_variables'): df_variables = parent_window.get_dataframe_variables()
        else: df_variables = global_variables
        return _MssqlMergeDialog(global_variables, df_variables, parent_window, initial_config)

    # --- THIS IS THE NEW, MORE ROBUST EXECUTOR METHOD ---
    # --- THIS IS THE CORRECTED EXECUTOR METHOD ---

# --- In CLASS Mssql_Merge ---

    def _execute_mssql_merge(self, context: ExecutionContext, config_data: dict):
        self.context = context
        db_engine = context.get_variable(config_data["connection_var"])
        source_df = context.get_variable(config_data["dataframe_var"])
        target_schema = config_data["target_schema"]
        target_table = config_data["target_table"]

        if not hasattr(db_engine, 'connect'):
            raise TypeError(f"Variable '@{config_data['connection_var']}' is not a valid database engine.")
        if not isinstance(source_df, pd.DataFrame) or source_df.empty:
            raise TypeError(f"Variable '@{config_data['dataframe_var']}' is not a non-empty pandas DataFrame.")

        # Best practice: create a copy to avoid SettingWithCopyWarning
        source_df = source_df.copy()
        
        # Ensure all object columns are safe for SQL transfer
        for col in source_df.select_dtypes(include=['object']).columns:
            source_df[col] = source_df[col].astype(str).replace('<NA>', None)

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        # Use a permanent, uniquely named staging table that we will explicitly drop
        staging_table_name = f"ZZ_STAGE_MERGE_{timestamp}"
        
        self._log(f"Preparing to merge {len(source_df)} rows into [{target_schema}].[{target_table}].")
        self._log(f"Staging table will be: [{target_schema}].[{staging_table_name}]")

        # The 'fast_executemany' engine is for the initial high-speed data dump.
        fast_engine = create_engine(db_engine.url, fast_executemany=True)

        try:
            # --- PHASE 1: Create and Populate Staging Table ---
            # This uses the high-performance engine.
            with fast_engine.connect() as conn:
                self._log(f"Writing {len(source_df)} rows to staging table [{target_schema}].[{staging_table_name}]...")
                source_df.to_sql(
                    name=staging_table_name,
                    con=conn,
                    schema=target_schema,
                    if_exists='replace', # Creates the table and inserts data
                    index=False,
                    chunksize=10000
                )
            self._log("Staging table populated successfully.")

            # --- PHASE 2: Execute MERGE using the main engine ---
            # This ensures transactionality and avoids issues with fast_executemany state.
            with db_engine.connect() as conn:
                on_conditions = " AND ".join([f"T.[{c['target']}] = S.[{c['source']}]" for c in config_data["on_conditions"]])
                update_clauses = ", ".join([f"T.[{c['target']}] = S.[{c['source']}]" for c in config_data["update_columns"]])
                insert_cols = ", ".join([f"[{col}]" for col in source_df.columns])
                source_cols_for_values = ", ".join([f"S.[{col}]" for col in source_df.columns])

                merge_sql = text(f"""
                MERGE [{target_schema}].[{target_table}] AS T
                USING [{target_schema}].[{staging_table_name}] AS S ON ({on_conditions})
                WHEN MATCHED THEN
                    UPDATE SET {update_clauses}
                WHEN NOT MATCHED BY TARGET THEN
                    INSERT ({insert_cols}) VALUES ({source_cols_for_values});
                """)

                self._log("Executing MERGE statement in a transaction...")
                trans = conn.begin()
                try:
                    result = conn.execute(merge_sql)
                    trans.commit()
                    self._log(f"MERGE operation successful. {result.rowcount} rows affected.")
                except Exception as merge_err:
                    self._log(f"ERROR during MERGE. Rolling back transaction. Error: {merge_err}")
                    trans.rollback()
                    raise merge_err # Re-raise the error to be caught by the outer block

        except Exception as e:
            # If any part of the process fails, log it and re-raise the exception
            self._log(f"FATAL ERROR during merge process: {e}")
            raise

        finally:
            # --- PHASE 3: GUARANTEED CLEANUP ---
            # This block executes whether an exception occurred or not.
            self._log("Initiating cleanup phase...")
            try:
                # Use the original engine to drop the table
                with db_engine.connect() as conn:
                    with conn.begin(): # Use a transaction for the drop
                        conn.execute(text(f"DROP TABLE [{target_schema}].[{staging_table_name}]"))
                    self._log(f"Successfully dropped staging table [{target_schema}].[{staging_table_name}].")
            except Exception as drop_err:
                # If dropping fails, log a clear warning for manual cleanup
                self._log(f"CRITICAL WARNING: Failed to drop staging table [{target_schema}].[{staging_table_name}].")
                self._log(f"Please drop it manually. Error: {drop_err}")
            
            # Dispose of the temporary high-speed engine
            if fast_engine:
                fast_engine.dispose()
            self._log("Cleanup complete. The main engine remains open.")

#
# --- [CLASS 5] HELPER: The GUI Dialog for Closing a Connection ---
#

#
# --- [CLASS 6] HELPER: The GUI Dialog for MS SQL Execute Query ---
#
class _MssqlExecuteQueryDialog(QDialog):
    """A dialog to configure a non-data-returning SQL query execution."""
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None):
        super().__init__(parent)
        self.setWindowTitle("MS SQL Execute Statement")
        self.setMinimumSize(600, 500)
        self.global_variables = global_variables
        
        main_layout = QVBoxLayout(self)

        # 1. Connection Selection
        conn_group = QGroupBox("Database Connection")
        conn_layout = QFormLayout(conn_group)
        self.conn_var_combo = QComboBox()
        self.conn_var_combo.addItem("-- Select Connection Variable --")
        self.conn_var_combo.addItems(self.global_variables)
        conn_layout.addRow("Connection Engine from Variable:", self.conn_var_combo)
        main_layout.addWidget(conn_group)

        # 2. SQL Statement Source
        sql_group = QGroupBox("SQL Statement")
        sql_layout = QVBoxLayout(sql_group)
        self.sql_from_text_radio = QRadioButton("Enter SQL Statement directly:")
        self.sql_statement_edit = QTextEdit()
        self.sql_statement_edit.setPlaceholderText("e.g., UPDATE MyTable SET Status = 'Processed' WHERE ID = 1")
        self.sql_statement_edit.setFontFamily("Courier New")
        
        self.sql_from_var_radio = QRadioButton("Get SQL Statement from Variable:")
        self.sql_var_combo = QComboBox()
        self.sql_var_combo.addItem("-- Select Variable --")
        self.sql_var_combo.addItems(self.global_variables)

        sql_layout.addWidget(self.sql_from_text_radio)
        sql_layout.addWidget(self.sql_statement_edit)
        sql_layout.addWidget(self.sql_from_var_radio)
        sql_layout.addWidget(self.sql_var_combo)
        main_layout.addWidget(sql_group)

        # 3. Buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # --- Connections & Initial State ---
        self.sql_from_text_radio.toggled.connect(self._toggle_sql_input_widgets)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.sql_from_text_radio.setChecked(True)
        self._toggle_sql_input_widgets()
        
        if initial_config:
            self._populate_from_initial_config(initial_config)

    def _toggle_sql_input_widgets(self):
        """Enable/disable input widgets based on radio button selection."""
        is_direct_input = self.sql_from_text_radio.isChecked()
        self.sql_statement_edit.setEnabled(is_direct_input)
        self.sql_var_combo.setEnabled(not is_direct_input)

    def _populate_from_initial_config(self, config: Dict[str, Any]):
        """Load settings from a previously saved configuration."""
        self.conn_var_combo.setCurrentText(config.get("connection_var", "-- Select Connection Variable --"))
        
        if config.get("sql_source", "direct") == "variable":
            self.sql_from_var_radio.setChecked(True)
            self.sql_var_combo.setCurrentText(config.get("sql_statement_or_var", "-- Select Variable --"))
        else:
            self.sql_from_text_radio.setChecked(True)
            self.sql_statement_edit.setText(config.get("sql_statement_or_var", ""))
        self._toggle_sql_input_widgets()

    def get_executor_method_name(self) -> str:
        """Specifies which method in the main class will do the work."""
        return "_execute_mssql_execute"

    def get_assignment_variable(self) -> Optional[str]:
        """This action does not return a value, so it cannot be assigned."""
        return None

    def get_config_data(self) -> Optional[Dict[str, Any]]:
        """Validate and return the configuration from the dialog."""
        conn_var = self.conn_var_combo.currentText()
        if conn_var == "-- Select Connection Variable --":
            QMessageBox.warning(self, "Input Error", "Please select a global variable containing the database connection engine.")
            return None
            
        config = {"connection_var": conn_var}

        if self.sql_from_text_radio.isChecked():
            sql_statement = self.sql_statement_edit.toPlainText().strip()
            if not sql_statement:
                QMessageBox.warning(self, "Input Error", "The SQL statement cannot be empty.")
                return None
            config["sql_source"] = "direct"
            config["sql_statement_or_var"] = sql_statement
        else:
            sql_var = self.sql_var_combo.currentText()
            if sql_var == "-- Select Variable --":
                QMessageBox.warning(self, "Input Error", "Please select a variable containing the SQL statement.")
                return None
            config["sql_source"] = "variable"
            config["sql_statement_or_var"] = sql_var
            
        return config

#
# --- [CLASS 6] The Public-Facing Module Class for Executing a Query ---
#
class Mssql_execute_query:
    """A module to execute a non-data-returning SQL statement (e.g., UPDATE, INSERT, DELETE)."""
    def __init__(self, context: Optional[ExecutionContext] = None):
        self.context = context

    def _log(self, message: str):
        if self.context:
            self.context.add_log(message)
        else:
            print(message)

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str],
                           initial_config: Optional[Dict[str, Any]] = None, **kwargs) -> QDialog:
        """The entry point called by the main application to show the configuration GUI."""
        self._log("Opening MS SQL Execute Statement configuration...")
        return _MssqlExecuteQueryDialog(
            global_variables=global_variables,
            parent=parent_window,
            initial_config=initial_config
        )

    def _execute_mssql_execute(self, context: ExecutionContext, config_data: dict) -> str:
        """The executor method that performs the database operation."""
        self.context = context
        
        # 1. Get the connection engine from the specified global variable
        conn_var_name = config_data["connection_var"]
        db_engine = self.context.get_variable(conn_var_name)
        if not hasattr(db_engine, 'connect'):
            raise TypeError(f"Variable '@{conn_var_name}' does not contain a valid database engine.")
        
        # 2. Get the SQL statement from either the direct input or a variable
        sql_statement = ""
        sql_source = config_data.get("sql_source")
        if sql_source == "direct":
            sql_statement = config_data.get("sql_statement_or_var")
        elif sql_source == "variable":
            sql_statement = self.context.get_variable(config_data.get("sql_statement_or_var"))
        
        if not isinstance(sql_statement, str) or not sql_statement.strip():
            raise ValueError("SQL statement is empty or not a valid string.")

        self._log(f"Executing statement: {sql_statement[:200]}...")

        # 3. Execute the statement within a transaction
        try:
            with db_engine.connect() as connection:
                with connection.begin() as transaction: # Automatically commits or rolls back
                    result = connection.execute(text(sql_statement))
                    # result.rowcount is useful for UPDATE, INSERT, DELETE
                    rowcount = result.rowcount
                    self._log(f"Execution successful. {rowcount} row(s) affected.")
            
            return f"Execution successful. {rowcount} row(s) affected."
            
        except Exception as e:
            self._log(f"FATAL ERROR during SQL execution: {e}")
            raise  # Re-raise the exception to make the step fail in the bot
            
class _MssqlCloseDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None, initial_config: Optional[Dict[str, Any]] = None):
        super().__init__(parent)
        self.setWindowTitle("MS SQL Close Connection"); self.setMinimumWidth(400)
        main_layout, form_layout = QVBoxLayout(self), QFormLayout()
        self.conn_var_combo = QComboBox(); self.conn_var_combo.addItems(["-- Select Connection to Close --"] + global_variables)
        form_layout.addRow("Connection Engine Variable:", self.conn_var_combo); main_layout.addLayout(form_layout)
        if initial_config: self.conn_var_combo.setCurrentText(initial_config.get("connection_var", "-- Select Connection to Close --"))
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept); self.button_box.rejected.connect(self.reject); main_layout.addWidget(self.button_box)

    def get_executor_method_name(self): return "_execute_mssql_close"
    def get_assignment_variable(self): return None
    def get_config_data(self):
        conn_var = self.conn_var_combo.currentText()
        if conn_var == "-- Select Connection to Close --":
            QMessageBox.warning(self, "Input Error", "Please select a connection engine variable to close."); return None
        return {"connection_var": conn_var}

#
# --- [CLASS 5] The Public-Facing Module Class for Closing a Connection ---
#
class Mssql_Close_Connection:
    """A module to explicitly close an open MS SQL database connection engine."""
    def __init__(self, context: Optional[ExecutionContext] = None): self.context = context
    def _log(self, message: str):
        if self.context: self.context.add_log(message)
        else: print(message)

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str],
                           initial_config: Optional[Dict[str, Any]] = None, **kwargs) -> QDialog:
        self._log("Opening MS SQL Close Connection configuration...")
        return _MssqlCloseDialog(global_variables, parent_window, initial_config)

    def _execute_mssql_close(self, context: ExecutionContext, config_data: dict):
        self.context = context
        conn_var = config_data["connection_var"]
        self._log(f"Attempting to close and dispose of database engine in '@{conn_var}'...")
        db_engine = self.context.get_variable(conn_var)
        if not db_engine:
            self._log(f"Warning: Variable '@{conn_var}' is empty. Nothing to close."); return
        if not hasattr(db_engine, 'dispose'):
            self._log(f"Warning: Object in '@{conn_var}' is not a database engine."); return
        try:
            db_engine.dispose()
            self._log(f"Successfully disposed of the database engine for '@{conn_var}'.")
        except Exception as e:
            self._log(f"ERROR while disposing engine '@{conn_var}': {e}")