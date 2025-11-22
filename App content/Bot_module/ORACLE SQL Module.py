# File: Bot_module/oracle_module.py

import sys
from typing import Optional, List, Dict, Any
import pandas as pd
from sqlalchemy import create_engine, text

# --- PyQt6 Imports ---
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QDialogButtonBox,
    QComboBox, QWidget, QGroupBox, QMessageBox, QLabel,
    QCheckBox, QApplication, QTextEdit, QTreeWidget, QTreeWidgetItem,
    QHBoxLayout, QHeaderView, QTreeWidgetItemIterator
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal

# --- Database Imports ---
try:
    # Using cx_Oracle as requested
    import cx_Oracle
except ImportError:
    print("Warning: 'cx_Oracle' library not found. Please install it using: pip install cx_Oracle")
    cx_Oracle = None

# --- Main App Imports (Fallback for standalone testing) ---
try:
    from my_lib.shared_context import ExecutionContext
except ImportError:
    print("Warning: Could not import main app libraries. Using fallbacks.")
    class ExecutionContext:
        def add_log(self, message: str): print(message)
        def get_variable(self, name: str):
            if "conn" in name: return create_engine("sqlite:///:memory:")
            if "df" in name: return pd.DataFrame({'ID': [1], 'VALUE': ['a'], 'TYPE': ['t']})
            return None

#
# --- [CLASS 1] HELPER: The GUI Dialog for Oracle Connection ---
#
class _OracleConfigDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None,
                 initial_variable: Optional[str] = None):
        super().__init__(parent)
        self.setWindowTitle("Oracle Database Configuration")
        self.setMinimumWidth(450)
        
        main_layout = QVBoxLayout(self)
        
        connection_group = QGroupBox("Connection Details")
        form_layout = QFormLayout(connection_group)
        self.host_edit = QLineEdit(); self.host_edit.setPlaceholderText("e.g., 192.168.1.100 or db-server.domain.com")
        self.port_edit = QLineEdit("1521"); self.port_edit.setFixedWidth(60)
        self.service_name_edit = QLineEdit(); self.service_name_edit.setPlaceholderText("e.g., ORCL or XEPDB1")
        form_layout.addRow("Host:", self.host_edit)
        form_layout.addRow("Port:", self.port_edit)
        form_layout.addRow("Service Name:", self.service_name_edit)
        main_layout.addWidget(connection_group)
        
        auth_group = QGroupBox("Authentication")
        auth_form = QFormLayout(auth_group)
        self.username_edit = QLineEdit()
        self.password_edit = QLineEdit(); self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        auth_form.addRow("Username:", self.username_edit)
        auth_form.addRow("Password:", self.password_edit)
        main_layout.addWidget(auth_group)
        
        self.test_button = QPushButton("Test Connection")
        main_layout.addWidget(self.test_button)
        
        assignment_group = QGroupBox("Assign Connection Engine to Variable")
        assign_layout = QFormLayout(assignment_group)
        self.new_var_input = QLineEdit("oracle_connection")
        assign_layout.addRow("Variable Name:", self.new_var_input)
        main_layout.addWidget(assignment_group)
        
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)
        
        self.test_button.clicked.connect(self._test_connection)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        
        if initial_config: self._populate_from_initial_config(initial_config, initial_variable)

    def _get_connection_dsn(self) -> Optional[str]:
        host, port, service = self.host_edit.text().strip(), self.port_edit.text().strip(), self.service_name_edit.text().strip()
        if not all([host, port, service]):
            QMessageBox.warning(self, "Input Error", "Host, Port, and Service Name are required.")
            return None
        return f"{host}:{port}/{service}"

    def _test_connection(self):
        if not cx_Oracle:
            QMessageBox.critical(self, "Library Not Found", "'cx_Oracle' is required. Please install it via: pip install cx_Oracle"); return
        
        dsn = self._get_connection_dsn()
        user, pwd = self.username_edit.text().strip(), self.password_edit.text()
        if not dsn or not user:
            QMessageBox.warning(self, "Input Error", "Host, Port, Service Name, and Username are required."); return

        try:
            self.test_button.setText("Testing..."); self.test_button.setEnabled(False)
            QApplication.processEvents()
            conn_url = f"oracle+cx_oracle://{user}:{pwd}@{dsn}"
            engine = create_engine(conn_url)
            with engine.connect(): pass
            QMessageBox.information(self, "Success", "Connection successful!")
        except Exception as e:
            QMessageBox.critical(self, "Connection Failed", f"Could not connect to the database.\n\nError: {e}")
        finally:
            self.test_button.setText("Test Connection"); self.test_button.setEnabled(True)

    def _populate_from_initial_config(self, config, variable):
        self.host_edit.setText(config.get("host", ""))
        self.port_edit.setText(config.get("port", "1521"))
        self.service_name_edit.setText(config.get("service_name", ""))
        self.username_edit.setText(config.get("username", ""))
        if variable: self.new_var_input.setText(variable)

    def get_executor_method_name(self) -> str: return "_execute_oracle_connection"
    def get_config_data(self) -> Optional[Dict[str, Any]]:
        host, port, service = self.host_edit.text().strip(), self.port_edit.text().strip(), self.service_name_edit.text().strip()
        user = self.username_edit.text().strip()
        if not all([host, port, service, user]):
            QMessageBox.warning(self, "Input Error", "Host, Port, Service Name, and Username are required."); return None
        return { "host": host, "port": port, "service_name": service, "username": user, "password": self.password_edit.text() }
    def get_assignment_variable(self) -> Optional[str]:
        var_name = self.new_var_input.text().strip()
        if not var_name:
            QMessageBox.warning(self, "Input Error", "Variable name cannot be empty."); return None
        return var_name

#
# --- [CLASS 1] The Public-Facing Module Class for Oracle Connection ---
#
class OracleDatabase:
    def __init__(self, context: Optional[ExecutionContext] = None): self.context = context
    def _log(self, m: str): (self.context.add_log(m) if self.context else print(m))
    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        self._log("Opening Oracle Database configuration...")
        return _OracleConfigDialog(global_variables, parent_window, **kwargs)

    def _execute_oracle_connection(self, context: ExecutionContext, config_data: dict):
        self.context = context
        if not cx_Oracle:
            raise ImportError("'cx_Oracle' library is not installed. This step cannot be executed.")
        
        user, pwd = config_data["username"], config_data["password"]
        dsn = f"{config_data['host']}:{config_data['port']}/{config_data['service_name']}"
        conn_url = f"oracle+cx_oracle://{user}:{pwd}@{dsn}"
        
        self._log(f"Creating Oracle SQLAlchemy engine for DSN: {dsn} with user: {user}")
        try:
            engine = create_engine(conn_url)
            with engine.connect():
                self._log("Connection successful. Oracle engine created.")
            return engine
        except Exception as e:
            self._log(f"FATAL ERROR during Oracle engine creation: {e}"); raise

#
# --- [CLASS 2] HELPER: The GUI Dialog for Oracle Query ---
#
class _OracleQueryDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None, initial_variable: Optional[str] = None):
        super().__init__(parent)
        self.setWindowTitle("Oracle Query Executor"); self.setMinimumSize(600, 500); self.global_variables = global_variables
        main_layout = QVBoxLayout(self)
        conn_group = QGroupBox("Database Connection"); conn_layout = QFormLayout(conn_group)
        self.conn_var_combo = QComboBox(); self.conn_var_combo.addItem("-- Select Connection Engine --"); self.conn_var_combo.addItems(self.global_variables)
        conn_layout.addRow("Connection Engine:", self.conn_var_combo); main_layout.addWidget(conn_group)
        sql_group = QGroupBox("SQL Statement"); sql_layout = QVBoxLayout(sql_group)
        self.sql_from_text_radio = QRadioButton("Enter SQL Statement directly:")
        self.sql_statement_edit = QTextEdit(); self.sql_statement_edit.setPlaceholderText('SELECT * FROM "MyTable" WHERE "ID" = 1'); self.sql_statement_edit.setFontFamily("Courier New")
        self.sql_from_var_radio = QRadioButton("Get SQL Statement from Variable:")
        self.sql_var_combo = QComboBox(); self.sql_var_combo.addItem("-- Select Variable --"); self.sql_var_combo.addItems(self.global_variables)
        sql_layout.addWidget(self.sql_from_text_radio); sql_layout.addWidget(self.sql_statement_edit); sql_layout.addWidget(self.sql_from_var_radio); sql_layout.addWidget(self.sql_var_combo); main_layout.addWidget(sql_group)
        assignment_group = QGroupBox("Assign Query Result to Variable"); assign_layout = QVBoxLayout(assignment_group)
        self.assign_checkbox = QCheckBox("Assign results (DataFrame) to a variable"); assign_layout.addWidget(self.assign_checkbox)
        self.new_var_radio = QRadioButton("New Variable Name:"); self.new_var_input = QLineEdit("oracle_results")
        self.existing_var_radio = QRadioButton("Existing Variable:")
        self.existing_var_combo = QComboBox(); self.existing_var_combo.addItem("-- Select Variable --"); self.existing_var_combo.addItems(self.global_variables)
        assign_form = QFormLayout(); assign_form.addRow(self.new_var_radio, self.new_var_input); assign_form.addRow(self.existing_var_radio, self.existing_var_combo); assign_layout.addLayout(assign_form); main_layout.addWidget(assignment_group)
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel); main_layout.addWidget(self.button_box)
        self.sql_from_text_radio.toggled.connect(self._toggle_sql_input_widgets); self.assign_checkbox.toggled.connect(self._toggle_assignment_widgets); self.new_var_radio.toggled.connect(self._toggle_assignment_widgets)
        self.button_box.accepted.connect(self.accept); self.button_box.rejected.connect(self.reject)
        self.sql_from_text_radio.setChecked(True); self.assign_checkbox.setChecked(True); self.new_var_radio.setChecked(True)
        self._toggle_sql_input_widgets(); self._toggle_assignment_widgets()
        if initial_config: self._populate_from_initial_config(initial_config, initial_variable)
    def _toggle_sql_input_widgets(self): self.sql_statement_edit.setEnabled(self.sql_from_text_radio.isChecked()); self.sql_var_combo.setEnabled(not self.sql_from_text_radio.isChecked())
    def _toggle_assignment_widgets(self):
        is_assign_enabled = self.assign_checkbox.isChecked()
        self.new_var_radio.setVisible(is_assign_enabled); self.new_var_input.setVisible(is_assign_enabled and self.new_var_radio.isChecked())
        self.existing_var_radio.setVisible(is_assign_enabled); self.existing_var_combo.setVisible(is_assign_enabled and self.existing_var_radio.isChecked())
    def _populate_from_initial_config(self, config, variable):
        self.conn_var_combo.setCurrentText(config.get("connection_var", "-- Select Connection Engine --"))
        if config.get("sql_source", "direct") == "variable": self.sql_from_var_radio.setChecked(True); self.sql_var_combo.setCurrentText(config.get("sql_statement_or_var", "-- Select Variable --"))
        else: self.sql_from_text_radio.setChecked(True); self.sql_statement_edit.setText(config.get("sql_statement_or_var", ""))
        if variable:
            self.assign_checkbox.setChecked(True)
            if variable in self.global_variables: self.existing_var_radio.setChecked(True); self.existing_var_combo.setCurrentText(variable)
            else: self.new_var_radio.setChecked(True); self.new_var_input.setText(variable)
        else: self.assign_checkbox.setChecked(False)
        self._toggle_assignment_widgets()
    def get_executor_method_name(self): return "_execute_oracle_query"
    def get_config_data(self):
        conn_var = self.conn_var_combo.currentText()
        if conn_var == "-- Select Connection Engine --": QMessageBox.warning(self, "Input Error", "Please select a connection engine variable."); return None
        config = {"connection_var": conn_var}
        if self.sql_from_text_radio.isChecked():
            sql = self.sql_statement_edit.toPlainText().strip()
            if not sql: QMessageBox.warning(self, "Input Error", "SQL statement cannot be empty."); return None
            config["sql_source"], config["sql_statement_or_var"] = "direct", sql
        else:
            sql_var = self.sql_var_combo.currentText()
            if sql_var == "-- Select Variable --": QMessageBox.warning(self, "Input Error", "Please select a variable for the SQL statement."); return None
            config["sql_source"], config["sql_statement_or_var"] = "variable", sql_var
        return config
    def get_assignment_variable(self):
        if not self.assign_checkbox.isChecked(): return None
        if self.new_var_radio.isChecked():
            var_name = self.new_var_input.text().strip()
            if not var_name: QMessageBox.warning(self, "Input Error", "New variable name cannot be empty."); return None
            return var_name
        else:
            var_name = self.existing_var_combo.currentText()
            if var_name == "-- Select Variable --": QMessageBox.warning(self, "Input Error", "Please select an existing variable."); return None
            return var_name

#
# --- [CLASS 2] The Public-Facing Module Class for Oracle Query ---
#
class Oracle_Query:
    def __init__(self, context: Optional[ExecutionContext] = None): self.context = context
    def _log(self, m: str): (self.context.add_log(m) if self.context else print(m))
    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        self._log("Opening Oracle Query configuration...")
        return _OracleQueryDialog(global_variables, parent_window, **kwargs)
    def _execute_oracle_query(self, context: ExecutionContext, config_data: dict) -> pd.DataFrame:
        self.context = context
        db_engine = context.get_variable(config_data["connection_var"])
        if not hasattr(db_engine, 'connect'): raise TypeError(f"Variable '@{config_data['connection_var']}' is not a valid database engine.")
        sql = context.get_variable(config_data["sql_statement_or_var"]) if config_data["sql_source"] == "variable" else config_data["sql_statement_or_var"]
        if not isinstance(sql, str) or not sql.strip(): raise ValueError("SQL statement is empty.")
        try:
            self._log(f"Oracle Query: {sql[:200]}...")
            df = pd.read_sql(sql, db_engine)
            self._log(f"Query successful. Returned DataFrame with {len(df)} rows.")
            return df
        except Exception as e:
            self._log(f"FATAL ERROR during Oracle query execution: {e}"); raise

#
# --- [CLASS 3 & 4 & 5] HELPER: Worker thread for fetching Oracle schema ---
#
class _OracleSchemaLoaderThread(QThread):
    schema_ready = pyqtSignal(dict)
    error_occurred = pyqtSignal(str)
    def __init__(self, engine_obj): super().__init__(); self.engine_obj = engine_obj
    def run(self):
        try:
            with self.engine_obj.connect() as connection:
                query = "SELECT OWNER, TABLE_NAME FROM ALL_TABLES ORDER BY OWNER, TABLE_NAME"
                result = connection.execute(text(query))
                schema_dict = {}
                for owner, table_name in result:
                    if owner not in schema_dict: schema_dict[owner] = []
                    schema_dict[owner].append(table_name)
                self.schema_ready.emit(schema_dict)
        except Exception as e: self.error_occurred.emit(f"Failed to load schema. Check connection and permissions. Error: {e}")

#
# --- [CLASSES 3, 4] HELPER: Base GUI Dialog for Oracle Write/Merge ---
#
class _OracleWriteMergeBaseDialog(QDialog):
    def __init__(self, global_variables: List[str], df_variables: List[str], parent: Optional[QWidget] = None, initial_config: Optional[Dict[str, Any]] = None):
        super().__init__(parent)
        self.global_variables, self.df_variables, self.initial_config = global_variables, df_variables, initial_config or {}
        self.source_df_columns, self.target_table_columns = [], []
    def _get_context_variable(self, var_name: str):
        if not hasattr(self.parent(), 'get_variable_for_dialog'):
            QMessageBox.critical(self, "Error", "Main application support function 'get_variable_for_dialog' is missing."); return None
        return self.parent().get_variable_for_dialog(var_name)
    def _load_schema(self):
        engine_obj = self._get_context_variable(self.conn_var_combo.currentText())
        if not engine_obj: QMessageBox.warning(self, "Connection Missing", "Please select a valid connection engine first."); return
        self.load_schema_button.setText("Loading..."); self.load_schema_button.setEnabled(False); self.table_tree.clear()
        self.schema_loader_thread = _OracleSchemaLoaderThread(engine_obj)
        self.schema_loader_thread.schema_ready.connect(self._on_schema_loaded)
        self.schema_loader_thread.error_occurred.connect(lambda e: (QMessageBox.critical(self, "Schema Load Failed", str(e)), self.load_schema_button.setText("Reload Schema"), self.load_schema_button.setEnabled(True)))
        self.schema_loader_thread.start()
    def _on_schema_loaded(self, schema_dict: dict):
        self.table_tree.clear()
        for schema, tables in sorted(schema_dict.items()):
            schema_item = QTreeWidgetItem(self.table_tree, [schema])
            for table in sorted(tables):
                QTreeWidgetItem(schema_item, [table])
        self.load_schema_button.setText("Reload Schema"); self.load_schema_button.setEnabled(True)
        self._restore_tree_selection()
    def _restore_tree_selection(self):
        schema_key, table_key = ("schema", "table") if "schema" in self.initial_config else ("target_schema", "target_table")
        saved_schema, saved_table = self.initial_config.get(schema_key), self.initial_config.get(table_key)
        if not saved_schema or not saved_table: return
        iterator = QTreeWidgetItemIterator(self.table_tree)
        while iterator.value():
            item = iterator.value()
            if item.parent() and item.parent().text(0) == saved_schema and item.text(0) == saved_table:
                self.table_tree.setCurrentItem(item); self.table_tree.scrollToItem(item); break
            iterator += 1
    def get_assignment_variable(self): return None

#
# --- [CLASS 3] HELPER: The GUI Dialog for Oracle Write ---
#
class _OracleWriteDialog(_OracleWriteMergeBaseDialog):
    def __init__(self, global_variables, df_variables, parent, initial_config):
        super().__init__(global_variables, df_variables, parent, initial_config)
        self.setWindowTitle("Oracle Write"); self.setMinimumSize(650, 700)
        main_layout = QVBoxLayout(self)
        source_group = QGroupBox("Source Data"); source_form = QFormLayout(source_group)
        self.conn_var_combo = QComboBox(); self.conn_var_combo.addItems(["-- Select Connection --"] + self.global_variables)
        self.df_var_combo = QComboBox(); self.df_var_combo.addItems(["-- Select DataFrame --"] + self.df_variables)
        source_form.addRow("Connection Engine:", self.conn_var_combo); source_form.addRow("DataFrame:", self.df_var_combo)
        main_layout.addWidget(source_group)
        target_group = QGroupBox("Target Table"); target_layout = QVBoxLayout(target_group)
        self.load_schema_button = QPushButton("Load Database Schema"); target_layout.addWidget(self.load_schema_button)
        self.table_tree = QTreeWidget(); self.table_tree.setHeaderLabels(["OWNER/Schema", "Tables"]); self.table_tree.header().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        target_layout.addWidget(self.table_tree)
        self.create_new_table_checkbox = QCheckBox("Create new table")
        new_table_layout = QHBoxLayout(); self.new_schema_input = QLineEdit(); self.new_schema_input.setPlaceholderText("OWNER/Schema"); self.new_table_input = QLineEdit(); self.new_table_input.setPlaceholderText("New Table Name")
        new_table_layout.addWidget(self.create_new_table_checkbox); new_table_layout.addWidget(QLabel("Schema:")); new_table_layout.addWidget(self.new_schema_input); new_table_layout.addWidget(QLabel("Table:")); new_table_layout.addWidget(self.new_table_input)
        target_layout.addLayout(new_table_layout); main_layout.addWidget(target_group)
        mode_group = QGroupBox("Write Mode"); mode_layout = QFormLayout(mode_group)
        self.write_mode_combo = QComboBox(); self.write_mode_combo.addItems(["append", "replace"])
        mode_layout.addRow("Action:", self.write_mode_combo); main_layout.addWidget(mode_group)
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel); main_layout.addWidget(self.button_box)
        self.load_schema_button.clicked.connect(self._load_schema); self.create_new_table_checkbox.toggled.connect(self._toggle_target_widgets)
        self.table_tree.itemDoubleClicked.connect(lambda item: (self.new_schema_input.setText(item.parent().text(0)), self.new_table_input.setText(item.text(0))) if item.parent() else None)
        self.button_box.accepted.connect(self.accept); self.button_box.rejected.connect(self.reject)
        self._toggle_target_widgets(); self._populate_from_initial_config()
    def _toggle_target_widgets(self): is_new = self.create_new_table_checkbox.isChecked(); self.new_schema_input.setEnabled(is_new); self.new_table_input.setEnabled(is_new); self.table_tree.setEnabled(not is_new)
    def _populate_from_initial_config(self):
        if not self.initial_config: self.create_new_table_checkbox.setChecked(True); self._toggle_target_widgets(); return
        self.conn_var_combo.setCurrentText(self.initial_config.get("connection_var", "-- Select Connection --"))
        self.df_var_combo.setCurrentText(self.initial_config.get("dataframe_var", "-- Select DataFrame --"))
        self.write_mode_combo.setCurrentText(self.initial_config.get("write_mode", "append"))
        is_new = self.initial_config.get("target_type") == "new"
        self.create_new_table_checkbox.setChecked(is_new)
        self.new_schema_input.setText(self.initial_config.get("schema", "")); self.new_table_input.setText(self.initial_config.get("table", ""))
        self._toggle_target_widgets()
    def get_executor_method_name(self): return "_execute_oracle_write"
    def get_config_data(self):
        config = {"connection_var": self.conn_var_combo.currentText(), "dataframe_var": self.df_var_combo.currentText()}
        if config["connection_var"] == "-- Select Connection --" or config["dataframe_var"] == "-- Select DataFrame --": QMessageBox.warning(self, "Input Error", "Please select a Connection and DataFrame."); return None
        if self.create_new_table_checkbox.isChecked():
            config["target_type"] = "new"; schema, table = self.new_schema_input.text().strip(), self.new_table_input.text().strip()
            if not table or not schema: QMessageBox.warning(self, "Input Error", "Schema and Table Name are required for new tables."); return None
            config["schema"], config["table"] = schema, table
        else:
            config["target_type"] = "existing"; item = self.table_tree.currentItem()
            if not item or not item.parent(): QMessageBox.warning(self, "Input Error", "Please select an existing table."); return None
            config["schema"], config["table"] = item.parent().text(0), item.text(0)
        config["write_mode"] = self.write_mode_combo.currentText(); return config

#
# --- [CLASS 3] The Public-Facing Module Class for Oracle Writing ---
#
class Oracle_Write:
    def __init__(self, context: Optional[ExecutionContext] = None): self.context = context
    def _log(self, m: str): (self.context.add_log(m) if self.context else print(m))
    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs):
        self._log("Opening Oracle Write configuration...")
        df_vars = getattr(parent_window, 'get_dataframe_variables', lambda: global_variables)()
        return _OracleWriteDialog(global_variables, df_vars, parent_window, kwargs.get("initial_config"))
    def _execute_oracle_write(self, context: ExecutionContext, config_data: dict):
        self.context = context
        db_engine, df = context.get_variable(config_data["connection_var"]), context.get_variable(config_data["dataframe_var"])
        if not hasattr(db_engine, 'connect'): raise TypeError(f"Variable '@{config_data['connection_var']}' is not a valid database engine.")
        if not isinstance(df, pd.DataFrame): raise TypeError(f"Variable '@{config_data['dataframe_var']}' is not a pandas DataFrame.")
        table_name, schema, write_mode = config_data["table"].upper(), config_data["schema"].upper(), config_data["write_mode"]
        self._log(f"Preparing to write {len(df)} rows to {schema}.{table_name} in '{write_mode}' mode.")
        df.columns = [c.upper() for c in df.columns]
        try:
            df.to_sql(name=table_name, con=db_engine, schema=schema, if_exists=write_mode, index=False)
            self._log(f"Successfully wrote {len(df)} rows. Connection remains open.")
        except Exception as e:
            self._log(f"FATAL ERROR during Oracle write: {e}"); raise

#
# --- [CLASS 4] HELPER: The GUI Dialog for Oracle Merge ---
#
class _OracleMergeDialog(_OracleWriteMergeBaseDialog):
    def __init__(self, global_variables, df_variables, parent, initial_config):
        super().__init__(global_variables, df_variables, parent, initial_config)
        self.setWindowTitle("Oracle Merge (Update/Insert)"); self.setMinimumSize(700, 750)
        self.join_condition_widgets, self.update_column_widgets = [], []
        main_layout = QVBoxLayout(self)
        source_group = QGroupBox("1. Source Data"); source_form = QFormLayout(source_group)
        self.conn_var_combo = QComboBox(); self.conn_var_combo.addItems(["-- Select Connection --"] + self.global_variables)
        self.df_var_combo = QComboBox(); self.df_var_combo.addItems(["-- Select DataFrame --"] + self.df_variables)
        source_form.addRow("Connection Engine:", self.conn_var_combo); source_form.addRow("Source DataFrame:", self.df_var_combo); main_layout.addWidget(source_group)
        target_group = QGroupBox("2. Target Table"); target_layout = QVBoxLayout(target_group)
        self.load_schema_button = QPushButton("Load Database Schema"); target_layout.addWidget(self.load_schema_button)
        self.table_tree = QTreeWidget(); self.table_tree.setHeaderLabels(["OWNER/Schema", "Tables"]); self.table_tree.header().setSectionResizeMode(QHeaderView.ResizeMode.Stretch); target_layout.addWidget(self.table_tree); main_layout.addWidget(target_group)
        config_group = QGroupBox("3. Merge Configuration"); config_layout = QVBoxLayout(config_group)
        join_group = QGroupBox("ON (Join Condition)"); self.join_layout = QVBoxLayout(join_group)
        self.add_join_condition_button = QPushButton("➕ Add Condition"); self.join_layout.addWidget(self.add_join_condition_button); config_layout.addWidget(join_group)
        update_group = QGroupBox("WHEN MATCHED THEN UPDATE"); self.update_layout = QVBoxLayout(update_group)
        self.add_update_column_button = QPushButton("➕ Add Column to Update"); self.update_layout.addWidget(self.add_update_column_button); config_layout.addWidget(update_group)
        insert_group = QGroupBox("WHEN NOT MATCHED THEN INSERT"); insert_layout = QVBoxLayout(insert_group); insert_layout.addWidget(QLabel("All columns from the source DataFrame will be inserted.")); config_layout.addWidget(insert_group)
        main_layout.addWidget(config_group)
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel); main_layout.addWidget(self.button_box)
        self.load_schema_button.clicked.connect(self._load_schema); self.df_var_combo.currentTextChanged.connect(self._on_source_df_changed); self.table_tree.currentItemChanged.connect(self._on_target_table_changed)
        self.add_join_condition_button.clicked.connect(lambda: self._add_row_to_layout(self.join_layout, self.join_condition_widgets, self._create_join_condition_row))
        self.add_update_column_button.clicked.connect(lambda: self._add_row_to_layout(self.update_layout, self.update_column_widgets, self._create_update_column_row))
        self.button_box.accepted.connect(self.accept); self.button_box.rejected.connect(self.reject)
        self._populate_from_initial_config(); self._on_source_df_changed(self.df_var_combo.currentText())
    def _on_source_df_changed(self, df_var_name):
        df = self._get_context_variable(df_var_name) if df_var_name != "-- Select DataFrame --" else None
        self.source_df_columns = sorted([c.upper() for c in df.columns]) if isinstance(df, pd.DataFrame) else []
        self._rebuild_dynamic_combos()
    def _on_target_table_changed(self, current, previous):
        if not (current and current.parent()): return
        engine_obj = self._get_context_variable(self.conn_var_combo.currentText())
        if not engine_obj: self.target_table_columns = []; self._rebuild_dynamic_combos(); return
        schema, table = current.parent().text(0), current.text(0)
        try:
            with engine_obj.connect() as connection:
                cursor = connection.connection.cursor()
                self.target_table_columns = sorted([row[2] for row in cursor.columns(schema=schema, table=table)])
        except Exception as e:
            QMessageBox.warning(self, "Column Load Error", f"Could not fetch columns for {schema}.{table}:\n{e}"); self.target_table_columns = []
        self._rebuild_dynamic_combos()
    def _create_join_condition_row(self):
        widget, layout = QWidget(), QHBoxLayout(widget); layout.setContentsMargins(0,0,0,0)
        source_combo = QComboBox(); source_combo.addItems(["-- Source --"] + self.source_df_columns); target_combo = QComboBox(); target_combo.addItems(["-- Target --"] + self.target_table_columns)
        rm_btn = QPushButton("➖"); rm_btn.setFixedWidth(30); rm_btn.clicked.connect(lambda: self._remove_row_from_layout(widget, self.join_layout, self.join_condition_widgets))
        layout.addWidget(QLabel("Source:")); layout.addWidget(source_combo); layout.addWidget(QLabel(" = Target:")); layout.addWidget(target_combo); layout.addWidget(rm_btn)
        return widget, {'source': source_combo, 'target': target_combo}
    def _create_update_column_row(self):
        widget, layout = QWidget(), QHBoxLayout(widget); layout.setContentsMargins(0,0,0,0)
        target_combo = QComboBox(); target_combo.addItems(["-- Target --"] + self.target_table_columns); source_combo = QComboBox(); source_combo.addItems(["-- Source --"] + self.source_df_columns)
        rm_btn = QPushButton("➖"); rm_btn.setFixedWidth(30); rm_btn.clicked.connect(lambda: self._remove_row_from_layout(widget, self.update_layout, self.update_column_widgets))
        layout.addWidget(QLabel("SET Target:")); layout.addWidget(target_combo); layout.addWidget(QLabel(" = Source:")); layout.addWidget(source_combo); layout.addWidget(rm_btn)
        return widget, {'target': target_combo, 'source': source_combo}
    def _add_row_to_layout(self, layout, widget_list, create_row_func): row_widget, combo_dict = create_row_func(); layout.addWidget(row_widget); widget_list.append((row_widget, combo_dict))
    def _remove_row_from_layout(self, widget_to_remove, layout, widget_list): [widget_list.remove(item) for item in widget_list if item[0] is widget_to_remove]; widget_to_remove.deleteLater()
    def _rebuild_dynamic_combos(self):
        for _, combos in self.join_condition_widgets:
            combos['source'].clear(); combos['source'].addItems(["-- Source --"] + self.source_df_columns); combos['target'].clear(); combos['target'].addItems(["-- Target --"] + self.target_table_columns)
        for _, combos in self.update_column_widgets:
            combos['source'].clear(); combos['source'].addItems(["-- Source --"] + self.source_df_columns); combos['target'].clear(); combos['target'].addItems(["-- Target --"] + self.target_table_columns)
    def _populate_from_initial_config(self):
        if not self.initial_config: return
        self.conn_var_combo.setCurrentText(self.initial_config.get("connection_var", "-- Select Connection --")); self.df_var_combo.setCurrentText(self.initial_config.get("dataframe_var", "-- Select DataFrame --"))
        if self.initial_config.get("target_table"): self._load_schema()
        for c in self.initial_config.get("on_conditions", []): self._add_row_to_layout(self.join_layout, self.join_condition_widgets, self._create_join_condition_row); self.join_condition_widgets[-1][1]['source'].setCurrentText(c['source']); self.join_condition_widgets[-1][1]['target'].setCurrentText(c['target'])
        for u in self.initial_config.get("update_columns", []): self._add_row_to_layout(self.update_layout, self.update_column_widgets, self._create_update_column_row); self.update_column_widgets[-1][1]['target'].setCurrentText(u['target']); self.update_column_widgets[-1][1]['source'].setCurrentText(u['source'])
    def get_executor_method_name(self): return "_execute_oracle_merge"
    def get_config_data(self):
        config = {"connection_var": self.conn_var_combo.currentText(), "dataframe_var": self.df_var_combo.currentText()}
        if config["connection_var"] == "-- Select Connection --" or config["dataframe_var"] == "-- Select DataFrame --": QMessageBox.warning(self, "Input Error", "Please select Connection and DataFrame."); return None
        item = self.table_tree.currentItem();
        if not item or not item.parent(): QMessageBox.warning(self, "Input Error", "Please select a target table."); return None
        config["target_schema"], config["target_table"] = item.parent().text(0), item.text(0)
        config["on_conditions"] = [{'source': c['source'].currentText(), 'target': c['target'].currentText()} for _, c in self.join_condition_widgets]
        if not config["on_conditions"] or any(c['source'] == "-- Source --" or c['target'] == "-- Target --" for c in config["on_conditions"]): QMessageBox.warning(self, "Input Error", "All 'ON' conditions must be fully specified."); return None
        config["update_columns"] = [{'target': c['target'].currentText(), 'source': c['source'].currentText()} for _, c in self.update_column_widgets]
        if not config["update_columns"]: QMessageBox.warning(self, "Input Error", "At least one 'UPDATE' column is required."); return None
        return config

#
# --- [CLASS 4] The Public-Facing Module Class for Oracle Merging ---
#
class Oracle_Merge:
    def __init__(self, context: Optional[ExecutionContext] = None): self.context = context
    def _log(self, m: str): (self.context.add_log(m) if self.context else print(m))
    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs):
        self._log("Opening Oracle Merge configuration...")
        df_vars = getattr(parent_window, 'get_dataframe_variables', lambda: global_variables)()
        return _OracleMergeDialog(global_variables, df_vars, parent_window, kwargs.get("initial_config"))
    def _execute_oracle_merge(self, context: ExecutionContext, config_data: dict):
        self.context = context
        db_engine = context.get_variable(config_data["connection_var"])
        source_df = context.get_variable(config_data["dataframe_var"])
        target_schema, target_table = config_data["target_schema"].upper(), config_data["target_table"].upper()
        if not hasattr(db_engine, 'connect'): raise TypeError(f"Variable '@{config_data['connection_var']}' is not a valid database engine.")
        if not isinstance(source_df, pd.DataFrame) or source_df.empty: raise TypeError(f"Variable '@{config_data['dataframe_var']}' is not a non-empty pandas DataFrame.")
        staging_table_name = "MERGE_STAGE_TEMP"
        self._log(f"Preparing to merge {len(source_df)} rows into {target_schema}.{target_table}.")
        source_df.columns = [c.upper() for c in source_df.columns]
        try:
            self._log(f"Writing data to staging table '{staging_table_name}'.")
            source_df.to_sql(staging_table_name, db_engine, if_exists='replace', index=False)
        except Exception as e: raise IOError(f"Failed to write data to staging table. Error: {e}")
        on_conditions = " AND ".join([f'T."{c["target"].upper()}" = S."{c["source"].upper()}"' for c in config_data["on_conditions"]])
        update_clauses = ", ".join([f'T."{c["target"].upper()}" = S."{c["source"].upper()}"' for c in config_data["update_columns"]])
        insert_cols = ", ".join([f'"{col.upper()}"' for col in source_df.columns])
        source_cols_for_values = ", ".join([f'S."{col.upper()}"' for col in source_df.columns])
        merge_sql = f"""
        MERGE INTO "{target_schema}"."{target_table}" T USING {staging_table_name} S ON ({on_conditions})
        WHEN MATCHED THEN UPDATE SET {update_clauses}
        WHEN NOT MATCHED THEN INSERT ({insert_cols}) VALUES ({source_cols_for_values})
        """
        self._log("Executing Oracle MERGE statement...")
        try:
            with db_engine.begin() as conn:
                result = conn.execute(text(merge_sql))
                self._log(f"MERGE successful. {result.rowcount} rows affected.")
                conn.execute(text(f"DROP TABLE {staging_table_name}"))
        except Exception as e:
            self._log(f"FATAL ERROR during Oracle MERGE execution: {e}"); raise

#
# --- [CLASS 5] Close Connection ---
#
class Oracle_Close_Connection:
    def __init__(self, context: Optional[ExecutionContext] = None): self.context = context
    def _log(self, m: str): (self.context.add_log(m) if self.context else print(m))
    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs):
        self._log("Opening Oracle Close Connection configuration..."); return _OracleCloseDialog(global_variables, parent_window, **kwargs)
    def _execute_oracle_close(self, context: ExecutionContext, config_data: dict):
        self.context = context
        conn_var = config_data["connection_var"]
        self._log(f"Attempting to close and dispose of Oracle engine in '@{conn_var}'...")
        db_engine = context.get_variable(conn_var)
        if not db_engine: self._log(f"Warning: Variable '@{conn_var}' is empty."); return
        if not hasattr(db_engine, 'dispose'): self._log(f"Warning: Object in '@{conn_var}' is not a database engine."); return
        try:
            db_engine.dispose()
            self._log(f"Successfully disposed of the Oracle database engine for '@{conn_var}'.")
        except Exception as e:
            self._log(f"ERROR while disposing Oracle engine '@{conn_var}': {e}")