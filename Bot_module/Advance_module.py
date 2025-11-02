# File: Bot_module/data_hub_module.py

import os
import sys
import pandas as pd
import win32com.client
import pythoncom
from datetime import datetime
from typing import Optional, List, Dict, Any
import fnmatch

# --- PyQt6 Imports ---
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QDialogButtonBox,
    QComboBox, QFileDialog, QTableView, QListWidget, QListWidgetItem, QHeaderView,
    QGroupBox, QCheckBox, QRadioButton, QWidget, QHBoxLayout, QMessageBox, QLabel,
    QDateTimeEdit, QTreeWidgetItemIterator, QTreeWidget, QTreeWidgetItem, QTabWidget
)
from PyQt6.QtCore import Qt, QVariant, QThread, pyqtSignal, QDateTime
from PyQt6.QtGui import QStandardItemModel, QStandardItem

# --- Main App Imports ---
try:
    from my_lib.shared_context import ExecutionContext
except ImportError:
    print("Warning: Could not import main app libraries. Using fallbacks.")
    class ExecutionContext:
        def add_log(self, message: str):
            print(message)

#
# --- HELPER: Pandas Table Model ---
#
class _PandasModel(QStandardItemModel):
    def __init__(self, df: pd.DataFrame, parent=None):
        super().__init__(parent)
        self.setColumnCount(len(df.columns))
        self.setHorizontalHeaderLabels(list(df.columns))
        self.setRowCount(len(df))
        
        for i in range(len(df)):
            for j in range(len(df.columns)):
                value = df.iloc[i, j]
                item = QStandardItem(str(value))
                self.setItem(i, j, item)

#
# --- HELPER: Outlook Folder Loading Thread ---
#
class _OutlookFolderLoader(QThread):
    folders_loaded = pyqtSignal(object)
    error_occurred = pyqtSignal(str)

    def run(self):
        try:
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            folder_structure = []
            for folder in outlook.Folders:
                folder_data = self._process_folder(folder)
                if folder_data:
                    folder_structure.append(folder_data)
            self.folders_loaded.emit(folder_structure)
        except Exception as e:
            self.error_occurred.emit(f"Could not connect to Outlook or read folders.\n"
                                     f"Please ensure Outlook is running.\nError: {e}")
        finally:
            pythoncom.CoUninitialize()

    def _process_folder(self, folder):
        try:
            folder_data = {"name": folder.Name, "entry_id": folder.EntryID, "subfolders": []}
            for subfolder in folder.Folders:
                subfolder_data = self._process_folder(subfolder)
                if subfolder_data:
                    folder_data["subfolders"].append(subfolder_data)
            return folder_data
        except Exception as e:
            print(f"Skipping folder {folder.Name}: {e}")
            return None

#
# --- HELPER: The Main Custom GUI Dialog ---
#
class _DataHubConfigDialog(QDialog):
    """
    The single custom GUI with two tabs: File Loader and Outlook Reader.
    """
    
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Data Hub Configuration")
        self.setMinimumSize(800, 700)
        self.global_variables = global_variables
        self.df_preview: Optional[pd.DataFrame] = None
        self.folder_loader_thread = None

        main_layout = QVBoxLayout(self)

        # --- 1. Tab Widget ---
        self.tab_widget = QTabWidget()
        
        # Create the two tabs
        self.file_loader_tab = QWidget()
        self.outlook_reader_tab = QWidget()
        
        # Populate each tab
        self._create_file_loader_tab(self.file_loader_tab)
        self._create_outlook_reader_tab(self.outlook_reader_tab)
        
        self.tab_widget.addTab(self.file_loader_tab, "File Loader")
        self.tab_widget.addTab(self.outlook_reader_tab, "Outlook Reader")
        
        main_layout.addWidget(self.tab_widget)

        # --- 2. Shared Variable Assignment Group ---
        # This is outside the tabs and applies to whichever task is run
        assignment_group = QGroupBox("Assign Results to Variable")
        assignment_layout = QVBoxLayout(assignment_group)
        self.assign_checkbox = QCheckBox("Assign results (DataFrame) to a variable")
        self.assign_checkbox.setChecked(True)
        assignment_layout.addWidget(self.assign_checkbox)
        
        self.new_var_radio = QRadioButton("New Variable Name:")
        self.new_var_input = QLineEdit("new_data") # Default name
        self.existing_var_radio = QRadioButton("Existing Variable:")
        self.existing_var_combo = QComboBox()
        self.existing_var_combo.addItem("-- Select Variable --")
        self.existing_var_combo.addItems(self.global_variables)
        self.new_var_radio.setChecked(True)

        assign_form = QFormLayout()
        assign_form.addRow(self.new_var_radio, self.new_var_input)
        assign_form.addRow(self.existing_var_radio, self.existing_var_combo)
        assignment_layout.addLayout(assign_form)
        
        main_layout.addWidget(assignment_group)

        # --- 3. Dialog Buttons ---
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # --- Connections ---
        self.assign_checkbox.toggled.connect(self._toggle_assignment_widgets)
        self.new_var_radio.toggled.connect(self._toggle_assignment_widgets)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        
        # --- Initial State ---
        self._toggle_assignment_widgets()

    # --- TAB 1: FILE LOADER ---
    def _create_file_loader_tab(self, tab_widget: QWidget):
        layout = QVBoxLayout(tab_widget)
        
        # File Selection
        file_group = QGroupBox("File Source")
        file_layout = QFormLayout(file_group)
        self.file_type_combo = QComboBox()
        self.file_type_combo.addItems(["Excel", "CSV", "JSON", "Text (Delimited)"])
        file_layout.addRow("File Type:", self.file_type_combo)
        path_layout = QHBoxLayout()
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("Select a file...")
        self.browse_button = QPushButton("Browse...")
        path_layout.addWidget(self.path_edit)
        path_layout.addWidget(self.browse_button)
        file_layout.addRow("File Path:", path_layout)
        layout.addWidget(file_group)

        # File Options
        self.options_group = QGroupBox("File Options")
        self.options_layout = QFormLayout(self.options_group)
        self.sheet_name_edit = QLineEdit("0")
        self.sheet_name_label = QLabel("Sheet Name (or index):")
        self.options_layout.addRow(self.sheet_name_label, self.sheet_name_edit)
        self.delimiter_edit = QLineEdit(",")
        self.delimiter_label = QLabel("Delimiter:")
        self.options_layout.addRow(self.delimiter_label, self.delimiter_edit)
        self.orient_edit = QLineEdit("records")
        self.orient_label = QLabel("JSON Orient:")
        self.options_layout.addRow(self.orient_label, self.orient_edit)
        layout.addWidget(self.options_group)

        # Preview
        preview_group = QGroupBox("Data Preview (First 50 Rows)")
        preview_layout = QVBoxLayout(preview_group)
        self.preview_button = QPushButton("Load Preview")
        self.preview_table = QTableView()
        self.preview_table.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        preview_layout.addWidget(self.preview_button)
        preview_layout.addWidget(self.preview_table)
        layout.addWidget(preview_group)

        # Column Selection
        column_group = QGroupBox("Column Selection")
        column_layout = QVBoxLayout(column_group)
        self.column_list_widget = QListWidget()
        column_layout.addWidget(self.column_list_widget)
        col_button_layout = QHBoxLayout()
        self.select_all_cols_button = QPushButton("Select All")
        self.deselect_all_cols_button = QPushButton("Deselect All")
        col_button_layout.addWidget(self.select_all_cols_button)
        col_button_layout.addWidget(self.deselect_all_cols_button)
        column_layout.addLayout(col_button_layout)
        layout.addWidget(column_group)
        
        # Connections
        self.browse_button.clicked.connect(self._on_browse_file)
        self.preview_button.clicked.connect(self._on_preview_file)
        self.file_type_combo.currentTextChanged.connect(self._update_file_options_ui)
        self.select_all_cols_button.clicked.connect(self._select_all_columns)
        self.deselect_all_cols_button.clicked.connect(self._deselect_all_columns)

        self._update_file_options_ui()

    # --- TAB 2: OUTLOOK READER ---
    def _create_outlook_reader_tab(self, tab_widget: QWidget):
        layout = QVBoxLayout(tab_widget)
        
        # Folder Selection
        folder_group = QGroupBox("Select Folders to Read")
        folder_layout = QVBoxLayout(folder_group)
        self.load_folders_button = QPushButton("Load Outlook Folders")
        self.folder_tree = QTreeWidget()
        self.folder_tree.setHeaderLabels(["Outlook Folders"])
        self.folder_tree.setToolTip("Check the boxes for all folders you want to search.")
        folder_layout.addWidget(self.load_folders_button)
        folder_layout.addWidget(self.folder_tree)
        layout.addWidget(folder_group)
        
        # Time/Status Filter
        time_group = QGroupBox("Filter by Time and Status")
        time_layout = QFormLayout(time_group)
        self.end_time_edit = QDateTimeEdit(QDateTime.currentDateTime())
        self.end_time_edit.setCalendarPopup(True)
        self.start_time_edit = QDateTimeEdit(QDateTime.currentDateTime().addDays(-7))
        self.start_time_edit.setCalendarPopup(True)
        self.use_current_end_time_check = QCheckBox("Use Current Datetime")
        self.use_current_end_time_check.setToolTip("If checked, the 'To' date will be the exact time the bot runs.")
        end_time_layout = QHBoxLayout()
        end_time_layout.addWidget(self.end_time_edit)
        end_time_layout.addWidget(self.use_current_end_time_check)
        time_layout.addRow("From (Start Date):", self.start_time_edit)
        time_layout.addRow("To (End Date):", end_time_layout)
        self.read_status_combo = QComboBox()
        self.read_status_combo.addItems(["All Emails", "Unread Only", "Read Only"])
        time_layout.addRow("Status:", self.read_status_combo)
        layout.addWidget(time_group)
        
        # Content Filter
        filter_group = QGroupBox("Filter by Content (Use * as wildcard)")
        filter_layout = QFormLayout(filter_group)
        self.sender_filter = QLineEdit()
        self.sender_filter.setPlaceholderText("e.g., *example.com or user@*")
        self.subject_filter = QLineEdit()
        self.subject_filter.setPlaceholderText("e.g., *invoice* or Report*")
        self.body_filter = QLineEdit()
        self.body_filter.setPlaceholderText("e.g., *payment due*")
        self.attachment_filter = QLineEdit()
        self.attachment_filter.setPlaceholderText("e.g., *.pdf or *data.xlsx")
        filter_layout.addRow("From Sender:", self.sender_filter)
        filter_layout.addRow("Subject:", self.subject_filter)
        filter_layout.addRow("Body contains:", self.body_filter)
        filter_layout.addRow("Attachment Name:", self.attachment_filter)
        layout.addWidget(filter_group)

        # Attachment Options
        attachment_group = QGroupBox("Attachment Options")
        attachment_layout = QFormLayout(attachment_group)
        self.save_attachments_checkbox = QCheckBox("Save attachments from matching emails")
        attachment_layout.addRow(self.save_attachments_checkbox)
        attachment_path_layout = QHBoxLayout()
        self.attachment_path_edit = QLineEdit()
        self.attachment_path_edit.setPlaceholderText("Select a folder to save attachments...")
        self.browse_save_path_button = QPushButton("Browse...")
        attachment_path_layout.addWidget(self.attachment_path_edit)
        attachment_path_layout.addWidget(self.browse_save_path_button)
        attachment_layout.addRow("Save Location:", attachment_path_layout)
        layout.addWidget(attachment_group)
        
        # Connections
        self.load_folders_button.clicked.connect(self._start_folder_load)
        self.browse_save_path_button.clicked.connect(self._on_browse_save_path)
        self.save_attachments_checkbox.toggled.connect(self._toggle_save_path_widgets)
        self.use_current_end_time_check.toggled.connect(self.end_time_edit.setDisabled)
        
        # Initial State
        self._toggle_save_path_widgets(False)
        self.use_current_end_time_check.setChecked(True)
        self.end_time_edit.setDisabled(True)

    # --- FILE LOADER METHODS ---
    def _update_file_options_ui(self):
        file_type = self.file_type_combo.currentText()
        is_excel = (file_type == "Excel")
        is_csv = (file_type == "CSV")
        is_json = (file_type == "JSON")
        is_text = (file_type == "Text (Delimited)")
        self.sheet_name_label.setVisible(is_excel)
        self.sheet_name_edit.setVisible(is_excel)
        self.delimiter_label.setVisible(is_csv or is_text)
        self.delimiter_edit.setVisible(is_csv or is_text)
        if is_text: self.delimiter_edit.setText(r"\t")
        else: self.delimiter_edit.setText(",")
        self.orient_label.setVisible(is_json)
        self.orient_edit.setVisible(is_json)

    def _on_browse_file(self):
        file_type = self.file_type_combo.currentText()
        if file_type == "Excel": filt = "Excel Files (*.xlsx *.xls)"
        elif file_type == "CSV": filt = "CSV Files (*.csv)"
        elif file_type == "JSON": filt = "JSON Files (*.json)"
        elif file_type == "Text (Delimited)": filt = "Text Files (*.txt *.log *.dat)"
        else: filt = "All Files (*)"
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Data File", "", filt)
        if file_path: self.path_edit.setText(file_path)

    def _on_preview_file(self):
        file_path = self.path_edit.text()
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "File Not Found", "Please select a valid file path.")
            return
        file_type = self.file_type_combo.currentText()
        try:
            if file_type == "Excel":
                sheet = self.sheet_name_edit.text()
                sheet = int(sheet) if sheet.isdigit() else sheet
                self.df_preview = pd.read_excel(file_path, sheet_name=sheet, nrows=50)
            elif file_type == "CSV":
                self.df_preview = pd.read_csv(file_path, delimiter=self.delimiter_edit.text(), nrows=50)
            elif file_type == "JSON":
                self.df_preview = pd.read_json(file_path, orient=self.orient_edit.text(), nrows=50)
            elif file_type == "Text (Delimited)":
                delim = self.delimiter_edit.text()
                if delim == r'\t': delim = '\t'
                self.df_preview = pd.read_csv(file_path, delimiter=delim, nrows=50)
            
            model = _PandasModel(self.df_preview)
            self.preview_table.setModel(model)
            self.column_list_widget.clear()
            for col in self.df_preview.columns:
                item = QListWidgetItem(col)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Checked)
                self.column_list_widget.addItem(item)
        except Exception as e:
            QMessageBox.critical(self, "Preview Error", f"Could not load file preview:\n{e}")
            self.preview_table.setModel(None)
            self.column_list_widget.clear()
            self.df_preview = None

    def _select_all_columns(self):
        for i in range(self.column_list_widget.count()):
            self.column_list_widget.item(i).setCheckState(Qt.CheckState.Checked)

    def _deselect_all_columns(self):
        for i in range(self.column_list_widget.count()):
            self.column_list_widget.item(i).setCheckState(Qt.CheckState.Unchecked)

    # --- OUTLOOK READER METHODS ---
    def _start_folder_load(self):
        self.load_folders_button.setText("Loading... Please Wait")
        self.load_folders_button.setEnabled(False)
        self.folder_tree.clear()
        self.folder_loader_thread = _OutlookFolderLoader()
        self.folder_loader_thread.folders_loaded.connect(self._on_folders_loaded)
        self.folder_loader_thread.error_occurred.connect(self._on_folder_load_error)
        self.folder_loader_thread.start()

    def _on_folder_load_error(self, error_message: str):
        QMessageBox.critical(self, "Outlook Error", error_message)
        self.load_folders_button.setText("Load Outlook Folders")
        self.load_folders_button.setEnabled(True)

    def _on_folders_loaded(self, folder_structure: List[Dict]):
        self.folder_tree.clear()
        for folder_data in folder_structure:
            self._add_folder_to_tree(self.folder_tree.invisibleRootItem(), folder_data)
        self.folder_tree.expandToDepth(0)
        self.load_folders_button.setText("Reload Folders")
        self.load_folders_button.setEnabled(True)

    def _add_folder_to_tree(self, parent_item: QTreeWidgetItem, folder_data: Dict):
        item = QTreeWidgetItem(parent_item, [folder_data["name"]])
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(0, Qt.CheckState.Unchecked)
        item.setData(0, Qt.ItemDataRole.UserRole, folder_data["entry_id"])
        for subfolder_data in folder_data["subfolders"]:
            self._add_folder_to_tree(item, subfolder_data)

    def _get_selected_folders(self) -> List[str]:
        selected_ids = []
        iterator = QTreeWidgetItemIterator(self.folder_tree, QTreeWidgetItemIterator.IteratorFlag.All)
        while iterator.value():
            item = iterator.value()
            if item.checkState(0) == Qt.CheckState.Checked:
                entry_id = item.data(0, Qt.ItemDataRole.UserRole)
                if entry_id:
                    selected_ids.append(entry_id)
            iterator += 1
        return selected_ids

    def _on_browse_save_path(self):
        path = QFileDialog.getExistingDirectory(self, "Select Folder to Save Attachments")
        if path: self.attachment_path_edit.setText(path)

    def _toggle_save_path_widgets(self, checked: bool):
        self.attachment_path_edit.setEnabled(checked)
        self.browse_save_path_button.setEnabled(checked)

    # --- SHARED METHODS ---
    def _toggle_assignment_widgets(self):
        is_assign_enabled = self.assign_checkbox.isChecked()
        self.new_var_radio.setVisible(is_assign_enabled)
        self.new_var_input.setVisible(is_assign_enabled and self.new_var_radio.isChecked())
        self.existing_var_radio.setVisible(is_assign_enabled)
        self.existing_var_combo.setVisible(is_assign_enabled and self.existing_var_radio.isChecked())

    # --- Methods for main_app to call ---

    def get_executor_method_name(self) -> str:
        # This points to the single, smart executor
        return "_execute_data_hub_task"

    def _get_file_loader_config(self) -> Optional[Dict[str, Any]]:
        if not self.path_edit.text() or not os.path.exists(self.path_edit.text()):
            QMessageBox.warning(self, "Input Error (File Loader)", "Please provide a valid file path.")
            return None
        selected_columns = []
        for i in range(self.column_list_widget.count()):
            item = self.column_list_widget.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                selected_columns.append(item.text())
        if self.df_preview is not None and not selected_columns:
            QMessageBox.warning(self, "Input Error (File Loader)", "Please select at least one column.")
            return None
        return {
            "file_path": self.path_edit.text(),
            "file_type": self.file_type_combo.currentText(),
            "sheet_name": self.sheet_name_edit.text(),
            "delimiter": self.delimiter_edit.text(),
            "json_orient": self.orient_edit.text(),
            "selected_columns": selected_columns
        }

    def _get_outlook_reader_config(self) -> Optional[Dict[str, Any]]:
        selected_folder_ids = self._get_selected_folders()
        if not selected_folder_ids:
            QMessageBox.warning(self, "Input Error (Outlook Reader)", "Please load and select at least one Outlook folder.")
            return None
        save_attachments = self.save_attachments_checkbox.isChecked()
        save_path = self.attachment_path_edit.text()
        if save_attachments and not os.path.isdir(save_path):
            QMessageBox.warning(self, "Input Error (Outlook Reader)", "Please select a valid folder for saving attachments.")
            return None
        return {
            "selected_folder_ids": selected_folder_ids,
            "start_time": self.start_time_edit.dateTime().toString(Qt.DateFormat.ISODate),
            "end_time": self.end_time_edit.dateTime().toString(Qt.DateFormat.ISODate),
            "use_current_end_time": self.use_current_end_time_check.isChecked(),
            "read_status_filter": self.read_status_combo.currentText(),
            "sender_filter": self.sender_filter.text(),
            "subject_filter": self.subject_filter.text(),
            "body_filter": self.body_filter.text(),
            "attachment_filter": self.attachment_filter.text(),
            "save_attachments": save_attachments,
            "save_path": save_path
        }

    def get_config_data(self) -> Optional[Dict[str, Any]]:
        """Returns a master config dict. The executor will know which task to run."""
        
        current_tab_index = self.tab_widget.currentIndex()
        task_to_run = "file_loader" if current_tab_index == 0 else "outlook_reader"
        
        file_config = None
        outlook_config = None
        
        # Validate the config for the *active* tab
        if task_to_run == "file_loader":
            file_config = self._get_file_loader_config()
            if file_config is None: return None # Validation failed
        else: # task_to_run == "outlook_reader"
            outlook_config = self._get_outlook_reader_config()
            if outlook_config is None: return None # Validation failed
            
        # Return the master config
        return {
            "task_to_run": task_to_run,
            "file_loader_config": file_config or {}, # Send empty dict if not active
            "outlook_reader_config": outlook_config or {}, # Send empty dict if not active
        }

    def get_assignment_variable(self) -> Optional[str]:
        """Returns the shared variable name."""
        if not self.assign_checkbox.isChecked():
            return None
        
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
# --- The Module Class main_app Discovers ---
#
class DataHub:
    """
    A unified module to load data from local files (Excel, CSV)
    or from Microsoft Outlook emails.
    """

    def __init__(self, context: Optional[ExecutionContext] = None):
        self.context = context

    def _log(self, message: str):
        if self.context:
            self.context.add_log(message)
        else:
            print(message)

    #
    # --- 1. The "Configuration" Method (User-facing) ---
    #
    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str]) -> QDialog:
        """
        Configures data loading from either a File or Outlook.
        The active tab in the dialog will be the task that runs.
        
        MAGIC METHOD: main_app will call this to get the custom GUI.
        """
        self._log("Opening Data Hub configuration...")
        return _DataHubConfigDialog(global_variables=global_variables, parent=parent_window)

    #
    # --- 2. The "Execution" Methods (Hidden) ---
    #
    def _execute_data_hub_task(self, context, config_data: dict) -> pd.DataFrame:
        """
        EXECUTOR: This single method is called by the bot.
        It looks at the 'task_to_run' key to decide which sub-task to perform.
        """
        self.context = context
        task = config_data.get("task_to_run")
        
        if task == "file_loader":
            self._log("Data Hub executing: File Loader")
            file_config = config_data.get("file_loader_config")
            return self._execute_file_load(file_config)
        elif task == "outlook_reader":
            self._log("Data Hub executing: Outlook Reader")
            outlook_config = config_data.get("outlook_reader_config")
            return self._execute_email_read(outlook_config)
        else:
            raise ValueError(f"Invalid task specified in Data Hub config: {task}")

    def _execute_file_load(self, config_data: dict) -> pd.DataFrame:
        """Internal logic for loading a file from disk."""
        path = config_data.get("file_path")
        file_type = config_data.get("file_type")
        
        try:
            df = None
            if file_type == "Excel":
                sheet = config_data.get("sheet_name", "0")
                sheet = int(sheet) if sheet.isdigit() else sheet
                self._log(f"Loading Excel: {os.path.basename(path)} (Sheet: {sheet})")
                df = pd.read_excel(path, sheet_name=sheet)
            
            elif file_type == "CSV":
                delim = config_data.get("delimiter", ",")
                self._log(f"Loading CSV: {os.path.basename(path)} (Delimiter: '{delim}')")
                df = pd.read_csv(path, delimiter=delim)

            elif file_type == "JSON":
                orient = config_data.get("json_orient", "records")
                self._log(f"Loading JSON: {os.path.basename(path)} (Orient: {orient})")
                df = pd.read_json(path, orient=orient)

            elif file_type == "Text (Delimited)":
                delim = config_data.get("delimiter", r"\t")
                if delim == r'\t': delim = '\t'
                self._log(f"Loading Text: {os.path.basename(path)} (Delimiter: '{delim}')")
                df = pd.read_csv(path, delimiter=delim)
            
            if df is None:
                raise ValueError("File type not supported or loading failed.")

            selected_cols = config_data.get("selected_columns")
            if selected_cols:
                existing_cols = [col for col in selected_cols if col in df.columns]
                missing_cols = set(selected_cols) - set(existing_cols)
                if missing_cols:
                    self._log(f"Warning: Could not find columns: {missing_cols}")
                if not existing_cols:
                    raise ValueError("No valid columns were selected or found.")
                df = df[existing_cols]
            
            self._log(f"Successfully loaded DataFrame with {len(df)} rows and {len(df.columns)} columns.")
            return df
        except Exception as e:
            self._log(f"FATAL ERROR during data load: {e}")
            raise

    def _execute_email_read(self, config_data: dict) -> pd.DataFrame:
        """Internal logic for reading emails from Outlook."""
        self._log("Connecting to Outlook and processing emails...")
        try:
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            
            selected_folder_ids = config_data.get("selected_folder_ids", [])
            start_time = datetime.fromisoformat(config_data.get("start_time"))
            
            if config_data.get("use_current_end_time", False):
                end_time = datetime.now()
                self._log("Using current datetime as end time.")
            else:
                end_time = datetime.fromisoformat(config_data.get("end_time"))
                self._log(f"Using fixed end time: {end_time}")
            
            read_status_filter = config_data.get("read_status_filter", "All Emails")
            save_attachments = config_data.get("save_attachments", False)
            save_path = config_data.get("save_path", "")
            if save_attachments:
                os.makedirs(save_path, exist_ok=True)
                self._log(f"Will save attachments to: {save_path}")

            def prep_filter(f):
                f = config_data.get(f, "").strip()
                return f if f == "" or "*" in f or "?" in f else f"*{f}*"

            sender_f = prep_filter("sender_filter").lower()
            subject_f = prep_filter("subject_filter").lower()
            body_f = prep_filter("body_filter").lower()
            attach_f = prep_filter("attachment_filter").lower()

            found_emails = []
            self._log(f"Searching {len(selected_folder_ids)} folders from {start_time} to {end_time}...")
            
            total_processed = 0
            total_found = 0

            for folder_id in selected_folder_ids:
                try:
                    folder = outlook.GetFolderFromID(folder_id)
                except Exception as e:
                    self._log(f"Warning: Could not access folder with ID {folder_id}. Skipping. Error: {e}")
                    continue
                
                self._log(f"Searching folder: {folder.Name}")
                
                filters = []
                filters.append(f"[ReceivedTime] >= '{start_time.strftime('%Y-%m-%d %H:%M')}'")
                filters.append(f"[ReceivedTime] <= '{end_time.strftime('%Y-%m-%d %H:%M')}'")
                
                if read_status_filter == "Unread Only":
                    filters.append("[Unread] = true")
                elif read_status_filter == "Read Only":
                    filters.append("[Unread] = false")
                
                combined_filter = " AND ".join(f"({f})" for f in filters)
                
                try:
                    items = folder.Items.Restrict(combined_filter)
                except Exception:
                    self._log(f"Could not apply combined filter to {folder.Name}, filtering manually (slower).")
                    items = folder.Items

                for item in items:
                    total_processed += 1
                    
                    if item.Class != 43: # olMail
                        continue
                        
                    try:
                        item_received_time = item.ReceivedTime
                        if item_received_time.timestamp() < start_time.timestamp() or \
                           item_received_time.timestamp() > end_time.timestamp():
                            continue
                        
                        if read_status_filter == "Unread Only" and item.UnRead == False:
                            continue
                        if read_status_filter == "Read Only" and item.UnRead == True:
                            continue

                        sender = getattr(item, 'SenderEmailAddress', '').lower()
                        if sender_f and not fnmatch.fnmatch(sender, sender_f):
                            continue
                            
                        subject = getattr(item, 'Subject', '').lower()
                        if subject_f and not fnmatch.fnmatch(subject, subject_f):
                            continue

                        body = getattr(item, 'Body', '').lower()
                        if body_f and not fnmatch.fnmatch(body, body_f):
                            continue
                            
                        attachment_names = [att.FileName.lower() for att in item.Attachments]
                        if attach_f:
                            if not any(fnmatch.fnmatch(name, attach_f) for name in attachment_names):
                                continue
                        
                        if save_attachments:
                            for att in item.Attachments:
                                try:
                                    att_filename = att.FileName
                                    save_full_path = os.path.join(save_path, att_filename)
                                    att.SaveAsFile(save_full_path)
                                except Exception as e:
                                    self._log(f"Error saving attachment {att.FileName}: {e}")
                        
                        recipients_list = []
                        try: recipients_list.append(getattr(item, 'To', ''))
                        except Exception: pass
                        try: recipients_list.append(getattr(item, 'CC', ''))
                        except Exception: pass
                        try: recipients_list.append(getattr(item, 'BCC', ''))
                        except Exception: pass
                        
                        recipients = "; ".join(filter(None, recipients_list))
                        status = "Unread" if item.UnRead else "Read"
                        
                        received_dt = datetime(
                            item_received_time.year, item_received_time.month, item_received_time.day,
                            item_received_time.hour, item_received_time.minute, item_received_time.second
                        )
                        
                        total_found += 1
                        found_emails.append({
                            "Sender": sender,
                            "Recipients": recipients,
                            "Subject": item.Subject,
                            "Status": status,
                            "ReceivedTime": received_dt,
                            "Body": item.Body,
                            "AttachmentCount": len(attachment_names),
                            "AttachmentNames": ", ".join(attachment_names),
                            "EntryID": item.EntryID
                        })

                    except Exception as e:
                        self._log(f"Error processing item '{getattr(item, 'Subject', 'N/A')}': {e}")
                        
            self._log(f"Search complete. Processed {total_processed} items, found {total_found} matching emails.")

            df_columns = ["Sender", "Recipients", "Subject", "Status", "ReceivedTime", "Body", "AttachmentCount", "AttachmentNames", "EntryID"]
            if not found_emails:
                self._log("No matching emails found. Returning empty DataFrame.")
                return pd.DataFrame(columns=df_columns)

            df = pd.DataFrame(found_emails, columns=df_columns)
            return df

        except Exception as e:
            self._log(f"FATAL ERROR during Outlook read: {e}")
            raise
        finally:
            pythoncom.CoUninitialize()