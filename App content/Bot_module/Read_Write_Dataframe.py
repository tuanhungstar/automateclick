# File: Bot_module/file_module.py

import sys
from typing import Optional, List, Dict, Any
import pandas as pd
import os
import math

# --- PyQt6 Imports ---
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QDialogButtonBox,
    QComboBox, QWidget, QGroupBox, QMessageBox, QLabel,
    QCheckBox, QApplication, QFileDialog, QTableView, QListWidget, QListWidgetItem,
    QHBoxLayout, QRadioButton, QSpinBox, QProgressBar
)
from PyQt6.QtGui import QStandardItemModel, QStandardItem
from PyQt6.QtCore import Qt, QThread, pyqtSignal


# --- Main App Imports (Fallback for standalone testing) ---
try:
    from my_lib.shared_context import ExecutionContext
except ImportError:
    print("Warning: Could not import main app libraries. Using fallbacks.")
    class ExecutionContext:
        def add_log(self, message: str): print(message)
        def get_variable(self, name: str):
            print(f"Fallback: Getting variable '{name}'")
            return None # Fallback behavior

#
# --- HELPER: Worker thread for loading file preview ---
#
class _FilePreviewLoaderThread(QThread):
    preview_ready = pyqtSignal(pd.DataFrame)
    error_occurred = pyqtSignal(str)

    def __init__(self, config: dict):
        super().__init__()
        self.config = config

    def run(self):
        try:
            file_path = self.config['file_path']
            file_type = self.config['file_type']
            
            df_preview = pd.DataFrame()

            if file_type == 'Excel':
                sheet_name = self.config['sheet_name']
                excel_file = pd.ExcelFile(file_path)
                all_sheets = excel_file.sheet_names

                # Allow user to specify sheet by name or index
                target_sheet = None
                if str(sheet_name).isdigit():
                    sheet_index = int(sheet_name)
                    if 0 <= sheet_index < len(all_sheets):
                        target_sheet = all_sheets[sheet_index]
                if target_sheet is None:
                    target_sheet = sheet_name if sheet_name in all_sheets else all_sheets[0]

                df_preview = pd.read_excel(excel_file, sheet_name=target_sheet, nrows=50)

            elif file_type == 'CSV':
                df_preview = pd.read_csv(file_path, nrows=50)
                
            elif file_type == 'TXT':
                # For TXT, we'll assume a simple one-column file for preview
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    lines = [next(f) for _ in range(50) if f]
                df_preview = pd.DataFrame(lines, columns=['Content'])

            self.preview_ready.emit(df_preview)
        except Exception as e:
            self.error_occurred.emit(f"Failed to load preview. Error: {e}")

#
# --- HELPER: The GUI Dialog for the File Loader ---
#
class _FileLoaderDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None,
                 initial_variable: Optional[str] = None):
        super().__init__(parent)
        self.setWindowTitle("File Loader")
        self.setMinimumSize(800, 800)
        self.global_variables = global_variables
        self.preview_df = None

        main_layout = QVBoxLayout(self)

        # 1. File Source
        source_group = QGroupBox("File Source")
        source_layout = QFormLayout(source_group)
        self.file_type_combo = QComboBox(); self.file_type_combo.addItems(['Excel', 'CSV', 'TXT'])
        self.file_path_edit = QLineEdit(); self.file_path_edit.setReadOnly(True)
        browse_button = QPushButton("Browse...")
        path_layout = QHBoxLayout(); path_layout.addWidget(self.file_path_edit); path_layout.addWidget(browse_button)
        source_layout.addRow("File Type:", self.file_type_combo)
        source_layout.addRow("File Path:", path_layout)
        main_layout.addWidget(source_group)

        # 2. File Options (dynamically shown)
        self.options_group = QGroupBox("File Options")
        options_layout = QFormLayout(self.options_group)
        self.sheet_name_edit = QLineEdit("0")
        self.sheet_name_label = QLabel("Sheet Name (or index):")
        options_layout.addRow(self.sheet_name_label, self.sheet_name_edit)
        main_layout.addWidget(self.options_group)

        # 3. Data Preview
        preview_group = QGroupBox("Data Preview (First 50 Rows)")
        preview_layout = QVBoxLayout(preview_group)
        self.load_preview_button = QPushButton("Load Preview")
        self.preview_table = QTableView(); self.preview_table.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
        preview_layout.addWidget(self.load_preview_button); preview_layout.addWidget(self.preview_table)
        main_layout.addWidget(preview_group)

        # 4. Column Selection
        column_group = QGroupBox("Column Selection")
        column_layout = QVBoxLayout(column_group)
        self.column_list_widget = QListWidget(); self.column_list_widget.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        select_buttons_layout = QHBoxLayout()
        select_all_button = QPushButton("Select All"); deselect_all_button = QPushButton("Deselect All")
        select_buttons_layout.addWidget(select_all_button); select_buttons_layout.addWidget(deselect_all_button)
        column_layout.addWidget(self.column_list_widget); column_layout.addLayout(select_buttons_layout)
        main_layout.addWidget(column_group)

        # 5. Assign Results
        assign_group = QGroupBox("Assign Results to Variable"); assign_layout = QFormLayout(assign_group)
        self.assign_results_check = QCheckBox("Assign results (DataFrame) to a variable")
        self.new_var_radio = QRadioButton("New Variable Name:"); self.new_var_input = QLineEdit("new_data")
        self.existing_var_radio = QRadioButton("Existing Variable:"); self.existing_var_combo = QComboBox(); self.existing_var_combo.addItems(["-- Select --"] + global_variables)
        assign_layout.addRow(self.assign_results_check); assign_layout.addRow(self.new_var_radio, self.new_var_input); assign_layout.addRow(self.existing_var_radio, self.existing_var_combo)
        main_layout.addWidget(assign_group)
        
        # Dialog Buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # Connections
        browse_button.clicked.connect(self._browse_for_file)
        self.file_type_combo.currentTextChanged.connect(self._on_file_type_changed)
        self.load_preview_button.clicked.connect(self._load_preview)
        select_all_button.clicked.connect(self.column_list_widget.selectAll)
        deselect_all_button.clicked.connect(self.column_list_widget.clearSelection)
        self.assign_results_check.toggled.connect(self._toggle_assignment_widgets)
        self.button_box.accepted.connect(self.accept); self.button_box.rejected.connect(self.reject)

        self._on_file_type_changed(self.file_type_combo.currentText())
        if initial_config: self._populate_from_initial_config(initial_config, initial_variable)
        self._toggle_assignment_widgets(True)
        self.assign_results_check.setChecked(True)
        self.new_var_radio.setChecked(True)

    def _browse_for_file(self):
        file_type = self.file_type_combo.currentText()
        filters = "All Files (*)"
        if file_type == 'Excel': filters = "Excel Files (*.xlsx *.xls);;All Files (*)"
        elif file_type == 'CSV': filters = "CSV Files (*.csv);;All Files (*)"
        elif file_type == 'TXT': filters = "Text Files (*.txt);;All Files (*)"
        
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", filters)
        if file_path:
            self.file_path_edit.setText(file_path)
            self.preview_table.setModel(None) # Clear preview on new file
            self.column_list_widget.clear()

    def _on_file_type_changed(self, file_type: str):
        is_excel = (file_type == 'Excel')
        self.sheet_name_edit.setVisible(is_excel)
        self.sheet_name_label.setVisible(is_excel)
        
    def _load_preview(self):
        config = self._get_preview_config()
        if not config: return
        self.load_preview_button.setText("Loading..."); self.load_preview_button.setEnabled(False)
        self.preview_loader_thread = _FilePreviewLoaderThread(config)
        self.preview_loader_thread.preview_ready.connect(self._on_preview_loaded)
        self.preview_loader_thread.error_occurred.connect(lambda e: (QMessageBox.critical(self, "Error", e), self.load_preview_button.setText("Load Preview"), self.load_preview_button.setEnabled(True)))
        self.preview_loader_thread.start()

    def _on_preview_loaded(self, df: pd.DataFrame):
        self.preview_df = df
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(df.columns)
        for i, row in df.iterrows():
            items = [QStandardItem(str(val)) for val in row]
            model.appendRow(items)
        self.preview_table.setModel(model)
        
        self.column_list_widget.clear()
        for col in df.columns:
            item = QListWidgetItem(col)
            # Use checkable items instead of selection mode for clarity
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked)
            self.column_list_widget.addItem(item)
            
        self.load_preview_button.setText("Reload Preview"); self.load_preview_button.setEnabled(True)

    def _get_preview_config(self) -> Optional[Dict[str, Any]]:
        file_path = self.file_path_edit.text()
        if not file_path:
            QMessageBox.warning(self, "Input Error", "Please select a file path first."); return None
        return {
            'file_path': file_path,
            'file_type': self.file_type_combo.currentText(),
            'sheet_name': self.sheet_name_edit.text()
        }

    def _toggle_assignment_widgets(self, checked):
        self.new_var_radio.setEnabled(checked); self.new_var_input.setEnabled(checked)
        self.existing_var_radio.setEnabled(checked); self.existing_var_combo.setEnabled(checked)

    def _populate_from_initial_config(self, config, variable):
        self.file_type_combo.setCurrentText(config.get("file_type", "Excel"))
        self.file_path_edit.setText(config.get("file_path", ""))
        self.sheet_name_edit.setText(config.get("sheet_name", "0"))
        
        if config.get("file_path"):
            self._load_preview()
        
        if variable:
            self.assign_results_check.setChecked(True)
            if variable in self.global_variables: self.existing_var_radio.setChecked(True); self.existing_var_combo.setCurrentText(variable)
            else: self.new_var_radio.setChecked(True); self.new_var_input.setText(variable)

    def get_executor_method_name(self) -> str: return "_load_file_data"
    def get_config_data(self) -> Optional[Dict[str, Any]]:
        file_path = self.file_path_edit.text()
        if not file_path: QMessageBox.warning(self, "Input Error", "Please select a file path."); return None
        
        selected_columns = []
        for i in range(self.column_list_widget.count()):
            item = self.column_list_widget.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                selected_columns.append(item.text())
        
        if self.column_list_widget.count() > 0 and not selected_columns:
            QMessageBox.warning(self, "Input Error", "Please select at least one column to load."); return None

        return {
            "file_path": file_path,
            "file_type": self.file_type_combo.currentText(),
            "sheet_name": self.sheet_name_edit.text(),
            "selected_columns": selected_columns
        }
    def get_assignment_variable(self) -> Optional[str]:
        if not self.assign_results_check.isChecked(): return None
        if self.new_var_radio.isChecked():
            var_name = self.new_var_input.text().strip()
            if not var_name: QMessageBox.warning(self, "Input Error", "New variable name cannot be empty."); return None
            return var_name
        else:
            var_name = self.existing_var_combo.currentText()
            if var_name == "-- Select --": QMessageBox.warning(self, "Input Error", "Please select an existing variable."); return None
            return var_name

#
# --- The Public-Facing Module Class for File Loading ---
#
class File_Reader:
    def __init__(self, context: Optional[ExecutionContext] = None): self.context = context
    def _log(self, m: str): (self.context.add_log(m) if self.context else print(m))

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        self._log("Opening File Loader configuration...")
        return _FileLoaderDialog(global_variables, parent_window, **kwargs)

    def _load_file_data(self, context: ExecutionContext, config_data: dict) -> pd.DataFrame:
        self.context = context
        file_path, file_type = config_data["file_path"], config_data["file_type"]
        selected_columns = config_data.get("selected_columns")
        
        self._log(f"Loading data from {file_type} file: {os.path.basename(file_path)}")
        df = pd.DataFrame()
        use_cols = selected_columns if selected_columns else None

        try:
            if file_type == 'Excel':
                sheet_name_str = config_data.get('sheet_name', "0")
                excel_file = pd.ExcelFile(file_path)
                all_sheets = excel_file.sheet_names
                target_sheet = None

                if sheet_name_str.isdigit():
                    sheet_index = int(sheet_name_str)
                    if 0 <= sheet_index < len(all_sheets):
                        target_sheet = all_sheets[sheet_index]
                if target_sheet is None:
                    target_sheet = sheet_name_str if sheet_name_str in all_sheets else all_sheets[0]
                
                self._log(f"Reading from sheet: '{target_sheet}'")
                df = pd.read_excel(excel_file, sheet_name=target_sheet, usecols=use_cols)
            
            elif file_type == 'CSV':
                df = pd.read_csv(file_path, usecols=use_cols)
            
            elif file_type == 'TXT':
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    lines = f.readlines()
                col_name = use_cols[0] if use_cols else 'Content'
                df = pd.DataFrame(lines, columns=[col_name])
            
            self._log(f"Successfully loaded {len(df)} rows and {len(df.columns)} columns.")
            return df
        except Exception as e:
            self._log(f"FATAL ERROR during file loading: {e}"); raise
            
class _FileWriterDialog(QDialog):
    def __init__(self, df_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None):
        super().__init__(parent)
        self.setWindowTitle("File Writer")
        self.setMinimumWidth(500)
        self.df_variables = df_variables
        
        main_layout = QVBoxLayout(self)

        # 1. File Source & Destination
        source_group = QGroupBox("File Source & Destination")
        source_layout = QFormLayout(source_group)
        
        self.df_var_combo = QComboBox(); self.df_var_combo.addItems(["-- Select DataFrame --"] + self.df_variables)
        self.file_type_combo = QComboBox(); self.file_type_combo.addItems(['Excel', 'CSV', 'TXT'])
        self.file_path_edit = QLineEdit(); self.file_path_edit.setPlaceholderText("e.g., C:\\data\\output.xlsx")
        browse_button = QPushButton("Browse...")
        path_layout = QHBoxLayout(); path_layout.addWidget(self.file_path_edit); path_layout.addWidget(browse_button)
        
        source_layout.addRow("DataFrame to Save:", self.df_var_combo)
        source_layout.addRow("Save as File Type:", self.file_type_combo)
        source_layout.addRow("Save to File Path:", path_layout)
        main_layout.addWidget(source_group)

        # 2. File Options
        self.options_group = QGroupBox("File Options")
        options_layout = QFormLayout(self.options_group)
        self.sheet_name_edit = QLineEdit("Sheet1")
        self.sheet_name_label = QLabel("Sheet Name:")
        self.include_index_check = QCheckBox("Include DataFrame index in file")
        options_layout.addRow(self.sheet_name_label, self.sheet_name_edit)
        options_layout.addRow(self.include_index_check)
        main_layout.addWidget(self.options_group)

        # Dialog Buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # Connections
        browse_button.clicked.connect(self._browse_for_save_path)
        self.file_type_combo.currentTextChanged.connect(self._on_file_type_changed)
        self.button_box.accepted.connect(self.accept); self.button_box.rejected.connect(self.reject)

        self._on_file_type_changed(self.file_type_combo.currentText())
        if initial_config: self._populate_from_initial_config(initial_config)

    def _browse_for_save_path(self):
        file_type = self.file_type_combo.currentText()
        filters = "All Files (*)"
        if file_type == 'Excel': filters = "Excel Files (*.xlsx);;All Files (*)"
        elif file_type == 'CSV': filters = "CSV Files (*.csv);;All Files (*)"
        elif file_type == 'TXT': filters = "Text Files (*.txt);;All Files (*)"
        
        file_path, _ = QFileDialog.getSaveFileName(self, "Save File As", "", filters)
        if file_path:
            self.file_path_edit.setText(file_path)

    def _on_file_type_changed(self, file_type: str):
        is_excel = (file_type == 'Excel')
        self.sheet_name_edit.setVisible(is_excel)
        self.sheet_name_label.setVisible(is_excel)
        
    def _populate_from_initial_config(self, config):
        self.df_var_combo.setCurrentText(config.get("dataframe_var", "-- Select DataFrame --"))
        self.file_type_combo.setCurrentText(config.get("file_type", "Excel"))
        self.file_path_edit.setText(config.get("file_path", ""))
        self.sheet_name_edit.setText(config.get("sheet_name", "Sheet1"))
        self.include_index_check.setChecked(config.get("include_index", False))

    def get_executor_method_name(self) -> str: return "_save_file_data"
    def get_assignment_variable(self) -> Optional[str]: return None 

    def get_config_data(self) -> Optional[Dict[str, Any]]:
        df_var = self.df_var_combo.currentText()
        if df_var == "-- Select DataFrame --":
            QMessageBox.warning(self, "Input Error", "Please select a DataFrame to save."); return None
            
        file_path = self.file_path_edit.text().strip()
        if not file_path:
            QMessageBox.warning(self, "Input Error", "Please specify a file path to save to."); return None

        return {
            "dataframe_var": df_var,
            "file_path": file_path,
            "file_type": self.file_type_combo.currentText(),
            "sheet_name": self.sheet_name_edit.text(),
            "include_index": self.include_index_check.isChecked()
        }

#
# --- The Public-Facing Module Class for File Writing ---
#
class File_Writer:
    def __init__(self, context: Optional[ExecutionContext] = None): self.context = context
    def _log(self, m: str): (self.context.add_log(m) if self.context else print(m))

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        self._log("Opening File Writer configuration...")
        df_variables = []
        # A more robust way to get dataframe variables if the parent provides them
        if hasattr(parent_window, 'get_dataframe_variables'):
             df_variables = parent_window.get_dataframe_variables()
        else: # Fallback for testing
             df_variables = global_variables

        initial_config = kwargs.get("initial_config")

        return _FileWriterDialog(
            df_variables=df_variables, 
            parent=parent_window, 
            initial_config=initial_config
        )


    def _save_file_data(self, context: ExecutionContext, config_data: dict):
        self.context = context
        df_var, file_path, file_type = config_data["dataframe_var"], config_data["file_path"], config_data["file_type"]
        
        df_to_save = self.context.get_variable(df_var)
        if not isinstance(df_to_save, pd.DataFrame):
            raise TypeError(f"Variable '{df_var}' is not a pandas DataFrame.")
        
        try:
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
        except Exception: pass

        self._log(f"Saving DataFrame '{df_var}' to {file_type} file: {os.path.basename(file_path)}")
        
        try:
            if file_type == 'Excel':
                sheet_name = config_data.get('sheet_name', 'Sheet1')
                if not sheet_name: sheet_name = 'Sheet1'
                df_to_save.to_excel(
                    file_path, 
                    sheet_name=sheet_name, 
                    index=config_data.get("include_index", False)
                )
            
            elif file_type == 'CSV':
                df_to_save.to_csv(
                    file_path,
                    index=config_data.get("include_index", False)
                )
            
            elif file_type == 'TXT':
                df_string = df_to_save.to_string(index=config_data.get("include_index", False))
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(df_string)
            
            self._log(f"Successfully saved {len(df_to_save)} rows to {file_path}.")
        except Exception as e:
            self._log(f"FATAL ERROR during file write: {e}"); raise

################################################################################
# --- NEWLY ADDED CODE STARTS HERE ---
################################################################################

#
# --- [NEW] HELPER: Worker thread for splitting files ---
#
class _FileSplitterThread(QThread):
    """Handles the long-running file splitting task in a separate thread."""
    progress_update = pyqtSignal(int, str)  # (percentage, message)
    finished_signal = pyqtSignal(str)       # (final_message)
    error_signal = pyqtSignal(str)          # (error_message)

    def __init__(self, config: dict, context: Optional[ExecutionContext] = None):
        super().__init__()
        self.config = config
        self.context = context
        self.is_interrupted = False

    def _log(self, message: str):
        """Log message using the execution context if available."""
        if self.context:
            self.context.add_log(message)
        else:
            print(message)

    def stop(self):
        self.is_interrupted = True

    def run(self):
        try:
            file_path = self.config['file_path']
            file_type = self.config['file_type']
            chunk_size = self.config['chunk_size']
            target_folder = self.config['target_folder']
            sheet_name = self.config.get('sheet_name')
            base_filename, original_ext = os.path.splitext(os.path.basename(file_path))

            self._log(f"Preparing to split '{base_filename}{original_ext}' into chunks of {chunk_size} rows.")
            os.makedirs(target_folder, exist_ok=True)

            if file_type == 'Excel':
                self._split_excel(file_path, sheet_name, chunk_size, target_folder, base_filename)
            elif file_type == 'CSV':
                self._split_csv(file_path, chunk_size, target_folder, base_filename)
            elif file_type == 'TXT':
                self._split_txt(file_path, chunk_size, target_folder, base_filename)

        except Exception as e:
            self.error_signal.emit(f"An error occurred: {e}")
            self._log(f"ERROR: {e}")

    def _split_excel(self, file_path, sheet_name, chunk_size, target_folder, base_filename):
        self._log(f"Loading Excel sheet: '{sheet_name}'. This may take time for large files.")
        self.progress_update.emit(0, f"Loading sheet '{sheet_name}'...")
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        total_rows = len(df)
        if total_rows == 0:
            self.finished_signal.emit("Sheet is empty. No files created.")
            return

        total_chunks = math.ceil(total_rows / chunk_size)
        self._log(f"Splitting {total_rows} rows into {total_chunks} files.")

        for i in range(total_chunks):
            if self.is_interrupted:
                self._log("Splitting process was cancelled.")
                return

            start_row = i * chunk_size
            end_row = start_row + chunk_size
            chunk_df = df.iloc[start_row:end_row]
            
            output_filename = f"{base_filename}_{sheet_name}_part_{i+1:04d}.xlsx"
            output_path = os.path.join(target_folder, output_filename)
            
            chunk_df.to_excel(output_path, index=False, sheet_name=sheet_name)
            
            progress = int(((i + 1) / total_chunks) * 100)
            self.progress_update.emit(progress, f"Saved chunk {i+1}/{total_chunks}")
        
        self.finished_signal.emit(f"Successfully split into {total_chunks} files.")

    def _split_csv(self, file_path, chunk_size, target_folder, base_filename):
        self._log("Processing CSV file in chunks (memory efficient).")
        chunk_iterator = pd.read_csv(file_path, chunksize=chunk_size)
        
        for i, chunk_df in enumerate(chunk_iterator):
            if self.is_interrupted:
                self._log("Splitting process was cancelled.")
                return

            output_filename = f"{base_filename}_part_{i+1:04d}.csv"
            output_path = os.path.join(target_folder, output_filename)
            chunk_df.to_csv(output_path, index=False)
            self.progress_update.emit(0, f"Saved chunk {i+1} with {len(chunk_df)} rows")
        
        self.finished_signal.emit(f"Successfully finished splitting CSV.")

    def _split_txt(self, file_path, chunk_size, target_folder, base_filename):
        self._log("Reading entire TXT file to split by lines.")
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()
        
        total_lines = len(lines)
        if total_lines == 0:
            self.finished_signal.emit("File is empty. No files created.")
            return

        total_chunks = math.ceil(total_lines / chunk_size)
        self._log(f"Splitting {total_lines} lines into {total_chunks} files.")

        for i in range(total_chunks):
            if self.is_interrupted:
                self._log("Splitting process was cancelled.")
                return

            start_line = i * chunk_size
            end_line = start_line + chunk_size
            chunk_lines = lines[start_line:end_line]
            
            output_filename = f"{base_filename}_part_{i+1:04d}.txt"
            output_path = os.path.join(target_folder, output_filename)

            with open(output_path, 'w', encoding='utf-8') as out_f:
                out_f.writelines(chunk_lines)
            
            progress = int(((i + 1) / total_chunks) * 100)
            self.progress_update.emit(progress, f"Saved chunk {i+1}/{total_chunks}")
        
        self.finished_signal.emit(f"Successfully split into {total_chunks} files.")

#
# --- [NEW] HELPER: The GUI Dialog for the File Splitter ---
#
class _FileSplitterDialog(QDialog):
    def __init__(self, parent: Optional[QWidget] = None, initial_config: Optional[Dict[str, Any]] = None):
        super().__init__(parent)
        self.setWindowTitle("File Splitter")
        self.setMinimumWidth(600)

        main_layout = QVBoxLayout(self)

        # 1. File Source
        source_group = QGroupBox("File to Split")
        source_layout = QFormLayout(source_group)
        self.file_type_combo = QComboBox(); self.file_type_combo.addItems(['Excel', 'CSV', 'TXT'])
        self.file_path_edit = QLineEdit(); self.file_path_edit.setReadOnly(True)
        browse_file_button = QPushButton("Browse File...")
        path_layout = QHBoxLayout(); path_layout.addWidget(self.file_path_edit); path_layout.addWidget(browse_file_button)
        source_layout.addRow("File Type:", self.file_type_combo)
        source_layout.addRow("File Path:", path_layout)
        main_layout.addWidget(source_group)

        # 2. Excel Options
        self.excel_options_group = QGroupBox("Excel Options")
        excel_layout = QFormLayout(self.excel_options_group)
        self.sheet_list_combo = QComboBox()
        self.load_sheets_button = QPushButton("Load Sheets from File")
        excel_layout.addRow("Sheet to Split:", self.sheet_list_combo)
        excel_layout.addRow(self.load_sheets_button)
        main_layout.addWidget(self.excel_options_group)

        # 3. Splitting Options
        split_group = QGroupBox("Splitting Options")
        split_layout = QFormLayout(split_group)
        self.chunk_size_spinbox = QSpinBox(); self.chunk_size_spinbox.setRange(1, 10_000_000); self.chunk_size_spinbox.setValue(10000)
        self.chunk_size_spinbox.setToolTip("The maximum number of rows each smaller file will contain.")
        self.target_folder_edit = QLineEdit(); self.target_folder_edit.setReadOnly(True)
        browse_folder_button = QPushButton("Browse Folder...")
        folder_layout = QHBoxLayout(); folder_layout.addWidget(self.target_folder_edit); folder_layout.addWidget(browse_folder_button)
        split_layout.addRow("Rows per File (Chunk Size):", self.chunk_size_spinbox)
        split_layout.addRow("Target Folder:", folder_layout)
        main_layout.addWidget(split_group)
        
        # Dialog Buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # Connections
        browse_file_button.clicked.connect(self._browse_for_file)
        browse_folder_button.clicked.connect(self._browse_for_folder)
        self.load_sheets_button.clicked.connect(self._load_excel_sheets)
        self.file_type_combo.currentTextChanged.connect(self._on_file_type_changed)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        self._on_file_type_changed(self.file_type_combo.currentText())
        if initial_config: self._populate_from_initial_config(initial_config)

    def _browse_for_file(self):
        file_type = self.file_type_combo.currentText()
        filters = {"Excel": "Excel Files (*.xlsx *.xls)", "CSV": "CSV Files (*.csv)", "TXT": "Text Files (*.txt)"}
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File to Split", "", filters.get(file_type, "All Files (*)"))
        if file_path:
            self.file_path_edit.setText(file_path)
            self.sheet_list_combo.clear() # Clear sheets from previous file
            if self.file_type_combo.currentText() == 'Excel':
                self._load_excel_sheets()

    def _browse_for_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Target Folder")
        if folder_path: self.target_folder_edit.setText(folder_path)

    def _on_file_type_changed(self, file_type: str):
        self.excel_options_group.setVisible(file_type == 'Excel')
        # Automatically select a default folder if not set
        if self.file_path_edit.text() and not self.target_folder_edit.text():
            default_folder = os.path.dirname(self.file_path_edit.text())
            self.target_folder_edit.setText(default_folder)

    def _load_excel_sheets(self):
        file_path = self.file_path_edit.text()
        if not file_path:
            QMessageBox.warning(self, "Input Missing", "Please select an Excel file first.")
            return
        try:
            self.load_sheets_button.setText("Loading..."); self.load_sheets_button.setEnabled(False)
            QApplication.processEvents() # Ensure UI updates
            xls = pd.ExcelFile(file_path)
            self.sheet_list_combo.clear()
            self.sheet_list_combo.addItems(xls.sheet_names)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not read Excel file: {e}")
        finally:
            self.load_sheets_button.setText("Load Sheets from File")
            self.load_sheets_button.setEnabled(True)

    def _populate_from_initial_config(self, config: dict):
        self.file_type_combo.setCurrentText(config.get("file_type", "Excel"))
        self.file_path_edit.setText(config.get("file_path", ""))
        self.chunk_size_spinbox.setValue(config.get("chunk_size", 10000))
        self.target_folder_edit.setText(config.get("target_folder", ""))
        
        if config.get("file_path") and config.get("file_type") == "Excel":
            # Post-load sheet names and set the previously selected one
            QApplication.processEvents()
            self._load_excel_sheets()
            self.sheet_list_combo.setCurrentText(config.get("sheet_name", ""))
            
    def get_executor_method_name(self) -> str: return "_split_file"
    def get_assignment_variable(self) -> Optional[str]: return None # This action doesn't create a new variable

    def get_config_data(self) -> Optional[Dict[str, Any]]:
        file_path = self.file_path_edit.text()
        if not file_path:
            QMessageBox.warning(self, "Input Error", "Please select a file to split."); return None
            
        target_folder = self.target_folder_edit.text()
        if not target_folder:
            QMessageBox.warning(self, "Input Error", "Please select a target folder."); return None
        
        file_type = self.file_type_combo.currentText()
        sheet_name = ""
        if file_type == 'Excel':
            sheet_name = self.sheet_list_combo.currentText()
            if not sheet_name:
                QMessageBox.warning(self, "Input Error", "Please select a sheet to split for the Excel file."); return None

        return {
            "file_path": file_path,
            "file_type": file_type,
            "sheet_name": sheet_name,
            "chunk_size": self.chunk_size_spinbox.value(),
            "target_folder": target_folder,
        }

#
# --- [NEW] The Public-Facing Module Class for File Splitting ---
#
class Split_file:
    """A module to split a large file (Excel, CSV, TXT) into multiple smaller files."""
    def __init__(self, context: Optional[ExecutionContext] = None):
        self.context = context
        self.worker: Optional[_FileSplitterThread] = None

    def _log(self, message: str):
        if self.context:
            self.context.add_log(message)
        else:
            print(message)
    
    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        """Opens the configuration dialog for the file splitter."""
        self._log("Opening File Splitter configuration...")
        initial_config = kwargs.get("initial_config")
        # This dialog doesn't use global variables, so we pass an empty list.
        return _FileSplitterDialog(parent=parent_window, initial_config=initial_config)

    def _split_file(self, context: ExecutionContext, config_data: dict):
        """
        Executes the file splitting process using a worker thread to keep the UI responsive.
        This method will block until the thread is finished.
        """
        self.context = context
        self.worker = _FileSplitterThread(config_data, context)
        
        final_message = ""
        error_message = ""

        def on_finished(msg):
            nonlocal final_message
            final_message = msg
            self._log(f"SUCCESS: {msg}")

        def on_error(msg):
            nonlocal error_message
            error_message = msg
            self._log(f"ERROR from worker: {msg}")

        # Connect signals
        # If the context has a progress bar, connect it. Otherwise, log progress.
        if hasattr(context, 'update_progress_bar'):
            self.worker.progress_update.connect(context.update_progress_bar)
        else:
            self.worker.progress_update.connect(lambda p, msg: self._log(f"Progress ({p}%): {msg}"))

        self.worker.finished_signal.connect(on_finished)
        self.worker.error_signal.connect(on_error)
        
        self.worker.start()
        
        # Block execution of this method until the thread completes its run() method.
        # This is crucial for sequential task execution in your main application.
        self.worker.wait()

        # After the worker is done, check if it produced an error and raise it
        # to halt the main execution flow if necessary.
        if error_message:
            raise Exception(f"File splitting failed: {error_message}")
        
        self._log(f"File splitting process finished. {final_message}")
      
class _FileMergerThread(QThread):
    """Handles the long-running file merging task in a separate thread."""
    progress_update = pyqtSignal(int, str)  # (percentage, message)
    finished_signal = pyqtSignal(pd.DataFrame)  # (merged_dataframe)
    error_signal = pyqtSignal(str)          # (error_message)

    def __init__(self, config: dict, context: Optional[ExecutionContext] = None):
        super().__init__()
        self.config = config
        self.context = context

    def _log(self, message: str):
        if self.context: self.context.add_log(message)
        else: print(message)

    def run(self):
        try:
            folder_path = self.config['folder_path']
            file_type = self.config['file_type']
            template_file = self.config['template_file']
            mismatch_strategy = self.config['mismatch_strategy']
            keyword = self.config.get('keyword', '').lower()

            ext_map = {'Excel': ('.xlsx', '.xls'), 'CSV': ('.csv',), 'TXT': ('.txt',)}
            extensions = ext_map.get(file_type)

            self._log(f"Scanning '{folder_path}' for '{file_type}' files...")
            if keyword:
                self._log(f"Filtering by keyword: '{keyword}'")

            files_to_merge = sorted([
                f for f in os.listdir(folder_path)
                if os.path.isfile(os.path.join(folder_path, f))
                and f.endswith(extensions)
                and (keyword in f.lower()) # Case-insensitive keyword filtering
            ])

            if not files_to_merge:
                # Fail early if no files match the criteria at all.
                raise ValueError("No matching files found in the folder with the specified criteria.")

            self._log(f"Found {len(files_to_merge)} files to merge.")

            # 1. Define the master column structure from the template file
            template_path = os.path.join(folder_path, template_file)
            self._log(f"Reading template file '{template_file}' to define columns.")
            self.progress_update.emit(0, f"Reading template: {template_file}")

            if file_type == 'Excel':
                template_df = pd.read_excel(template_path)
            elif file_type == 'CSV':
                template_df = pd.read_csv(template_path)
            else: # TXT
                # For TXT, assume a single column named 'Content' based on the template.
                template_df = pd.read_csv(template_path, header=None, names=['Content'])

            master_columns = template_df.columns.tolist()
            self._log(f"Template columns defined as: {master_columns}")

            # 2. Iterate and merge all files
            all_dfs = []
            total_files = len(files_to_merge)
            successfully_processed_count = 0
            for i, filename in enumerate(files_to_merge):
                file_path = os.path.join(folder_path, filename)
                self.progress_update.emit(int((i / total_files) * 100), f"Processing {filename}...")

                try:
                    if file_type == 'Excel':
                        current_df = pd.read_excel(file_path)
                    elif file_type == 'CSV':
                        current_df = pd.read_csv(file_path)
                    else: # TXT
                        current_df = pd.read_csv(file_path, header=None, names=master_columns)

                    if not current_df.columns.equals(pd.Index(master_columns)):
                        self._log(f"Column mismatch in '{filename}'. Strategy: {mismatch_strategy}")
                        if mismatch_strategy == 'add_missing_cols':
                            # Re-index to match master_columns, adding NA for missing, dropping extra
                            current_df = current_df.reindex(columns=master_columns)
                        else: # skip_file
                            self._log(f"Skipping '{filename}' due to column mismatch.")
                            continue
                    
                    all_dfs.append(current_df)
                    successfully_processed_count += 1
                except Exception as e:
                    # Log a warning but continue processing other files
                    self._log(f"Warning: Could not process file '{filename}'. Error: {e}. Skipping.")

            if not all_dfs:
                # This is the key change: raise a specific, informative error.
                raise ValueError(f"Files were found ({len(files_to_merge)}), but none could be processed. "
                                 f"Check logs for warnings about individual files (e.g., corruption, format issues).")

            self._log(f"Concatenating {successfully_processed_count} successfully processed DataFrames...")
            self.progress_update.emit(99, "Finalizing merge...")
            merged_df = pd.concat(all_dfs, ignore_index=True)
            self.progress_update.emit(100, "Merge complete.")

            self.finished_signal.emit(merged_df)

        except Exception as e:
            # This is the single, reliable exit point for all errors in the thread.
            self.error_signal.emit(str(e))


# --- [NEW & FIXED] HELPER: The GUI Dialog for the File Merger ---
class _FileMergerDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None,
                 initial_variable: Optional[str] = None):
        super().__init__(parent)
        self.setWindowTitle("File Merger")
        self.setMinimumWidth(600)
        self.global_variables = global_variables

        main_layout = QVBoxLayout(self)

        # 1. Source Folder and File Type
        source_group = QGroupBox("Source Files")
        source_layout = QFormLayout(source_group)
        self.folder_path_edit = QLineEdit(); self.folder_path_edit.setReadOnly(True)
        browse_folder_button = QPushButton("Browse Folder...")
        folder_layout = QHBoxLayout(); folder_layout.addWidget(self.folder_path_edit); folder_layout.addWidget(browse_folder_button)

        self.file_type_combo = QComboBox(); self.file_type_combo.addItems(['CSV', 'Excel', 'TXT'])
        
        self.keyword_filter_edit = QLineEdit()
        self.keyword_filter_edit.setPlaceholderText("Optional: e.g., 'report' (case-insensitive)")

        source_layout.addRow("Folder Containing Files:", folder_layout)
        source_layout.addRow("File Type to Merge:", self.file_type_combo)
        source_layout.addRow("Filter by Keyword in Name:", self.keyword_filter_edit)
        main_layout.addWidget(source_group)

        # 2. Template and Options
        options_group = QGroupBox("Merging Options")
        options_layout = QFormLayout(options_group)
        self.template_file_combo = QComboBox()
        self.template_file_combo.setToolTip("Select one file to act as the 'master' for column structure.")

        self.mismatch_strategy_combo = QComboBox()
        self.mismatch_strategy_combo.addItems(['add_missing_cols', 'skip_file'])
        self.mismatch_strategy_combo.setItemText(0, 'Align to template (add/drop cols)')
        self.mismatch_strategy_combo.setItemText(1, 'Skip files with different columns')

        options_layout.addRow("Template File (for columns):", self.template_file_combo)
        options_layout.addRow("If Columns Differ:", self.mismatch_strategy_combo)
        main_layout.addWidget(options_group)

        # 3. Assign Results
        assign_group = QGroupBox("Assign Merged DataFrame to Variable")
        assign_layout = QFormLayout(assign_group)
        self.new_var_radio = QRadioButton("New Variable Name:"); self.new_var_input = QLineEdit("merged_data")
        self.existing_var_radio = QRadioButton("Existing Variable:"); self.existing_var_combo = QComboBox()
        self.existing_var_combo.addItems(["-- Select --"] + self.global_variables)
        assign_layout.addRow(self.new_var_radio, self.new_var_input)
        assign_layout.addRow(self.existing_var_radio, self.existing_var_combo)
        self.new_var_radio.setChecked(True)
        main_layout.addWidget(assign_group)

        # Dialog Buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # Connections
        browse_folder_button.clicked.connect(self._browse_for_folder)
        self.file_type_combo.currentTextChanged.connect(self._update_file_list)
        self.keyword_filter_edit.textChanged.connect(self._update_file_list) # Update list when keyword changes
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        if initial_config: self._populate_from_initial_config(initial_config, initial_variable)

    def _browse_for_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder Containing Files to Merge")
        if folder_path:
            self.folder_path_edit.setText(folder_path)
            self._update_file_list()

    def _update_file_list(self):
        folder_path = self.folder_path_edit.text()
        file_type = self.file_type_combo.currentText()
        keyword = self.keyword_filter_edit.text().lower()
        self.template_file_combo.clear()

        if not folder_path: return

        ext_map = {'Excel': ('.xlsx', '.xls'), 'CSV': ('.csv',), 'TXT': ('.txt',)}
        extensions = ext_map.get(file_type)
        if not extensions: return

        try:
            files = sorted([
                f for f in os.listdir(folder_path)
                if os.path.isfile(os.path.join(folder_path, f))
                and f.endswith(extensions)
                and (keyword in f.lower())
            ])
            self.template_file_combo.addItems(files)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not read directory: {e}")

    def _populate_from_initial_config(self, config, variable):
        self.folder_path_edit.setText(config.get("folder_path", ""))
        self.file_type_combo.setCurrentText(config.get("file_type", "CSV"))
        self.keyword_filter_edit.setText(config.get("keyword", ""))
        
        # This will populate the template combo based on folder, type, and keyword
        self._update_file_list() 
        
        self.template_file_combo.setCurrentText(config.get("template_file", ""))
        self.mismatch_strategy_combo.setCurrentText(config.get("mismatch_strategy", "add_missing_cols"))

        if variable:
            if variable in self.global_variables:
                self.existing_var_radio.setChecked(True); self.existing_var_combo.setCurrentText(variable)
            else:
                self.new_var_radio.setChecked(True); self.new_var_input.setText(variable)

    def get_executor_method_name(self) -> str: return "_merge_files"

    def get_assignment_variable(self) -> Optional[str]:
        if self.new_var_radio.isChecked():
            var_name = self.new_var_input.text().strip()
            if not var_name:
                QMessageBox.warning(self, "Input Error", "New variable name cannot be empty."); return None
            return var_name
        else:
            var_name = self.existing_var_combo.currentText()
            if var_name == "-- Select --":
                QMessageBox.warning(self, "Input Error", "Please select an existing variable."); return None
            return var_name

    def get_config_data(self) -> Optional[Dict[str, Any]]:
        if not self.folder_path_edit.text():
            QMessageBox.warning(self, "Input Error", "Please select a source folder."); return None
        if not self.template_file_combo.currentText():
            QMessageBox.warning(self, "Input Error", "No files match your criteria. Please select a template file."); return None

        return {
            "folder_path": self.folder_path_edit.text(),
            "file_type": self.file_type_combo.currentText(),
            "keyword": self.keyword_filter_edit.text(),
            "template_file": self.template_file_combo.currentText(),
            "mismatch_strategy": self.mismatch_strategy_combo.currentText(),
        }


# --- [NEW & FIXED] The Public-Facing Module Class for File Merging ---
class Merge_file:
    """A module to merge multiple files from a folder into a single DataFrame."""
    def __init__(self, context: Optional[ExecutionContext] = None):
        self.context = context
        self.worker: Optional[_FileMergerThread] = None

    def _log(self, message: str):
        if self.context: self.context.add_log(message)
        else: print(message)

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        self._log("Opening File Merger configuration...")
        return _FileMergerDialog(
            global_variables=global_variables,
            parent=parent_window,
            **kwargs
        )

    def _merge_files(self, context: ExecutionContext, config_data: dict) -> pd.DataFrame:
        """Executes the file merging process and returns the resulting DataFrame."""
        from PyQt6.QtCore import QEventLoop

        self.context = context
        self.worker = _FileMergerThread(config_data, context)

        merged_dataframe = None
        error_message = ""
        loop = QEventLoop()

        def on_finished(df):
            nonlocal merged_dataframe
            merged_dataframe = df
            loop.quit()

        def on_error(msg):
            nonlocal error_message
            error_message = msg
            loop.quit()

        if hasattr(context, 'update_progress_bar'):
            self.worker.progress_update.connect(context.update_progress_bar)
        else:
            self.worker.progress_update.connect(lambda p, msg: self._log(f"Progress ({p}%): {msg}"))
            
        self.worker.finished.connect(loop.quit)
        self.worker.finished_signal.connect(on_finished)
        self.worker.error_signal.connect(on_error)

        self.worker.start()
        loop.exec() # Block here until the worker calls loop.quit()

        # After the loop finishes, check the results
        if error_message:
            # Raise the specific error message from the thread
            raise Exception(f"File merging failed: {error_message}")

        if merged_dataframe is None:
            # This is a fallback for unexpected thread termination
            raise Exception("Merging process ended unexpectedly without a result or a clear error.")

        self._log(f"Successfully merged files into a DataFrame with {len(merged_dataframe)} rows.")
        return merged_dataframe