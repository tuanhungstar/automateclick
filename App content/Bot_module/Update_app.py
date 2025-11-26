# File: bot_compiler_module.py
# This module contains classes for compiling a bot's source code into a DataFrame
# and for updating/deploying a bot from a DataFrame back into a folder structure.

import sys
from typing import Optional, List, Dict, Any
import pandas as pd
import os
import datetime

# --- PyQt6 Imports ---
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QDialogButtonBox,
    QComboBox, QWidget, QGroupBox, QMessageBox, QLabel, QFileDialog,
    QHBoxLayout, QRadioButton, QProgressBar
)
from PyQt6.QtCore import QThread, pyqtSignal, QEventLoop

# --- Main App Imports (Fallback for standalone testing) ---
try:
    # This should be your actual shared context from the main application
    from my_lib.shared_context import ExecutionContext
except ImportError:
    print("Warning: Could not import main app libraries. Using fallbacks for testing.")
    # Fallback class for standalone testing
    class ExecutionContext:
        def add_log(self, message: str): print(message)
        def get_variable(self, name: str):
            print(f"Fallback: Getting variable '{name}'")
            if "compile" in name: # Simulate getting an empty dataframe for Update_App
                return pd.DataFrame(columns=['file_path', 'last_modified', 'content'])
            return None
        def update_progress_bar(self, percent, message):
            print(f"Progress ({percent}%): {message}")


#
# --- HELPER (Compile): Worker thread for reading all files ---
#
class _CompilerWorkerThread(QThread):
    """Reads files from a directory structure into a DataFrame in a background thread."""
    progress_update = pyqtSignal(int, str)  # (percentage, message)
    finished_signal = pyqtSignal(pd.DataFrame)
    error_signal = pyqtSignal(str)

    def __init__(self, config: dict, context: Optional[ExecutionContext] = None):
        super().__init__()
        self.config = config
        self.context = context

    def _log(self, message: str):
        if self.context: self.context.add_log(message)
        else: print(message)

    def run(self):
        try:
            root_folder = self.config['root_folder']
            extensions = self.config['extensions']
            exclude_folders = self.config['exclude_folders']
            
            self._log(f"Starting compilation from root folder: {root_folder}")
            self._log(f"Including file extensions: {extensions}")
            self._log(f"Excluding folders: {exclude_folders}")

            file_data = []
            files_to_process = []
            
            # First, walk the directory to find all relevant files
            for dirpath, dirnames, filenames in os.walk(root_folder):
                # Modify dirnames in-place to prevent os.walk from descending into excluded folders
                dirnames[:] = [d for d in dirnames if d not in exclude_folders]
                
                for filename in filenames:
                    if filename.lower().endswith(tuple(f".{ext}" for ext in extensions)):
                        full_path = os.path.join(dirpath, filename)
                        files_to_process.append(full_path)
            
            if not files_to_process:
                self._log("Warning: No matching files found to compile.")
                self.finished_signal.emit(pd.DataFrame(columns=['file_path', 'last_modified', 'content']))
                return

            # Now, process the files and read their content
            total_files = len(files_to_process)
            for i, full_path in enumerate(files_to_process):
                relative_path = os.path.relpath(full_path, root_folder)
                self.progress_update.emit(int((i / total_files) * 100), f"Reading: {relative_path}")

                try:
                    # Get last modified time
                    mod_time = os.path.getmtime(full_path)
                    last_modified = datetime.datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d %H:%M:%S')

                    # Read file content
                    with open(full_path, 'r', encoding='utf-8', errors='ignore') as f:
                        content = f.read()
                    
                    file_data.append({
                        'file_path': relative_path.replace(os.sep, '/'), # Standardize path separators
                        'last_modified': last_modified,
                        'content': content
                    })
                except Exception as e:
                    self._log(f"Warning: Could not read file '{relative_path}'. Error: {e}. Skipping.")

            self._log(f"Successfully compiled {len(file_data)} files into a DataFrame.")
            df = pd.DataFrame(file_data, columns=['file_path', 'last_modified', 'content'])
            self.finished_signal.emit(df)

        except Exception as e:
            self.error_signal.emit(f"A critical error occurred during compilation: {e}")

#
# --- HELPER (Compile): The GUI Dialog for the Compiler ---
#
class _CompileDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None, **kwargs):
        super().__init__(parent)
        self.setWindowTitle("Compile Application Files")
        self.setMinimumWidth(600)
        self.global_variables = global_variables
        initial_config = kwargs.get("initial_config", {})
        initial_variable = kwargs.get("initial_variable")

        main_layout = QVBoxLayout(self)

        # 1. Source Configuration
        source_group = QGroupBox("Source Configuration")
        source_layout = QFormLayout(source_group)
        
        self.folder_path_edit = QLineEdit()
        browse_button = QPushButton("Browse...")
        path_layout = QHBoxLayout(); path_layout.addWidget(self.folder_path_edit); path_layout.addWidget(browse_button)

        self.extensions_edit = QLineEdit("py;json;csv;txt")
        self.extensions_edit.setToolTip("File extensions to include, separated by semicolons (e.g., py;json;txt)")
        self.exclude_edit = QLineEdit("__pycache__;venv;.git")
        self.exclude_edit.setToolTip("Folder names to exclude, separated by semicolons")

        source_layout.addRow("Root Source Folder:", path_layout)
        source_layout.addRow("File Extensions:", self.extensions_edit)
        source_layout.addRow("Exclude Folders:", self.exclude_edit)
        main_layout.addWidget(source_group)

        # 2. Assign Results
        assign_group = QGroupBox("Assign Compiled DataFrame to Variable")
        assign_layout = QFormLayout(assign_group)
        self.new_var_radio = QRadioButton("New Variable Name:"); self.new_var_input = QLineEdit("compiled_app_df")
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
        browse_button.clicked.connect(self._browse_for_folder)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        # Populate from initial config if provided
        self._populate_from_initial(initial_config, initial_variable)

    def _browse_for_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Root Source Folder")
        if folder_path:
            self.folder_path_edit.setText(folder_path)

    def _populate_from_initial(self, config, variable):
        self.folder_path_edit.setText(config.get("root_folder", ""))
        self.extensions_edit.setText(config.get("extensions_str", "py;json;csv;txt"))
        self.exclude_edit.setText(config.get("exclude_str", "__pycache__;venv;.git"))
        if variable:
            if variable in self.global_variables:
                self.existing_var_radio.setChecked(True); self.existing_var_combo.setCurrentText(variable)
            else:
                self.new_var_radio.setChecked(True); self.new_var_input.setText(variable)

    def get_executor_method_name(self) -> str: return "_compile_app"
    
    def get_assignment_variable(self) -> Optional[str]:
        if self.new_var_radio.isChecked():
            var_name = self.new_var_input.text().strip()
            if not var_name: QMessageBox.warning(self, "Input Error", "New variable name cannot be empty."); return None
            return var_name
        else:
            var_name = self.existing_var_combo.currentText()
            if var_name == "-- Select --": QMessageBox.warning(self, "Input Error", "Please select an existing variable."); return None
            return var_name

    def get_config_data(self) -> Optional[Dict[str, Any]]:
        root_folder = self.folder_path_edit.text()
        if not root_folder or not os.path.isdir(root_folder):
            QMessageBox.warning(self, "Input Error", "Please select a valid root source folder."); return None
        
        extensions_str = self.extensions_edit.text().strip().lower()
        exclude_str = self.exclude_edit.text().strip()

        return {
            "root_folder": root_folder,
            "extensions_str": extensions_str,
            "exclude_str": exclude_str,
            "extensions": [ext.strip() for ext in extensions_str.split(';') if ext.strip()],
            "exclude_folders": {folder.strip() for folder in exclude_str.split(';') if folder.strip()}
        }

#
# --- PUBLIC CLASS: Compile_App ---
#
class Compile_App:
    """A module to scan a folder structure and compile specified files into a single DataFrame."""
    def __init__(self, context: Optional[ExecutionContext] = None):
        self.context = context
        self.worker: Optional[_CompilerWorkerThread] = None

    def _log(self, message: str):
        if self.context: self.context.add_log(message)
        else: print(message)

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        self._log("Opening Application Compiler configuration...")
        return _CompileDialog(global_variables, parent_window, **kwargs)

    def _compile_app(self, context: ExecutionContext, config_data: dict) -> pd.DataFrame:
        """Executes the file compilation process and returns the resulting DataFrame."""
        self.context = context
        self.worker = _CompilerWorkerThread(config_data, context)

        compiled_dataframe = None
        error_message = ""
        loop = QEventLoop()

        def on_finished(df):
            nonlocal compiled_dataframe
            compiled_dataframe = df
            loop.quit()

        def on_error(msg):
            nonlocal error_message
            error_message = msg
            loop.quit()

        # Connect signals
        if hasattr(context, 'update_progress_bar'):
            self.worker.progress_update.connect(context.update_progress_bar)
        else:
            self.worker.progress_update.connect(lambda p, msg: self._log(f"Progress ({p}%): {msg}"))
        
        self.worker.finished_signal.connect(on_finished)
        self.worker.error_signal.connect(on_error)
        self.worker.finished.connect(loop.quit)

        self.worker.start()
        loop.exec() # Block until the worker finishes

        if error_message:
            raise Exception(f"Application compilation failed: {error_message}")
        if compiled_dataframe is None:
            raise Exception("Compilation ended unexpectedly without a result or error.")

        return compiled_dataframe


################################################################################
# --- UPDATE APPLICATION CLASSES ---
################################################################################

#
# --- HELPER (Update): Worker thread for writing all files ---
#
class _UpdaterWorkerThread(QThread):
    """Writes files from a DataFrame to a directory structure in a background thread."""
    progress_update = pyqtSignal(int, str)  # (percentage, message)
    finished_signal = pyqtSignal(str) # (final_message)
    error_signal = pyqtSignal(str)

    def __init__(self, config: dict, dataframe: pd.DataFrame, context: Optional[ExecutionContext] = None):
        super().__init__()
        self.config = config
        self.dataframe = dataframe
        self.context = context

    def _log(self, message: str):
        if self.context: self.context.add_log(message)
        else: print(message)

    def run(self):
        try:
            target_folder = self.config['target_folder']
            df = self.dataframe
            
            self._log(f"Starting update/deployment to target folder: {target_folder}")

            # Validate DataFrame structure
            required_cols = {'file_path', 'content'}
            if not required_cols.issubset(df.columns):
                raise ValueError(f"DataFrame must contain columns: {required_cols}")

            total_files = len(df)
            if total_files == 0:
                self.finished_signal.emit("Input DataFrame is empty. No files were written.")
                return

            # Iterate through the DataFrame and create/overwrite files
            for i, row in df.iterrows():
                relative_path = row['file_path']
                content = row['content']
                
                # Normalize path for the current OS and create the full path
                os_rel_path = os.path.normpath(relative_path)
                full_path = os.path.join(target_folder, os_rel_path)
                
                self.progress_update.emit(int((i / total_files) * 100), f"Writing: {relative_path}")

                try:
                    # Ensure the directory for the file exists
                    os.makedirs(os.path.dirname(full_path), exist_ok=True)
                    
                    # Write the content to the file, overwriting if it exists
                    with open(full_path, 'w', encoding='utf-8') as f:
                        f.write(str(content)) # Ensure content is string
                
                except Exception as e:
                    self._log(f"Warning: Failed to write file '{relative_path}'. Error: {e}. Skipping.")

            final_message = f"Successfully wrote/updated {total_files} files in '{target_folder}'."
            self._log(final_message)
            self.finished_signal.emit(final_message)

        except Exception as e:
            self.error_signal.emit(f"A critical error occurred during the update: {e}")

#
# --- HELPER (Update): The GUI Dialog for the Updater ---
#
class _UpdateDialog(QDialog):
    def __init__(self, df_variables: List[str], parent: Optional[QWidget] = None, **kwargs):
        super().__init__(parent)
        self.setWindowTitle("Update/Deploy Application from DataFrame")
        self.setMinimumWidth(600)
        self.df_variables = df_variables
        initial_config = kwargs.get("initial_config", {})

        main_layout = QVBoxLayout(self)

        # 1. Source and Destination Configuration
        config_group = QGroupBox("Configuration")
        config_layout = QFormLayout(config_group)

        self.df_var_combo = QComboBox()
        self.df_var_combo.addItems(["-- Select DataFrame --"] + self.df_variables)
        self.df_var_combo.setToolTip("Select the DataFrame containing the file structure and content.")

        self.folder_path_edit = QLineEdit()
        self.folder_path_edit.setPlaceholderText("Select the root folder to write files into")
        browse_button = QPushButton("Browse...")
        path_layout = QHBoxLayout(); path_layout.addWidget(self.folder_path_edit); path_layout.addWidget(browse_button)

        config_layout.addRow("Source DataFrame:", self.df_var_combo)
        config_layout.addRow("Target Root Folder:", path_layout)
        main_layout.addWidget(config_group)

        # Dialog Buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # Connections
        browse_button.clicked.connect(self._browse_for_folder)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        # Populate from initial config
        self._populate_from_initial(initial_config)

    def _browse_for_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Target Root Folder")
        if folder_path:
            self.folder_path_edit.setText(folder_path)

    def _populate_from_initial(self, config):
        self.df_var_combo.setCurrentText(config.get("dataframe_var", "-- Select DataFrame --"))
        self.folder_path_edit.setText(config.get("target_folder", ""))

    def get_executor_method_name(self) -> str: return "_update_app"
    def get_assignment_variable(self) -> Optional[str]: return None # This action does not create a new variable

    def get_config_data(self) -> Optional[Dict[str, Any]]:
        df_var = self.df_var_combo.currentText()
        if df_var == "-- Select DataFrame --":
            QMessageBox.warning(self, "Input Error", "Please select a source DataFrame."); return None
        
        target_folder = self.folder_path_edit.text()
        if not target_folder:
            QMessageBox.warning(self, "Input Error", "Please select a valid target root folder."); return None

        return {
            "dataframe_var": df_var,
            "target_folder": target_folder,
        }

#
# --- PUBLIC CLASS: Update_App ---
#
class Update_App:
    """A module to write files from a specially formatted DataFrame to a folder structure."""
    def __init__(self, context: Optional[ExecutionContext] = None):
        self.context = context
        self.worker: Optional[_UpdaterWorkerThread] = None

    def _log(self, message: str):
        if self.context: self.context.add_log(message)
        else: print(message)

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        self._log("Opening Application Updater configuration...")
        
        # In a real app, you would have a more robust way to get only DataFrame variables
        # For now, we assume the main app passes a filtered list or we use all globals.
        df_variables = global_variables # Or a filtered list from the parent

        return _UpdateDialog(df_variables, parent_window, **kwargs)

    def _update_app(self, context: ExecutionContext, config_data: dict):
        """Executes the file writing process."""
        self.context = context
        df_var_name = config_data["dataframe_var"]
        
        # Retrieve the DataFrame from the context
        source_df = self.context.get_variable(df_var_name)
        if not isinstance(source_df, pd.DataFrame):
            raise TypeError(f"Variable '{df_var_name}' is not a pandas DataFrame.")

        self.worker = _UpdaterWorkerThread(config_data, source_df, context)

        error_message = ""
        loop = QEventLoop()

        def on_error(msg):
            nonlocal error_message
            error_message = msg
            loop.quit()

        # Connect signals
        if hasattr(context, 'update_progress_bar'):
            self.worker.progress_update.connect(context.update_progress_bar)
        else:
            self.worker.progress_update.connect(lambda p, msg: self._log(f"Progress ({p}%): {msg}"))

        self.worker.finished_signal.connect(loop.quit)
        self.worker.error_signal.connect(on_error)
        self.worker.finished.connect(loop.quit)
        
        self.worker.start()
        loop.exec() # Block until worker finishes

        if error_message:
            raise Exception(f"Application update failed: {error_message}")
        
        # No DataFrame is returned, the action is complete upon finishing without error.
        self._log("Update process finished.")
