# File: Bot_module/Gemini_API.py

import sys
from typing import Optional, List, Dict, Any
import pandas as pd
import os
import google.genai
from google.genai import types 
import mimetypes 

# --- PyQt6 Imports ---
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QDialogButtonBox,
    QComboBox, QWidget, QGroupBox, QMessageBox, QLabel,
    QHBoxLayout, QRadioButton, QFileDialog
)
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
            return None 

#
# --- HELPER: The GUI Dialog for the Gemini API Call ---
#
class _GeminiAPIDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None,
                 initial_variable: Optional[str] = None):
        super().__init__(parent)
        self.setWindowTitle("Gemini API Call")
        self.setMinimumWidth(600)
        self.global_variables = global_variables

        main_layout = QVBoxLayout(self)

        # 1. API Key Configuration
        api_key_group = QGroupBox("Gemini API Key")
        api_key_layout = QFormLayout(api_key_group)
        self.api_key_hardcode_radio = QRadioButton("Enter Key Directly:")
        self.api_key_hardcode_input = QLineEdit()
        self.api_key_variable_radio = QRadioButton("Select Global Variable:")
        self.api_key_variable_combo = QComboBox()
        self.api_key_variable_combo.addItems(["-- Select --"] + self.global_variables)
        
        api_key_layout.addRow(self.api_key_hardcode_radio, self.api_key_hardcode_input)
        api_key_layout.addRow(self.api_key_variable_radio, self.api_key_variable_combo)
        self.api_key_variable_radio.setChecked(True)
        main_layout.addWidget(api_key_group)

        # 2. Prompt Configuration
        prompt_group = QGroupBox("Prompt Input")
        prompt_layout = QFormLayout(prompt_group)
        self.prompt_hardcode_radio = QRadioButton("Enter Prompt Directly:")
        self.prompt_hardcode_input = QLineEdit()
        self.prompt_variable_radio = QRadioButton("Select Global Variable:")
        self.prompt_variable_combo = QComboBox()
        self.prompt_variable_combo.addItems(["-- Select --"] + self.global_variables)
        
        prompt_layout.addRow(self.prompt_hardcode_radio, self.prompt_hardcode_input)
        prompt_layout.addRow(self.prompt_variable_radio, self.prompt_variable_combo)
        self.prompt_hardcode_radio.setChecked(True)
        main_layout.addWidget(prompt_group)

        # 3. File Input (New)
        file_group = QGroupBox("Optional File Attachment (Image/PDF)")
        file_layout = QFormLayout(file_group)
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        self.browse_file_button = QPushButton("Browse File...")
        
        file_path_layout = QHBoxLayout()
        file_path_layout.addWidget(self.file_path_edit)
        file_path_layout.addWidget(self.browse_file_button)
        
        file_layout.addRow("File Path (Image/PDF):", file_path_layout)
        main_layout.addWidget(file_group)

        # 4. Assign Results
        assign_group = QGroupBox("Assign Results to Variable")
        assign_layout = QFormLayout(assign_group)
        self.new_var_radio = QRadioButton("New Variable Name:"); self.new_var_input = QLineEdit("gemini_response")
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
        self.browse_file_button.clicked.connect(self._browse_for_file)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        
        if initial_config: self._populate_from_initial_config(initial_config, initial_variable)

    def _browse_for_file(self):
        filters = "Multi-modal Files (*.jpg *.jpeg *.png *.pdf);;All Files (*)"
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Image or PDF File", "", filters)
        if file_path:
            self.file_path_edit.setText(file_path)

    def _populate_from_initial_config(self, config, variable):
        # API Key Config
        if config.get("api_key_source") == "variable":
            self.api_key_variable_radio.setChecked(True)
            self.api_key_variable_combo.setCurrentText(config.get("api_key_value", ""))
        else:
            self.api_key_hardcode_radio.setChecked(True)
            self.api_key_hardcode_input.setText(config.get("api_key_value", ""))
            
        # Prompt Config
        if config.get("prompt_source") == "variable":
            self.prompt_variable_radio.setChecked(True)
            self.prompt_variable_combo.setCurrentText(config.get("prompt_value", ""))
        else:
            self.prompt_hardcode_radio.setChecked(True)
            self.prompt_hardcode_input.setText(config.get("prompt_value", ""))

        # File Config
        self.file_path_edit.setText(config.get("file_path", ""))

        # Assignment Config
        if variable:
            if variable in self.global_variables:
                self.existing_var_radio.setChecked(True); self.existing_var_combo.setCurrentText(variable)
            else:
                self.new_var_radio.setChecked(True); self.new_var_input.setText(variable)

    def get_executor_method_name(self) -> str: return "_call_gemini_api"

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
        # API Key validation
        api_key_source = "hardcode" if self.api_key_hardcode_radio.isChecked() else "variable"
        api_key_value = self.api_key_hardcode_input.text().strip() if api_key_source == "hardcode" else self.api_key_variable_combo.currentText()
        if not api_key_value or api_key_value == "-- Select --":
            QMessageBox.warning(self, "Input Error", "Please provide or select a variable for the API Key."); return None

        # Prompt validation
        prompt_source = "hardcode" if self.prompt_hardcode_radio.isChecked() else "variable"
        prompt_value = self.prompt_hardcode_input.text().strip() if prompt_source == "hardcode" else self.prompt_variable_combo.currentText()
        
        # File Path (Optional, no validation needed yet)
        file_path = self.file_path_edit.text().strip()

        if not prompt_value and not file_path:
            QMessageBox.warning(self, "Input Error", "You must provide either a Prompt or a File to analyze."); return None

        return {
            "api_key_source": api_key_source,
            "api_key_value": api_key_value,
            "prompt_source": prompt_source,
            "prompt_value": prompt_value,
            "file_path": file_path
        }

#
# --- The Public-Facing Module Class for Gemini API Call ---
#
class Gemini_API:
    """A module to interact with the Google Gemini API using the 'google-genai' SDK."""
    def __init__(self, context: Optional[ExecutionContext] = None):
        self.context = context

    def _log(self, message: str):
        if self.context: self.context.add_log(message)
        else: print(message)

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        """Opens the configuration dialog for the Gemini API call."""
        self._log("Opening Gemini API Call configuration...")
        return _GeminiAPIDialog(
            global_variables=global_variables,
            parent=parent_window,
            **kwargs
        )

    def _call_gemini_api(self, context: ExecutionContext, config_data: dict) -> pd.DataFrame:
        """
        Executes the real Gemini API call, supporting multimodal content (Image/PDF).
        """
        self.context = context
        uploaded_file = None
        
        # 1. Resolve API Key
        api_key_value = config_data["api_key_value"]
        api_key = None
        if config_data["api_key_source"] == "variable":
            api_key = self.context.get_variable(api_key_value)
            self._log(f"Fetching API Key from variable: '{api_key_value}'")
        else:
            api_key = api_key_value
            self._log("Using hardcoded API Key.")
            
        if not api_key:
             raise ValueError(f"API Key not found or variable '{api_key_value}' is empty.")

        # 2. Resolve Prompt
        prompt_value = config_data["prompt_value"]
        prompt = None
        if config_data["prompt_source"] == "variable":
            prompt = self.context.get_variable(prompt_value)
            self._log(f"Fetching Prompt from variable: '{prompt_value}'")
        else:
            prompt = prompt_value
            self._log("Using hardcoded Prompt input.")
            
        # 3. Initialize client (needed for file upload too)
        try:
            client = google.genai.Client(api_key=str(api_key))
        except Exception as e:
            self._log(f"FATAL ERROR: Failed to initialize Gemini Client: {e}")
            raise

        # 4. Construct contents for multimodal prompt
        file_path = config_data.get("file_path")
        contents = []

        try:
            # Handle file upload if path is provided
            if file_path and os.path.exists(file_path):
                self._log(f"Attempting to upload file: {os.path.basename(file_path)}")
                
                # FIX: Explicitly pass the file path using the keyword 'file='
                # This resolves the "takes 1 positional argument but 2 were given" error.
                uploaded_file = client.files.upload(file=file_path) 
                
                contents.append(uploaded_file)
                self._log(f"File uploaded successfully to: {uploaded_file.uri}")

            # Add the user's text prompt (can be empty if file is provided)
            if prompt:
                contents.append(str(prompt))

            if not contents:
                raise ValueError("No content (prompt or file) was provided for the API call.")

            self._log(f"Sending request to Gemini API with {len(contents)} parts.")

            # Make the API call
            response = client.models.generate_content(
                model="gemini-2.5-flash", 
                contents=contents,
            )
            
            generated_text = response.text
            self._log(f"Successfully received response (length: {len(generated_text)}).")

        except Exception as e:
            error_message = f"FATAL ERROR during Gemini API call: {e}"
            self._log(error_message)
            # Re-raise the error for the execution context to catch
            raise RuntimeError(error_message)
            
        finally:
            # Clean up the uploaded file to avoid unnecessary storage/cost
            if uploaded_file:
                try:
                    client.files.delete(name=uploaded_file.name)
                    self._log(f"Cleaned up uploaded file: {uploaded_file.name}")
                except Exception as cleanup_e:
                    # Log a warning if cleanup fails, but don't halt execution
                    self._log(f"Warning: Failed to delete uploaded file {uploaded_file.name}. Error: {cleanup_e}")

        # 5. Return result as a DataFrame
        df = pd.DataFrame([generated_text], columns=['Response'])
        return df