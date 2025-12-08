import sys
import os
import base64
import json
from typing import Optional, List, Dict, Any

# --- External Library Imports ---
import pandas as pd
import requests

# --- PyQt6 Imports ---
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QDialogButtonBox,
    QComboBox, QWidget, QGroupBox, QMessageBox, QLabel, QTextEdit,
    QCheckBox, QFileDialog, QHBoxLayout, QRadioButton
)
from PyQt6.QtCore import Qt

# --- Main App Imports (with fallback for standalone testing) ---
try:
    from my_lib.shared_context import ExecutionContext
except ImportError:
    print("Warning: Could not import main app libraries. Using fallback for ExecutionContext.")
    class ExecutionContext:
        def add_log(self, message: str): print(f"LOG: {message}")
        def get_variable(self, name: str, default=None):
            print(f"Fallback: Getting variable '{name}'")
            return None

# --- Custom Dialog for Gemini API Configuration ---
class _GeminiConfigDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None,
                 initial_variable: Optional[str] = None):
        super().__init__(parent)
        self.setWindowTitle("Configure Gemini API Call")
        self.setMinimumWidth(600)
        self.global_variables = global_variables

        main_layout = QVBoxLayout(self)

        # 1. API Configuration Group
        api_group = QGroupBox("API Settings")
        api_layout = QFormLayout(api_group)
        self.url_edit = QLineEdit("https://aoai-farm.bosch-temp.com/api/openai/deployments/google-gemini-2-0-flash-lite/chat/completions")
        self.api_key_edit = QLineEdit()
        self.api_key_edit.setPlaceholderText("Enter your GenAI Platform API Key")
        self.api_key_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.proxies_edit = QLineEdit()
        self.proxies_edit.setPlaceholderText("Optional: e.g., http://127.0.0.1:3128")
        api_layout.addRow("API URL:", self.url_edit)
        api_layout.addRow("API Key:", self.api_key_edit)
        api_layout.addRow("Proxies:", self.proxies_edit)
        main_layout.addWidget(api_group)

        # 2. File Input Group (Optional)
        self.file_group = QGroupBox("File Input (Optional)")
        self.file_group.setCheckable(True)
        self.file_group.setChecked(False)
        file_layout = QFormLayout(self.file_group)

        # <<< NEW: Radio buttons for file source >>>
        self.file_source_layout = QHBoxLayout()
        self.browse_file_radio = QRadioButton("Browse for File")
        self.var_file_radio = QRadioButton("Use Global Variable")
        self.file_source_layout.addWidget(self.browse_file_radio)
        self.file_source_layout.addWidget(self.var_file_radio)
        file_layout.addRow(self.file_source_layout)

        # Hardcoded file path input
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        browse_button = QPushButton("Browse...")
        self.path_layout_widget = QWidget() # Widget to hold the browse UI
        path_layout = QHBoxLayout(self.path_layout_widget)
        path_layout.setContentsMargins(0,0,0,0)
        path_layout.addWidget(self.file_path_edit)
        path_layout.addWidget(browse_button)
        
        # Global variable file path input
        self.file_var_combo = QComboBox()
        self.file_var_combo.addItems(["-- Select Variable --"] + self.global_variables)

        file_layout.addRow("File Path:", self.path_layout_widget)
        file_layout.addRow("File Path Variable:", self.file_var_combo)
        
        self.file_type_label = QLabel("File will be converted to Base64.")
        file_layout.addRow(self.file_type_label)
        main_layout.addWidget(self.file_group)
        
        # 3. Prompt Input Group
        prompt_group = QGroupBox("Prompt")
        prompt_layout = QVBoxLayout(prompt_group)
        
        self.prompt_source_layout = QHBoxLayout()
        self.text_prompt_radio = QRadioButton("Enter Prompt Text")
        self.var_prompt_radio = QRadioButton("Use Global Variable for Prompt")
        self.prompt_source_layout.addWidget(self.text_prompt_radio)
        self.prompt_source_layout.addWidget(self.var_prompt_radio)
        
        self.prompt_text_edit = QTextEdit()
        self.prompt_text_edit.setPlaceholderText("Enter your prompt here. If including a file, you can reference it, e.g., 'Describe the attached image.'")
        
        self.prompt_var_combo = QComboBox()
        self.prompt_var_combo.addItems(["-- Select Variable --"] + self.global_variables)

        prompt_layout.addLayout(self.prompt_source_layout)
        prompt_layout.addWidget(self.prompt_text_edit)
        prompt_layout.addWidget(self.prompt_var_combo)
        main_layout.addWidget(prompt_group)

        # 4. Assign Results Group
        assign_group = QGroupBox("Assign Result (DataFrame) to Variable")
        assign_layout = QFormLayout(assign_group)
        self.new_var_radio = QRadioButton("New Variable Name:")
        self.new_var_input = QLineEdit("gemini_result_df")
        self.existing_var_radio = QRadioButton("Existing Variable:")
        self.existing_var_combo = QComboBox()
        self.existing_var_combo.addItems(["-- Select --"] + self.global_variables)
        assign_layout.addRow(self.new_var_radio, self.new_var_input)
        assign_layout.addRow(self.existing_var_radio, self.existing_var_combo)
        main_layout.addWidget(assign_group)

        # Dialog Buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # --- Connections ---
        browse_button.clicked.connect(self._browse_for_file)
        self.browse_file_radio.toggled.connect(self._toggle_file_input) # <<< NEW
        self.text_prompt_radio.toggled.connect(self._toggle_prompt_input)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        # --- Initial State ---
        self.browse_file_radio.setChecked(True) # <<< NEW
        self._toggle_file_input() # <<< NEW
        self.text_prompt_radio.setChecked(True)
        self._toggle_prompt_input()
        self.new_var_radio.setChecked(True)

        if initial_config:
            self._populate_from_initial_config(initial_config, initial_variable)

    def _browse_for_file(self):
        filters = "Supported Files (*.pdf *.png *.jpg *.jpeg *.webp);;All Files (*)"
        file_path, _ = QFileDialog.getOpenFileName(self, "Select PDF or Image File", "", filters)
        if file_path:
            self.file_path_edit.setText(file_path)

    def _toggle_file_input(self): # <<< NEW METHOD
        """Shows/hides the file input widgets based on radio button selection."""
        is_browse = self.browse_file_radio.isChecked()
        self.path_layout_widget.setVisible(is_browse)
        self.file_var_combo.setVisible(not is_browse)

    def _toggle_prompt_input(self):
        is_text_prompt = self.text_prompt_radio.isChecked()
        self.prompt_text_edit.setVisible(is_text_prompt)
        self.prompt_var_combo.setVisible(not is_text_prompt)

    def _populate_from_initial_config(self, config, variable):
        self.url_edit.setText(config.get("url", "https://aoai-farm.bosch-temp.com/api/openai/deployments/google-gemini-2-0-flash-lite/chat/completions"))
        self.api_key_edit.setText(config.get("api_key", ""))
        self.proxies_edit.setText(config.get("proxies", ""))

        if config.get("file_path_config"):
            self.file_group.setChecked(True)
            file_config = config.get("file_path_config", {})
            if file_config.get("type") == "variable":
                self.var_file_radio.setChecked(True)
                self.file_var_combo.setCurrentText(file_config.get("value", "-- Select Variable --"))
            else: # Default to hardcoded/browse
                self.browse_file_radio.setChecked(True)
                self.file_path_edit.setText(file_config.get("value", ""))
        
        if config.get("prompt_type") == "variable":
            self.var_prompt_radio.setChecked(True)
            self.prompt_var_combo.setCurrentText(config.get("prompt_value", "-- Select Variable --"))
        else:
            self.text_prompt_radio.setChecked(True)
            self.prompt_text_edit.setText(config.get("prompt_value", ""))

        if variable:
            if variable in self.global_variables:
                self.existing_var_radio.setChecked(True)
                self.existing_var_combo.setCurrentText(variable)
            else:
                self.new_var_radio.setChecked(True)
                self.new_var_input.setText(variable)

    def get_executor_method_name(self) -> str:
        return "_execute_gemini_call"

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
        url = self.url_edit.text().strip()
        if not url:
            QMessageBox.warning(self, "Input Error", "API URL is required."); return None

        api_key = self.api_key_edit.text().strip()
        if not api_key:
            QMessageBox.warning(self, "Input Error", "API Key is required."); return None

        prompt_type = "text" if self.text_prompt_radio.isChecked() else "variable"
        prompt_value = ""
        if prompt_type == "text":
            prompt_value = self.prompt_text_edit.toPlainText().strip()
            if not prompt_value:
                QMessageBox.warning(self, "Input Error", "Prompt text cannot be empty."); return None
        else:
            prompt_value = self.prompt_var_combo.currentText()
            if prompt_value == "-- Select Variable --":
                QMessageBox.warning(self, "Input Error", "Please select a global variable for the prompt."); return None
        
        # <<< NEW: File Path Configuration Logic >>>
        file_path_config = None
        if self.file_group.isChecked():
            if self.browse_file_radio.isChecked():
                file_path = self.file_path_edit.text()
                if not file_path:
                    QMessageBox.warning(self, "Input Error", "Please select a file or uncheck the 'File Input' box."); return None
                file_path_config = {"type": "hardcoded", "value": file_path}
            else: # Variable radio is checked
                file_var = self.file_var_combo.currentText()
                if file_var == "-- Select Variable --":
                    QMessageBox.warning(self, "Input Error", "Please select a global variable for the file path."); return None
                file_path_config = {"type": "variable", "value": file_var}

        return {
            "url": url,
            "api_key": api_key,
            "proxies": self.proxies_edit.text().strip(),
            "prompt_type": prompt_type,
            "prompt_value": prompt_value,
            "file_path_config": file_path_config, # Store the new config structure
        }

# --- The Public-Facing Module Class ---
class Gemini_API:
    def __init__(self, context: Optional[ExecutionContext] = None):
        self.context = context

    def _log(self, message: str):
        if self.context:
            self.context.add_log(f"Gemini_API: {message}")
        else:
            print(f"Gemini_API LOG: {message}")

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        self._log("Opening Gemini API configuration dialog.")
        return _GeminiConfigDialog(
            global_variables=global_variables,
            parent=parent_window,
            **kwargs
        )

    def _execute_gemini_call(self, context: ExecutionContext, config_data: dict) -> pd.DataFrame:
        self.context = context
        
        url = config_data['url']
        api_key = config_data['api_key']
        proxies_str = config_data.get('proxies')
        
        prompt_text = ""
        if config_data['prompt_type'] == 'variable':
            var_name = config_data['prompt_value']
            prompt_text = context.get_variable(var_name)
            if not isinstance(prompt_text, str):
                raise TypeError(f"Global variable '@{var_name}' must contain a string for the prompt.")
            self._log(f"Using prompt from global variable '@{var_name}'.")
        else:
            prompt_text = config_data['prompt_value']
            self._log("Using prompt from text input.")

        # <<< NEW: Resolve File Path and Encode >>>
        file_base64 = None
        file_mime_type = None
        attached_file_path = "N/A"
        file_path_config = config_data.get("file_path_config")
        
        if file_path_config:
            file_path = ""
            if file_path_config["type"] == "variable":
                var_name = file_path_config["value"]
                file_path = context.get_variable(var_name)
                if not isinstance(file_path, str) or not os.path.exists(file_path):
                     raise FileNotFoundError(f"Global variable '@{var_name}' does not contain a valid file path.")
                self._log(f"Reading file from global variable '@{var_name}': {file_path}")
            else: # hardcoded
                file_path = file_path_config["value"]
                if not os.path.exists(file_path):
                    raise FileNotFoundError(f"The specified file path does not exist: {file_path}")
                self._log(f"Reading file from path: {file_path}")
            
            attached_file_path = file_path
            try:
                with open(file_path, "rb") as f:
                    file_base64 = base64.b64encode(f.read()).decode('utf-8')
                
                ext = os.path.splitext(file_path)[1].lower()
                if ext == ".pdf": file_mime_type = "application/pdf"
                elif ext == ".png": file_mime_type = "image/png"
                elif ext in (".jpg", ".jpeg"): file_mime_type = "image/jpeg"
                elif ext == ".webp": file_mime_type = "image/webp"
                else:
                    raise ValueError(f"File type '{ext}' is not supported for inline content.")
            except Exception as e:
                self._log(f"FATAL FILE ERROR: Could not read or encode file '{file_path}'. Error: {e}")
                raise

        headers = {
            "Content-Type": "application/json",
            "genaiplatform-farm-subscription-key": api_key
        }

        message_content = [{"type": "text", "text": prompt_text}]

        if file_base64 and file_mime_type:
            self._log(f"Attaching file {os.path.basename(attached_file_path)} to the prompt.")
            message_content.append({
                "type": "image_url",
                "image_url": { "url": f"data:{file_mime_type};base64,{file_base64}" }
            })

        payload = {
            "model": "gemini-2.0-flash-lite",
            "messages": [{"role": "user", "content": message_content}]
        }

        proxies = {"http": proxies_str, "https": proxies_str} if proxies_str else None
        if proxies: self._log(f"Using proxy: {proxies_str}")

        self._log(f"Sending request to: {url}...")
        try:
            response = requests.post(url, headers=headers, data=json.dumps(payload), proxies=proxies)
            response.raise_for_status()
            
            json_response = response.json()
            self._log("Successfully received response from API.")
            
            records = []
            model_requested = payload.get("model", "N/A")

            if 'choices' in json_response and len(json_response['choices']) > 0:
                for choice in json_response['choices']:
                    record = {
                        "request_model": model_requested, "prompt_text": prompt_text,
                        "attached_file": os.path.basename(attached_file_path),
                        "response_id": json_response.get('id'), "response_created": json_response.get('created'),
                        "response_model": json_response.get('model'), "choice_index": choice.get('index'),
                        "choice_role": choice.get('message', {}).get('role'),
                        "response_content": choice.get('message', {}).get('content'),
                        "finish_reason": choice.get('finish_reason')
                    }
                    if 'usage' in json_response:
                        record['prompt_tokens'] = json_response['usage'].get('prompt_tokens')
                        record['completion_tokens'] = json_response['usage'].get('completion_tokens')
                        record['total_tokens'] = json_response['usage'].get('total_tokens')
                    records.append(record)
            else:
                self._log("API response did not contain any 'choices'.")
                records.append({
                    "request_model": model_requested, "prompt_text": prompt_text,
                    "attached_file": os.path.basename(attached_file_path),
                    "response_id": json_response.get('id'),
                    "response_content": "No content generated or choices array is empty.",
                })
            
            result_df = pd.DataFrame(records)
            self._log(f"Created DataFrame with {len(result_df)} row(s).")
            return result_df

        except requests.exceptions.ProxyError as e:
            self._log(f"FATAL PROXY ERROR: Could not connect to proxy. Error: {e}"); raise
        except requests.exceptions.RequestException as e:
            error_details = f"Error: {e}"
            if e.response is not None:
                error_details += f" | Status Code: {e.response.status_code} | Response: {e.response.text}"
            self._log(f"FATAL API ERROR: {error_details}"); raise