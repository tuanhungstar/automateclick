import sys
import os
import json
import tempfile
import requests
from typing import Optional, List, Dict, Any
import pandas as pd
import pyautogui
import pygetwindow as gw
import ctypes
import time
from PIL import Image, ImageDraw, ImageGrab

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QDialogButtonBox,
    QComboBox, QWidget, QGroupBox, QMessageBox, QLabel,
    QHBoxLayout, QRadioButton, QFileDialog, QTextEdit, QCheckBox, QSpinBox
)
from PyQt6.QtCore import Qt

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
# --- HELPER: The GUI Dialog for the Local AI API Call ---
#
class _LocalAIDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None,
                 initial_variable: Optional[str] = None):
        super().__init__(parent)
        self.setWindowTitle("Local AI Call")
        self.setMinimumWidth(600)
        self.global_variables = global_variables

        main_layout = QVBoxLayout(self)

        # 1. Server URL Configuration
        url_group = QGroupBox("Local AI Server URL")
        url_layout = QFormLayout(url_group)
        self.url_hardcode_radio = QRadioButton("Enter URL Directly:")
        self.url_hardcode_input = QLineEdit("http://localhost:8001")
        self.url_variable_radio = QRadioButton("Select Global Variable:")
        self.url_variable_combo = QComboBox()
        self.url_variable_combo.addItems(["-- Select --"] + [str(v) for v in self.global_variables])
        
        url_layout.addRow(self.url_hardcode_radio, self.url_hardcode_input)
        url_layout.addRow(self.url_variable_radio, self.url_variable_combo)
        self.url_hardcode_radio.setChecked(True)
        main_layout.addWidget(url_group)

        # 2. Action Configuration
        action_group = QGroupBox("Endpoint / Action")
        action_layout = QVBoxLayout(action_group)
        self.action_combo = QComboBox()
        self.action_combo.addItems([
            "/generate (Image Generation)",
            "/transcribe (Audio Transcription)",
            "/story (Story Generation)",
            "/tts (Text to Speech)",
            "/ocr (Image OCR)",
            "/invoice (Invoice Processing)",
            "/detect (Object Detection)",
            "/prompt (Prompt Generation)",
            "/extract-product (Product Extraction)",
            "/extract-product-from-text (Plaintext Product Extraction)",
            "--- RAG Model (port 8001) ---",
            "/classify (RAG: Classify Product)",
            "/ingest-history (RAG: Ingest History XLSX/CSV)",
            "/ingest (RAG: Ingest Document PDF/TXT)",
            "/health (RAG: Health Check)",
            "/reset (RAG: Reset Database)"
        ])
        self.action_combo.currentIndexChanged.connect(self._on_action_changed)
        action_layout.addWidget(self.action_combo)
        main_layout.addWidget(action_group)

        # 3. Primary Text / Prompt Configuration
        self.prompt_group = QGroupBox("Text Input / Prompt")
        prompt_layout = QFormLayout(self.prompt_group)
        self.prompt_hardcode_radio = QRadioButton("Enter Directly (Multi-line):")
        self.prompt_hardcode_input = QTextEdit() 
        self.prompt_hardcode_input.setFixedHeight(80)
        self.prompt_variable_radio = QRadioButton("Select Global Variable:")
        self.prompt_variable_combo = QComboBox()
        self.prompt_variable_combo.addItems(["-- Select --"] + [str(v) for v in self.global_variables])
        
        self.source_type_label = QLabel("Source Type:")
        self.source_type_combo = QComboBox()
        self.source_type_combo.addItems(["HTML Content", "URL"])
        
        prompt_layout.addRow(self.source_type_label, self.source_type_combo)
        prompt_layout.addRow(self.prompt_hardcode_radio, self.prompt_hardcode_input)
        prompt_layout.addRow(self.prompt_variable_radio, self.prompt_variable_combo)
        self.prompt_hardcode_radio.setChecked(True)
        main_layout.addWidget(self.prompt_group)

        # 4. Secondary Text Configuration (for Product Ext, Prompt Gen, etc)
        self.secondary_group = QGroupBox("Secondary Text (Product Name / Word / URL)")
        secondary_layout = QFormLayout(self.secondary_group)
        self.secondary_hardcode_radio = QRadioButton("Enter Directly:")
        self.secondary_hardcode_input = QLineEdit()
        self.secondary_variable_radio = QRadioButton("Select Global Variable:")
        self.secondary_variable_combo = QComboBox()
        self.secondary_variable_combo.addItems(["-- Select --"] + [str(v) for v in self.global_variables])
        
        secondary_layout.addRow(self.secondary_hardcode_radio, self.secondary_hardcode_input)
        secondary_layout.addRow(self.secondary_variable_radio, self.secondary_variable_combo)
        self.secondary_hardcode_radio.setChecked(True)
        main_layout.addWidget(self.secondary_group)

        # 5. File Input
        self.file_group = QGroupBox("File Attachment")
        file_layout = QFormLayout(self.file_group)
        
        self.file_variable_radio = QRadioButton("Select Global Variable:")
        self.file_variable_combo = QComboBox()
        self.file_variable_combo.addItems(["-- Select --"] + [str(v) for v in self.global_variables])
        
        self.file_hardcode_radio = QRadioButton("Enter File Path Directly:")
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        self.browse_file_button = QPushButton("Browse File...")
        
        file_path_layout = QHBoxLayout()
        file_path_layout.addWidget(self.file_path_edit)
        file_path_layout.addWidget(self.browse_file_button)
        
        self.capture_window_radio = QRadioButton("Activate Window & Capture:")
        self.win_title_input = QLineEdit()
        self.win_title_input.setPlaceholderText("Window Title (e.g. Chrome)")
        
        self.simulation_checkbox = QCheckBox("Simulation (Draw detection rectangles)")
        self.simulation_checkbox.setVisible(False) # Only for /detect

        # Chunk size — only for /ingest-history
        self.chunk_size_label = QLabel("Chunk Size (rows):")
        self.chunk_size_spin = QSpinBox()
        self.chunk_size_spin.setRange(0, 1000000)
        self.chunk_size_spin.setValue(500)
        self.chunk_size_spin.setSpecialValueText("No chunking (send all at once)")
        self.chunk_size_spin.setSingleStep(100)
        self.chunk_size_label.setVisible(False)
        self.chunk_size_spin.setVisible(False)

        file_layout.addRow(self.file_variable_radio, self.file_variable_combo)
        file_layout.addRow(self.file_hardcode_radio, file_path_layout)
        file_layout.addRow(self.capture_window_radio, self.win_title_input)
        file_layout.addRow(self.simulation_checkbox)
        file_layout.addRow(self.chunk_size_label, self.chunk_size_spin)
        
        self.file_hardcode_radio.setChecked(True)
        main_layout.addWidget(self.file_group)

        # 6. Assign Results
        assign_group = QGroupBox("Assign Results to Variable")
        assign_layout = QFormLayout(assign_group)
        self.new_var_radio = QRadioButton("New Variable Name:")
        self.new_var_input = QLineEdit("local_ai_response")
        self.existing_var_radio = QRadioButton("Existing Variable:")
        self.existing_var_combo = QComboBox()
        self.existing_var_combo.addItems(["-- Select --"] + [str(v) for v in self.global_variables])
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
        
        self._on_action_changed()
        if initial_config: self._populate_from_initial_config(initial_config, initial_variable)

    def _on_action_changed(self):
        action = self.action_combo.currentText()
        # Default visibility
        self.prompt_group.setVisible(False)
        self.secondary_group.setVisible(False)
        self.file_group.setVisible(False)
        self.source_type_label.setVisible(False)
        self.source_type_combo.setVisible(False)
        self.chunk_size_label.setVisible(False)
        self.chunk_size_spin.setVisible(False)
        
        if "/generate" in action:
            self.prompt_group.setTitle("Image Prompt")
            self.prompt_group.setVisible(True)
        elif "/transcribe" in action:
            self.file_group.setTitle("Audio File (MP3/WAV/etc.)")
            self.file_group.setVisible(True)
        elif "/story" in action:
            self.prompt_group.setTitle("Words (Comma separated)")
            self.prompt_group.setVisible(True)
        elif "/tts" in action:
            self.prompt_group.setTitle("Text to Speech")
            self.prompt_group.setVisible(True)
        elif "/ocr" in action:
            self.file_group.setTitle("Image File")
            self.file_group.setVisible(True)
        elif "/invoice" in action:
            self.file_group.setTitle("Invoice File (Image/PDF)")
            self.prompt_group.setTitle("Custom Prompt (Optional)")
            self.file_group.setVisible(True)
            self.prompt_group.setVisible(True)
        elif "/detect" in action:
            self.file_group.setTitle("Image File / Capture")
            self.prompt_group.setTitle("Custom Prompt (Optional, e.g. 'List all buttons')")
            self.file_group.setVisible(True)
            self.prompt_group.setVisible(True)
            self.simulation_checkbox.setVisible(True)
        elif "/extract-product-from-text" in action:
            self.prompt_group.setTitle("Text Content")
            self.prompt_group.setVisible(True)
            self.secondary_group.setTitle("Product Name to Identify")
            self.secondary_group.setVisible(True)
        elif "/extract-product" in action:
            self.prompt_group.setTitle("Content (HTML or URL)")
            self.prompt_group.setVisible(True)
            self.secondary_group.setTitle("Product Name to Identify")
            self.secondary_group.setVisible(True)
            self.source_type_label.setVisible(True)
            self.source_type_combo.setVisible(True)
        elif "/prompt" in action:
            self.prompt_group.setTitle("Original Word / Text")
            self.prompt_group.setVisible(True)
            self.secondary_group.setTitle("Translation / Context")
            self.secondary_group.setVisible(True)
        elif "/classify" in action:
            self.prompt_group.setTitle("Product Info / Description")
            self.prompt_group.setVisible(True)
        elif "/ingest-history" in action:
            self.file_group.setTitle("History File (XLSX or CSV)")
            self.file_group.setVisible(True)
            self.chunk_size_label.setVisible(True)
            self.chunk_size_spin.setVisible(True)
        elif "/ingest" in action:
            self.file_group.setTitle("Document File (PDF or TXT)")
            self.file_group.setVisible(True)
        elif "/health" in action or "/reset" in action:
            pass  # No inputs needed
        else:
            self.simulation_checkbox.setVisible(False)

    def _browse_for_file(self):
        filters = "All Files (*);;Images (*.jpg *.jpeg *.png);;PDF (*.pdf);;Audio (*.mp3 *.wav *.m4a)"
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", filters)
        if file_path:
            self.file_path_edit.setText(file_path)
            self.file_hardcode_radio.setChecked(True)

    def _populate_from_initial_config(self, config, variable):
        if config.get("action"):
            self.action_combo.setCurrentText(config.get("action"))

        if config.get("url_source") == "variable":
            self.url_variable_radio.setChecked(True)
            self.url_variable_combo.setCurrentText(config.get("url_value", ""))
        else:
            self.url_hardcode_radio.setChecked(True)
            self.url_hardcode_input.setText(config.get("url_value", ""))
            
        if config.get("prompt_source") == "variable":
            self.prompt_variable_radio.setChecked(True)
            self.prompt_variable_combo.setCurrentText(config.get("prompt_value", ""))
        else:
            self.prompt_hardcode_radio.setChecked(True)
            self.prompt_hardcode_input.setPlainText(config.get("prompt_value", ""))

        if config.get("secondary_source") == "variable":
            self.secondary_variable_radio.setChecked(True)
            self.secondary_variable_combo.setCurrentText(config.get("secondary_value", ""))
        else:
            self.secondary_hardcode_radio.setChecked(True)
            self.secondary_hardcode_input.setText(config.get("secondary_value", ""))

        if config.get("file_source") == "variable":
            self.file_variable_radio.setChecked(True)
            self.file_variable_combo.setCurrentText(config.get("file_path_value", ""))
        else:
            self.file_hardcode_radio.setChecked(True)
            self.file_path_edit.setText(config.get("file_path_value", ""))

        if config.get("file_source") == "capture":
            self.capture_window_radio.setChecked(True)
            self.win_title_input.setText(config.get("window_title", ""))
        
        self.simulation_checkbox.setChecked(config.get("simulation", False))
        self.source_type_combo.setCurrentText(config.get("source_type", "HTML Content"))
        self.chunk_size_spin.setValue(config.get("chunk_size", 500))

        if variable:
            if variable in self.global_variables:
                self.existing_var_radio.setChecked(True)
                self.existing_var_combo.setCurrentText(str(variable))
            else:
                self.new_var_radio.setChecked(True); self.new_var_input.setText(variable)
        self._on_action_changed()

    def get_executor_method_name(self) -> str: return "_call_local_ai_api"

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
        url_source = "hardcode" if self.url_hardcode_radio.isChecked() else "variable"
        url_value = self.url_hardcode_input.text().strip() if url_source == "hardcode" else self.url_variable_combo.currentText()

        prompt_source = "hardcode" if self.prompt_hardcode_radio.isChecked() else "variable"
        prompt_value = self.prompt_hardcode_input.toPlainText().strip() if prompt_source == "hardcode" else self.prompt_variable_combo.currentText()
        
        secondary_source = "hardcode" if self.secondary_hardcode_radio.isChecked() else "variable"
        secondary_value = self.secondary_hardcode_input.text().strip() if secondary_source == "hardcode" else self.secondary_variable_combo.currentText()
        
        file_source = "hardcode" if self.file_hardcode_radio.isChecked() else ("capture" if self.capture_window_radio.isChecked() else "variable")
        file_path_value = self.file_path_edit.text().strip() if file_source == "hardcode" else self.file_variable_combo.currentText()
        if file_source == "variable" and file_path_value == "-- Select --":
             file_path_value = ""

        return {
            "action": self.action_combo.currentText(),
            "url_source": url_source,
            "url_value": url_value,
            "prompt_source": prompt_source,
            "prompt_value": prompt_value,
            "secondary_source": secondary_source,
            "secondary_value": secondary_value,
            "file_source": file_source,
            "file_path_value": file_path_value,
            "window_title": self.win_title_input.text().strip(),
            "simulation": self.simulation_checkbox.isChecked(),
            "source_type": self.source_type_combo.currentText(),
            "chunk_size": self.chunk_size_spin.value(),
            "assign_to": self.get_assignment_variable()
        }

#
# --- The Public-Facing Module Class ---
#
class Local_AI_API:
    """A module to interact with the Local AI server (FastAPI)."""
    def __init__(self, context: Optional[ExecutionContext] = None):
        self.context = context

    def _log(self, message: str):
        if self.context: self.context.add_log(message)
        else: print(message)

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        self._log("Opening Local AI API Call configuration...")
        return _LocalAIDialog(
            global_variables=global_variables,
            parent=parent_window,
            **kwargs
        )

    def _call_local_ai_api(self, context: ExecutionContext, config_data: dict) -> pd.DataFrame:
        self.context = context
        
        # 1. Resolve URL
        url_val = config_data["url_value"]
        base_url = self.context.get_variable(url_val) if config_data["url_source"] == "variable" else url_val
        if not base_url:
            base_url = "http://api-localai.germantest.net"
        base_url = base_url.rstrip("/")

        # 2. Resolve Inputs
        prompt_val = config_data.get("prompt_value", "")
        prompt = self.context.get_variable(prompt_val) if config_data.get("prompt_source") == "variable" else prompt_val

        sec_val = config_data.get("secondary_value", "")
        secondary = self.context.get_variable(sec_val) if config_data.get("secondary_source") == "variable" else sec_val

        file_val = config_data.get("file_path_value", "")
        file_path = self.context.get_variable(file_val) if config_data.get("file_source") == "variable" else file_val

        action_str = config_data.get("action", "")
        endpoint = action_str.split(" ")[0] # extract "/generate" from "/generate (Image Generation)"

        self._log(f"Calling Local AI endpoint: {base_url}{endpoint}")

        result_data = None

        try:
            full_url = f"{base_url}{endpoint}"
            
            if endpoint == "/generate":
                if not prompt: raise ValueError("Prompt is required for /generate")
                response = requests.post(full_url, json={"prompt": str(prompt)}, timeout=300)
                response.raise_for_status()
                # Save binary content to temp file
                temp_dir = os.path.join(os.path.dirname(__file__), "..", "temps")
                os.makedirs(temp_dir, exist_ok=True)
                out_file = os.path.join(temp_dir, f"local_ai_gen_{os.urandom(4).hex()}.png")
                with open(out_file, "wb") as f:
                    f.write(response.content)
                result_data = out_file
                self._log(f"Image saved to {out_file}")

            elif endpoint == "/transcribe":
                if not file_path or not os.path.exists(file_path): raise ValueError(f"File not found: {file_path}")
                with open(file_path, "rb") as f:
                    files = {"file": (os.path.basename(file_path), f)}
                    response = requests.post(full_url, files=files, timeout=600)
                response.raise_for_status()
                result_data = response.json().get("text", "")

            elif endpoint == "/story":
                if not prompt: raise ValueError("Words are required for /story")
                words = [w.strip() for w in str(prompt).split(",") if w.strip()]
                response = requests.post(full_url, json={"words": words}, timeout=300)
                response.raise_for_status()
                result_data = response.json().get("story", "")

            elif endpoint == "/tts":
                if not prompt: raise ValueError("Text is required for /tts")
                response = requests.post(full_url, json={"text": str(prompt)}, timeout=300)
                response.raise_for_status()
                temp_dir = os.path.join(os.path.dirname(__file__), "..", "temps")
                os.makedirs(temp_dir, exist_ok=True)
                out_file = os.path.join(temp_dir, f"local_ai_tts_{os.urandom(4).hex()}.mp3")
                with open(out_file, "wb") as f:
                    f.write(response.content)
                result_data = out_file
                self._log(f"Audio saved to {out_file}")

            elif endpoint == "/ocr":
                if not file_path or not os.path.exists(file_path): raise ValueError(f"Image File not found: {file_path}")
                with open(file_path, "rb") as f:
                    files = {"file": (os.path.basename(file_path), f)}
                    response = requests.post(full_url, files=files, timeout=300)
                response.raise_for_status()
                result_data = json.dumps(response.json(), ensure_ascii=False, indent=2)

            elif endpoint == "/invoice":
                if not file_path or not os.path.exists(file_path): raise ValueError(f"File not found: {file_path}")
                data = {"prompt": str(prompt)} if prompt else {}
                with open(file_path, "rb") as f:
                    files = {"file": (os.path.basename(file_path), f)}
                    response = requests.post(full_url, data=data, files=files, timeout=300)
                response.raise_for_status()
                result_data = json.dumps(response.json(), ensure_ascii=False, indent=2)


            elif endpoint == "/extract-product-from-text":
                if not prompt or not secondary: raise ValueError("Text Content and Product Name are required")
                payload = {
                    "content": str(prompt),
                    "myproduct": str(secondary)
                }
                response = requests.post(full_url, json=payload, timeout=300)
                response.raise_for_status()
                result_data = json.dumps(response.json(), ensure_ascii=False, indent=2)

            elif endpoint == "/extract-product":
                if not prompt or not secondary: raise ValueError("Content/URL and Product Name are required")
                p_str = str(prompt)
                payload = {"myproduct": str(secondary)}
                
                source_type = config_data.get("source_type", "HTML Content")
                if "URL" in source_type:
                    payload["url"] = p_str
                elif "HTML" in source_type:
                    payload["content"] = p_str
                else:
                    # Fallback to auto-detection
                    if p_str.startswith("http://") or p_str.startswith("https://"):
                        payload["url"] = p_str
                    else:
                        payload["content"] = p_str
                        
                response = requests.post(full_url, json=payload, timeout=300)
                response.raise_for_status()
                result_data = json.dumps(response.json(), ensure_ascii=False, indent=2)

            elif endpoint == "/detect":
                # --- INTEGRATED CAPTURE FOR /detect ---
                if config_data.get("file_source") == "capture":
                    win_title = config_data.get("window_title", "")
                    if not win_title: raise ValueError("Window Title is required for Integrated Capture")
                    
                    self._log(f"Searching for window: '{win_title}'")
                    wins = gw.getWindowsWithTitle(win_title)
                    if not wins: raise RuntimeError(f"Could not find window with title: {win_title}")
                    
                    win = wins[0]
                    self._log(f"Activating window: {win.title}")
                    try:
                        if win.isMinimized: win.restore()
                        win.activate()
                        time.sleep(1.0)
                    except: pass
                    
                    temp_dir = os.path.join(os.path.dirname(__file__), "..", "temps")
                    os.makedirs(temp_dir, exist_ok=True)
                    file_path = os.path.join(temp_dir, f"detect_capture_{os.urandom(2).hex()}.png")
                    
                    # Ensure DPI awareness
                    try:
                        import ctypes
                        ctypes.windll.shcore.SetProcessDpiAwareness(1)
                    except: pass
                    
                    # Capture the window area more robustly
                    # Maximized windows often have a -8, -8 offset with extra border
                    left, top, width, height = win.left, win.top, win.width, win.height
                    if win.isMaximized:
                        # Trim the invisible borders typical of maximized windows
                        left += 8
                        top += 8
                        width -= 16
                        height -= 16
                    
                    self._log(f"Capturing region: L:{left}, T:{top}, W:{width}, H:{height}")
                    
                    # Use ImageGrab directly as it's often more reliable with DPI on Windows
                    # Capture the absolute screen coordinates in full physical resolution
                    bbox = (max(0, left), max(0, top), left + width, top + height)
                    shot = ImageGrab.grab(bbox=bbox)
                    
                    # DO NOT RESIZE. Sending the full resolution image to the AI 
                    # provides the best quality for OCR and object detection.
                    # Our click utility in Gui_Automate already handles scaling 
                    # by comparing AI image_size vs screen resolution.
                    
                    shot.save(file_path)
                    self._log(f"Captured high-res screenshot: {file_path} ({shot.width}x{shot.height})")

                if not file_path or not os.path.exists(file_path): raise ValueError(f"File not found: {file_path}")
                data = {"prompt": str(prompt)} if prompt else {}
                with open(file_path, "rb") as f:
                    files = {"file": (os.path.basename(file_path), f)}
                    response = requests.post(full_url, data=data, files=files, timeout=300)
                response.raise_for_status()
                ai_json = response.json()
                
                # The server now returns pixel-perfect bbox_pixels mapped back 
                # to the original image dimensions [xmin, ymin, xmax, ymax].
                
                # --- SIMULATION MODE (DRAW RECTANGLES) ---
                if config_data.get("simulation") and "elements" in ai_json:
                    try:
                        temp_dir = os.path.join(os.path.dirname(__file__), "..", "temps")
                        os.makedirs(temp_dir, exist_ok=True)
                        debug_path = os.path.join(temp_dir, "debug_detect.png")
                        
                        with Image.open(file_path) as img:
                            # Convert to RGB if necessary for drawing
                            if img.mode != "RGB":
                                img = img.convert("RGB")
                            draw = ImageDraw.Draw(img)
                            for el in ai_json["elements"]:
                                pix = el.get("bbox_pixels")
                                if pix:
                                    draw.rectangle([pix["xmin"], pix["ymin"], pix["xmax"], pix["ymax"]], outline="red", width=3)
                            
                            img.save(debug_path)
                            self._log(f"Simulation image saved: {debug_path}")
                    except Exception as de:
                        self._log(f"Simulation Error: {de}")

                result_data = json.dumps(ai_json, ensure_ascii=False, indent=2)

            elif endpoint == "/prompt":
                prompt_text = str(prompt)
                secondary_text = secondary
                response = requests.post(full_url, json={"word": prompt_text, "translation": secondary_text}, timeout=300)
                response.raise_for_status()
                result_data = json.dumps(response.json(), ensure_ascii=False, indent=2)

            # --- RAG Model Endpoints (rag_model.py, default port 8001) ---
            elif endpoint == "/classify":
                if not prompt: raise ValueError("Product Info / Description is required for /classify")
                payload = {"product_info": str(prompt)}
                response = requests.post(full_url, json=payload, timeout=300)
                response.raise_for_status()
                result_data = json.dumps(response.json(), ensure_ascii=False, indent=2)

            elif endpoint == "/ingest-history":
                if not file_path or not os.path.exists(file_path):
                    raise ValueError(f"History file not found: {file_path}")

                chunk_size = config_data.get("chunk_size", 0)

                # --- CHUNKED UPLOAD ---
                if chunk_size and chunk_size > 0:
                    fname = os.path.basename(file_path).lower()
                    if fname.endswith(".xlsx"):
                        df_full = pd.read_excel(file_path)
                    elif fname.endswith(".csv"):
                        df_full = pd.read_csv(file_path)
                    else:
                        raise ValueError("Chunk mode only supports .xlsx or .csv files.")

                    total_rows = len(df_full)
                    total_chunks = (total_rows + chunk_size - 1) // chunk_size
                    self._log(f"Chunked upload: {total_rows} rows → {total_chunks} chunks of {chunk_size} rows each.")

                    temp_dir = os.path.join(os.path.dirname(__file__), "..", "temps")
                    os.makedirs(temp_dir, exist_ok=True)

                    total_ingested = 0
                    chunk_results = []
                    for chunk_idx in range(total_chunks):
                        chunk_df = df_full.iloc[chunk_idx * chunk_size : (chunk_idx + 1) * chunk_size]
                        chunk_file = os.path.join(temp_dir, f"ingest_chunk_{os.urandom(3).hex()}.csv")
                        chunk_df.to_csv(chunk_file, index=False)
                        self._log(f"Sending chunk {chunk_idx + 1}/{total_chunks} ({len(chunk_df)} rows)...")
                        try:
                            with open(chunk_file, "rb") as f:
                                files = {"file": (os.path.basename(chunk_file), f, "text/csv")}
                                resp = requests.post(full_url, files=files, timeout=600)
                            resp.raise_for_status()
                            resp_json = resp.json()
                            chunk_results.append(resp_json.get("message", str(resp_json)))
                            total_ingested += len(chunk_df)
                            self._log(f"Chunk {chunk_idx + 1} done: {resp_json.get('message', 'OK')}")
                        finally:
                            try: os.remove(chunk_file)
                            except: pass

                    result_data = json.dumps({
                        "total_rows_sent": total_ingested,
                        "chunks": total_chunks,
                        "chunk_size": chunk_size,
                        "results": chunk_results
                    }, ensure_ascii=False, indent=2)

                # --- SINGLE UPLOAD (no chunking) ---
                else:
                    with open(file_path, "rb") as f:
                        files = {"file": (os.path.basename(file_path), f)}
                        response = requests.post(full_url, files=files, timeout=600)
                    response.raise_for_status()
                    result_data = json.dumps(response.json(), ensure_ascii=False, indent=2)

            elif endpoint == "/ingest":
                if not file_path or not os.path.exists(file_path):
                    raise ValueError(f"Document file not found: {file_path}")
                with open(file_path, "rb") as f:
                    files = {"file": (os.path.basename(file_path), f)}
                    response = requests.post(full_url, files=files, timeout=300)
                response.raise_for_status()
                result_data = json.dumps(response.json(), ensure_ascii=False, indent=2)

            elif endpoint == "/health":
                response = requests.get(full_url, timeout=30)
                response.raise_for_status()
                result_data = json.dumps(response.json(), ensure_ascii=False, indent=2)

            elif endpoint == "/reset":
                response = requests.post(full_url, timeout=60)
                response.raise_for_status()
                result_data = json.dumps(response.json(), ensure_ascii=False, indent=2)

            else:
                raise ValueError(f"Unknown endpoint: {endpoint}")

            self._log("Successfully received response from Local AI.")

        except Exception as e:
            error_message = f"FATAL ERROR during Local AI call: {e}"
            self._log(error_message)
            raise RuntimeError(error_message)

        # Return result as a DataFrame
        df = pd.DataFrame([str(result_data)], columns=['Response'])
        return df
