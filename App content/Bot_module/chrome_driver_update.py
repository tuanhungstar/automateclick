# File: Bot_module/Chrome_Driver_Updater.py

import sys
import os
import shutil
import platform
import json
import requests
import zipfile
import io
import subprocess
from typing import Optional, List, Dict, Any

# --- PyQt6 Imports ---
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QDialogButtonBox,
    QWidget, QGroupBox, QMessageBox, QLabel, QHBoxLayout, QProgressBar
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal

# --- Main App Imports (Fallback) ---
try:
    from my_lib.shared_context import ExecutionContext
except ImportError:
    class ExecutionContext:
        def add_log(self, message: str): print(message)
        def get_variable(self, name: str): return None

#
# --- HELPER: Logic to detect Chrome Version ---
#
def get_chrome_version() -> Optional[str]:
    """
    Detects the installed Google Chrome version based on the OS.
    Returns the version string (e.g., '120.0.6099.109') or None if not found.
    """
    system = platform.system()
    version = None

    try:
        if system == "Windows":
            # fast method using registry via command line
            cmd = 'reg query "HKEY_CURRENT_USER\\Software\\Google\\Chrome\\BLBeacon" /v version'
            output = subprocess.check_output(cmd, shell=True).decode()
            version = output.strip().split()[-1]
        
        elif system == "Darwin": # MacOS
            cmd = r'/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --version'
            output = subprocess.check_output(cmd, shell=True).decode()
            version = output.strip().split()[-1] # Format: "Google Chrome 120.0.x.x"
            
        elif system == "Linux":
            cmd = 'google-chrome --version'
            output = subprocess.check_output(cmd, shell=True).decode()
            version = output.strip().split()[-1]
            
    except Exception:
        # Fallback for Windows if registry fails: try wmic
        if system == "Windows":
            try:
                cmd = r'wmic datafile where name="C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" get Version /value'
                output = subprocess.check_output(cmd, shell=True).decode()
                # output looks like "Version=120.0.xxxx"
                for line in output.split('\n'):
                    if "Version=" in line:
                        version = line.split('=')[1].strip()
            except Exception:
                pass

    return version

#
# --- HELPER: Worker thread for Downloading Driver ---
#
class _DriverDownloaderThread(QThread):
    progress_update = pyqtSignal(int, str)
    finished_signal = pyqtSignal(str) # Path to downloaded driver
    error_signal = pyqtSignal(str)

    def __init__(self, config: dict, context: Optional[ExecutionContext] = None):
        super().__init__()
        self.config = config
        self.context = context

    def _log(self, msg: str):
        if self.context: self.context.add_log(msg)
        else: print(msg)

    def run(self):
        try:
            target_folder = self.config['target_folder']
            forced_version = self.config.get('forced_version', '')

            # FIX 1: Ensure target directory exists before doing anything
            if not os.path.exists(target_folder):
                try:
                    os.makedirs(target_folder, exist_ok=True)
                except OSError as e:
                    raise Exception(f"Could not create folder '{target_folder}'. Error: {e}")
            
            # 1. Determine Chrome Version
            chrome_version = forced_version
            if not chrome_version:
                self.progress_update.emit(10, "Detecting installed Chrome version...")
                chrome_version = get_chrome_version()
                
            if not chrome_version:
                raise Exception("Could not detect Google Chrome version. Is it installed?")
            
            self._log(f"Target Chrome Version: {chrome_version}")
            
            # 2. Determine Platform string for Google API
            system = platform.system()
            is_64bits = sys.maxsize > 2**32
            
            platform_str = ""
            if system == "Windows":
                platform_str = "win64" if is_64bits else "win32"
            elif system == "Darwin":
                platform_str = "mac-arm64" if platform.machine() == 'arm64' else "mac-x64"
            elif system == "Linux":
                platform_str = "linux64"
            
            if not platform_str:
                raise Exception(f"Unsupported Operating System: {system}")

            # 3. Fetch JSON metadata from Chrome for Testing API
            self.progress_update.emit(30, "Fetching driver metadata from Google...")
            # Note: This logic handles modern Chrome (v115+)
            json_url = "https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json"
            
            resp = requests.get(json_url, timeout=15)
            if resp.status_code != 200:
                raise Exception("Failed to connect to Google Chrome Driver API.")
            
            data = resp.json()
            
            # Logic: Try to match the Major version
            major_version = chrome_version.split('.')[0]
            download_url = ""
            
            # Check "channels" (Stable, Beta, etc.) first for a quick match
            channels = data.get('channels', {})
            if 'Stable' in channels and channels['Stable']['version'].split('.')[0] == major_version:
                 for platform_entry in channels['Stable']['downloads']['chromedriver']:
                     if platform_entry['platform'] == platform_str:
                         download_url = platform_entry['url']
                         break
            
            # Fallback: Just assume the user wants the latest stable driver available
            if not download_url:
                self._log("Exact version match not found in Stable. Using latest available Stable driver.")
                for platform_entry in channels['Stable']['downloads']['chromedriver']:
                     if platform_entry['platform'] == platform_str:
                         download_url = platform_entry['url']
                         break

            if not download_url:
                 raise Exception(f"Could not find a driver for platform '{platform_str}' and version '{major_version}'.")

            self._log(f"Download URL found: {download_url}")

            # 4. Download Zip
            self.progress_update.emit(50, "Downloading driver...")
            r = requests.get(download_url, stream=True)
            total_size = int(r.headers.get('content-length', 0))
            block_size = 1024
            downloaded = 0
            
            zip_in_memory = io.BytesIO()
            for data in r.iter_content(block_size):
                zip_in_memory.write(data)
                downloaded += len(data)
                if total_size > 0:
                    percent = 50 + int((downloaded / total_size) * 30)
                    self.progress_update.emit(percent, "Downloading...")

            # 5. Extract and Install
            self.progress_update.emit(85, "Extracting file...")
            found_binary = False
            
            with zipfile.ZipFile(zip_in_memory) as z:
                for member in z.infolist():
                    # FIX 2: Skip directory entries explicitly
                    if member.filename.endswith('/'):
                        continue
                    
                    # FIX 3: Check only the filename, ignoring the folders inside the zip
                    base_name = os.path.basename(member.filename)
                    
                    # Check for binary name (Windows: chromedriver.exe, Unix: chromedriver)
                    if base_name.lower() in ['chromedriver.exe', 'chromedriver']:
                        target_path = os.path.join(target_folder, base_name)
                        
                        self._log(f"Extracting {member.filename} to {target_path}...")
                        
                        try:
                            with z.open(member) as source, open(target_path, "wb") as target:
                                shutil.copyfileobj(source, target)
                        except PermissionError:
                             raise Exception(f"Permission Denied! Cannot write to '{target_folder}'. Please try selecting a different folder (e.g., Documents or Desktop).")

                        # Make executable on Unix/Mac
                        if system != "Windows":
                            st = os.stat(target_path)
                            os.chmod(target_path, st.st_mode | 0o111)
                            
                        self.finished_signal.emit(target_path)
                        found_binary = True
                        break # Stop after finding the binary

            if not found_binary:
                raise Exception("The downloaded zip file did not contain a 'chromedriver' executable.")

        except Exception as e:
            self.error_signal.emit(str(e))

#
# --- HELPER: The GUI Dialog ---
#
class _DriverUpdateDialog(QDialog):
    def __init__(self, parent: Optional[QWidget] = None, initial_config: Optional[Dict[str, Any]] = None):
        super().__init__(parent)
        self.setWindowTitle("Get/Update Chrome Driver")
        self.setMinimumWidth(500)
        
        main_layout = QVBoxLayout(self)

        # 1. Info Group
        info_group = QGroupBox("System Information")
        info_layout = QFormLayout(info_group)
        
        self.detected_version = get_chrome_version()
        self.version_label = QLabel(self.detected_version if self.detected_version else "Could not detect automatically")
        if not self.detected_version:
            self.version_label.setStyleSheet("color: red")
            
        info_layout.addRow("Detected Chrome Version:", self.version_label)
        main_layout.addWidget(info_group)

        # 2. Destination
        dest_group = QGroupBox("Destination")
        dest_layout = QFormLayout(dest_group)
        
        self.folder_path_edit = QLineEdit()
        self.folder_path_edit.setReadOnly(True)
        browse_button = QPushButton("Browse Folder...")
        path_layout = QHBoxLayout(); path_layout.addWidget(self.folder_path_edit); path_layout.addWidget(browse_button)
        
        dest_layout.addRow("Save Driver To:", path_layout)
        main_layout.addWidget(dest_group)

        # Dialog Buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # Connections
        browse_button.clicked.connect(self._browse_folder)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        # Init
        if initial_config:
            self.folder_path_edit.setText(initial_config.get("target_folder", ""))
        
        # Default to current directory if empty
        if not self.folder_path_edit.text():
            self.folder_path_edit.setText(os.getcwd())

    def _browse_folder(self):
        from PyQt6.QtWidgets import QFileDialog
        folder = QFileDialog.getExistingDirectory(self, "Select Folder to Save ChromeDriver")
        if folder:
            self.folder_path_edit.setText(folder)

    def get_executor_method_name(self) -> str:
        return "_get_driver"
    
    def get_assignment_variable(self) -> Optional[str]:
        return None 

    def get_config_data(self) -> Optional[Dict[str, Any]]:
        target_folder = self.folder_path_edit.text()
        if not target_folder:
             QMessageBox.warning(self, "Input Error", "Please select a destination folder.")
             return None
        
        return {
            "target_folder": target_folder,
            "forced_version": self.detected_version
        }

#
# --- The Public-Facing Module Class ---
#
class Get_update_chrome_driver:
    """
    Checks the installed Google Chrome version and downloads the matching 
    ChromeDriver binary from Google's Chrome for Testing API.
    """
    def __init__(self, context: Optional[ExecutionContext] = None):
        self.context = context
        self.worker: Optional[_DriverDownloaderThread] = None

    def _log(self, message: str):
        if self.context: self.context.add_log(message)
        else: print(message)

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        self._log("Opening Chrome Driver Updater configuration...")
        initial_config = kwargs.get("initial_config")
        return _DriverUpdateDialog(parent=parent_window, initial_config=initial_config)

    def _get_driver(self, context: ExecutionContext, config_data: dict):
        """
        Executes the download process.
        """
        from PyQt6.QtCore import QEventLoop
        
        self.context = context
        self.worker = _DriverDownloaderThread(config_data, context)
        
        loop = QEventLoop()
        final_path = None
        error_msg = ""
        
        def on_finished(path):
            nonlocal final_path
            final_path = path
            loop.quit()
            
        def on_error(msg):
            nonlocal error_msg
            error_msg = msg
            loop.quit()
            
        if hasattr(context, 'update_progress_bar'):
            self.worker.progress_update.connect(context.update_progress_bar)
        else:
            self.worker.progress_update.connect(lambda p, m: self._log(f"Progress {p}%: {m}"))
            
        self.worker.finished_signal.connect(on_finished)
        self.worker.error_signal.connect(on_error)
        
        self._log("Starting Chrome Driver download...")
        self.worker.start()
        loop.exec()
        
        if error_msg:
            raise Exception(f"Driver Update Failed: {error_msg}")
            
        self._log(f"SUCCESS: ChromeDriver saved to: {final_path}")
        return final_path