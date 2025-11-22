# File: Bot_module/outlook_module.py

import sys
import os
import re
from datetime import datetime, timedelta
from typing import Optional, List, Dict, Any
import pandas as pd
import fnmatch # For wildcard matching

# --- PyQt6 Imports ---
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QDialogButtonBox,
    QComboBox, QWidget, QGroupBox, QMessageBox, QLabel,
    QCheckBox, QApplication, QTreeWidget, QTreeWidgetItem,
    QFileDialog, QDateTimeEdit, QHBoxLayout, QTreeWidgetItemIterator, QRadioButton,
    QPlainTextEdit
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QDateTime

# --- COM Imports for Outlook ---
try:
    import win32com.client
except ImportError:
    print("Warning: 'pywin32' library not found. Please install it using: pip install pywin32")
    win32com = None

# --- Main App Imports (Fallback for standalone testing) ---
try:
    from my_lib.shared_context import ExecutionContext
except ImportError:
    print("Warning: Could not import main app libraries. Using fallbacks.")
    class ExecutionContext:
        def add_log(self, message: str): print(message)
        def get_variable(self, name: str): return None

#
# --- HELPER: Worker thread for fetching Outlook folders ---
#
class _OutlookFolderLoaderThread(QThread):
    """Worker thread to fetch Outlook folders without freezing the GUI."""
    folders_ready = pyqtSignal(list)
    error_occurred = pyqtSignal(str)

    def run(self):
        try:
            if not win32com:
                raise ImportError("'pywin32' library is not installed.")
            
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            # This will access the default mailbox. For multiple accounts, more complex logic is needed.
            root_folder = outlook.Folders.GetFirst() 
            
            folder_list = []
            # Recursive function to traverse all folders
            def recurse_folders(folder):
                # Use FullFolderPath for a unique, reliable path that win32com can parse
                path = folder.FullFolderPath
                folder_list.append(path)
                for subfolder in folder.Folders:
                    recurse_folders(subfolder)
            
            # Start recursion from the root of the mailbox
            for folder in outlook.Folders:
                 recurse_folders(folder)

            self.folders_ready.emit(folder_list)
        except Exception as e:
            self.error_occurred.emit(f"Could not access Outlook. Ensure it is installed and configured.\n\nError: {e}")

#
# --- HELPER: The GUI Dialog for Reading Outlook Mail ---
#
class _ReadOutlookDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None,
                 initial_variable: Optional[str] = None):
        super().__init__(parent)
        self.setWindowTitle("Read Outlook Emails")
        self.setMinimumSize(700, 800)
        self.global_variables = global_variables
        self.initial_config = initial_config or {}

        main_layout = QVBoxLayout(self)

        # 1. Folder Selection
        folder_group = QGroupBox("Select Folders to Read")
        folder_layout = QVBoxLayout(folder_group)
        self.load_folders_button = QPushButton("Load Outlook Folders")
        self.folder_tree = QTreeWidget()
        self.folder_tree.setHeaderHidden(True)
        
        self.selected_folders_display = QLineEdit()
        self.selected_folders_display.setReadOnly(True)
        self.selected_folders_display.setPlaceholderText("Selected folder paths will appear here...")
        
        folder_layout.addWidget(self.load_folders_button)
        folder_layout.addWidget(self.folder_tree)
        folder_layout.addWidget(QLabel("Selected Folders:"))
        folder_layout.addWidget(self.selected_folders_display)
        main_layout.addWidget(folder_group)

        # 2. Filter by Time and Status
        time_group = QGroupBox("Filter by Time and Status")
        time_layout = QFormLayout(time_group)
        self.from_date_edit = QDateTimeEdit(QDateTime.currentDateTime().addDays(-7)); self.from_date_edit.setCalendarPopup(True)
        self.to_date_edit = QDateTimeEdit(QDateTime.currentDateTime()); self.to_date_edit.setCalendarPopup(True)
        self.use_current_datetime_check = QCheckBox("Use Current Datetime for To (End Date)")
        self.status_combo = QComboBox(); self.status_combo.addItems(["All Emails", "Unread Only", "Read Only"])
        time_layout.addRow("From (Start Date):", self.from_date_edit)
        to_layout = QHBoxLayout(); to_layout.addWidget(self.to_date_edit); to_layout.addWidget(self.use_current_datetime_check)
        time_layout.addRow("To (End Date):", to_layout)
        time_layout.addRow("Status:", self.status_combo)
        main_layout.addWidget(time_group)

        # 3. Filter by Content
        content_group = QGroupBox("Filter by Content (Use * as wildcard)")
        content_layout = QFormLayout(content_group)
        self.sender_edit = QLineEdit(); self.sender_edit.setPlaceholderText("e.g., *@example.com or user@*")
        self.subject_edit = QLineEdit(); self.subject_edit.setPlaceholderText("e.g., *invoice* or Report*")
        self.body_edit = QLineEdit(); self.body_edit.setPlaceholderText("e.g., *payment due*")
        self.attachment_name_edit = QLineEdit(); self.attachment_name_edit.setPlaceholderText("e.g., *.pdf or *data.xlsx")
        content_layout.addRow("From Sender:", self.sender_edit); content_layout.addRow("Subject:", self.subject_edit)
        content_layout.addRow("Body contains:", self.body_edit); content_layout.addRow("Attachment Name:", self.attachment_name_edit)
        main_layout.addWidget(content_group)

        # 4. Attachment Options
        attachment_group = QGroupBox("Attachment Options")
        attachment_layout = QFormLayout(attachment_group)
        self.save_attachments_check = QCheckBox("Save attachments from matching emails")
        self.save_location_edit = QLineEdit(); self.save_location_edit.setReadOnly(True)
        browse_button = QPushButton("Browse...")
        location_layout = QHBoxLayout(); location_layout.addWidget(self.save_location_edit); location_layout.addWidget(browse_button)
        attachment_layout.addRow(self.save_attachments_check); attachment_layout.addRow("Save Location:", location_layout)
        main_layout.addWidget(attachment_group)

        # 5. Assign Results
        assign_group = QGroupBox("Assign Results to Variable")
        assign_layout = QFormLayout(assign_group)
        self.assign_results_check = QCheckBox("Assign results (DataFrame) to a variable")
        self.new_var_radio = QRadioButton("New Variable Name:"); self.new_var_input = QLineEdit("outlook_emails")
        self.existing_var_radio = QRadioButton("Existing Variable:"); self.existing_var_combo = QComboBox(); self.existing_var_combo.addItems(["-- Select --"] + global_variables)
        assign_layout.addRow(self.assign_results_check); assign_layout.addRow(self.new_var_radio, self.new_var_input); assign_layout.addRow(self.existing_var_radio, self.existing_var_combo)
        main_layout.addWidget(assign_group)
        
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel); main_layout.addWidget(self.button_box)

        # Connections
        self.load_folders_button.clicked.connect(self._load_folders)
        self.folder_tree.itemChanged.connect(self._update_selected_folders_display)
        self.use_current_datetime_check.toggled.connect(self.to_date_edit.setDisabled)
        self.save_attachments_check.toggled.connect(self._toggle_attachment_location)
        browse_button.clicked.connect(self._browse_save_location)
        self.assign_results_check.toggled.connect(self._toggle_assignment_widgets)
        self.button_box.accepted.connect(self.accept); self.button_box.rejected.connect(self.reject)

        self._populate_from_initial_config(initial_config, initial_variable)
        self._toggle_attachment_location(self.save_attachments_check.isChecked())
        self._toggle_assignment_widgets(self.assign_results_check.isChecked())
        
    def _load_folders(self):
        self.load_folders_button.setText("Loading..."); self.load_folders_button.setEnabled(False)
        self.folder_loader_thread = _OutlookFolderLoaderThread()
        self.folder_loader_thread.folders_ready.connect(self._on_folders_loaded)
        self.folder_loader_thread.error_occurred.connect(lambda e: (QMessageBox.critical(self, "Error", e), self.load_folders_button.setText("Load Outlook Folders"), self.load_folders_button.setEnabled(True)))
        self.folder_loader_thread.start()

    def _on_folders_loaded(self, folder_paths: List[str]):
        self.folder_tree.itemChanged.disconnect(self._update_selected_folders_display)
        self.folder_tree.clear()
        
        items = {}
        for path in sorted(folder_paths):
            parts = path.strip('\\').split('\\')
            parent = self.folder_tree.invisibleRootItem()
            item_path_key = ""
            for part in parts:
                item_path_key += f"\\{part}"
                if item_path_key not in items:
                    item = QTreeWidgetItem(parent, [part])
                    item.setData(0, Qt.ItemDataRole.UserRole, path)
                    item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                    item.setCheckState(0, Qt.CheckState.Unchecked)
                    items[item_path_key] = item
                parent = items[item_path_key]
        
        self.load_folders_button.setText("Reload Folders"); self.load_folders_button.setEnabled(True)
        self._restore_folder_selection()
        self.folder_tree.itemChanged.connect(self._update_selected_folders_display)
        self._update_selected_folders_display()

    def _restore_folder_selection(self):
        selected_paths = self.initial_config.get("folders", [])
        if not selected_paths: return
        iterator = QTreeWidgetItemIterator(self.folder_tree, QTreeWidgetItemIterator.IteratorFlag.All)
        while iterator.value():
            item = iterator.value()
            item_path = item.data(0, Qt.ItemDataRole.UserRole)
            if item_path in selected_paths:
                item.setCheckState(0, Qt.CheckState.Checked)
            iterator += 1

    def _update_selected_folders_display(self):
        checked_paths = []
        iterator = QTreeWidgetItemIterator(self.folder_tree, QTreeWidgetItemIterator.IteratorFlag.All)
        while iterator.value():
            item = iterator.value()
            if item.checkState(0) == Qt.CheckState.Checked:
                path = item.data(0, Qt.ItemDataRole.UserRole)
                if path: checked_paths.append(path)
            iterator += 1
        self.selected_folders_display.setText(", ".join(checked_paths))

    def _toggle_attachment_location(self, checked):
        self.save_location_edit.setEnabled(checked); self.save_location_edit.parent().findChild(QPushButton).setEnabled(checked)
    def _browse_save_location(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder to Save Attachments");
        if folder: self.save_location_edit.setText(folder)
    def _toggle_assignment_widgets(self, checked):
        self.new_var_radio.setEnabled(checked); self.new_var_input.setEnabled(checked)
        self.existing_var_radio.setEnabled(checked); self.existing_var_combo.setEnabled(checked)
    def _populate_from_initial_config(self, config, variable):
        if not config:
            self.assign_results_check.setChecked(True); self.new_var_radio.setChecked(True); self.use_current_datetime_check.setChecked(True)
            return
        if self.initial_config.get("folders"): self._load_folders()
        self.from_date_edit.setDateTime(QDateTime.fromString(config.get("from_date", QDateTime.currentDateTime().addDays(-7).toString(Qt.DateFormat.ISODate)), Qt.DateFormat.ISODate))
        self.to_date_edit.setDateTime(QDateTime.fromString(config.get("to_date", QDateTime.currentDateTime().toString(Qt.DateFormat.ISODate)), Qt.DateFormat.ISODate))
        self.use_current_datetime_check.setChecked(config.get("use_current_datetime", True))
        self.status_combo.setCurrentText(config.get("status", "All Emails"))
        self.sender_edit.setText(config.get("sender", "")); self.subject_edit.setText(config.get("subject", ""))
        self.body_edit.setText(config.get("body", "")); self.attachment_name_edit.setText(config.get("attachment_name", ""))
        self.save_attachments_check.setChecked(config.get("save_attachments", False))
        self.save_location_edit.setText(config.get("save_location", ""))
        self.assign_results_check.setChecked(bool(variable))
        if variable:
            if variable in self.global_variables: self.existing_var_radio.setChecked(True); self.existing_var_combo.setCurrentText(variable)
            else: self.new_var_radio.setChecked(True); self.new_var_input.setText(variable)
        else: self.new_var_radio.setChecked(True)
    def get_executor_method_name(self) -> str: return "_read_outlook_emails"
    def get_config_data(self) -> Optional[Dict[str, Any]]:
        folder_paths_str = self.selected_folders_display.text()
        if not folder_paths_str:
            QMessageBox.warning(self, "Input Error", "Please load and select at least one Outlook folder."); return None
        selected_folders = [path.strip() for path in folder_paths_str.split(',') if path.strip()]
        config = {
            "folders": selected_folders,
            "from_date": self.from_date_edit.dateTime().toString(Qt.DateFormat.ISODate),
            "to_date": self.to_date_edit.dateTime().toString(Qt.DateFormat.ISODate),
            "use_current_datetime": self.use_current_datetime_check.isChecked(),
            "status": self.status_combo.currentText(),
            "sender": self.sender_edit.text().strip(), "subject": self.subject_edit.text().strip(),
            "body": self.body_edit.text().strip(), "attachment_name": self.attachment_name_edit.text().strip(),
            "save_attachments": self.save_attachments_check.isChecked(), "save_location": self.save_location_edit.text()
        }
        if config["save_attachments"] and not config["save_location"]:
            QMessageBox.warning(self, "Input Error", "Please select a save location for attachments."); return None
        return config
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
# --- The Public-Facing Module Class for Reading Outlook Mail ---
#
class Read_Outlook_Mail:
    def __init__(self, context: Optional[ExecutionContext] = None): self.context = context
    def _log(self, m: str): (self.context.add_log(m) if self.context else print(m))

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        self._log("Opening Read Outlook Emails configuration...")
        return _ReadOutlookDialog(global_variables, parent_window, **kwargs)

    def _read_outlook_emails(self, context: ExecutionContext, config_data: dict) -> pd.DataFrame:
        self.context = context
        if not win32com: raise ImportError("'pywin32' library is not installed.")

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        from_date = QDateTime.fromString(config_data["from_date"], Qt.DateFormat.ISODate).toPyDateTime()
        to_date = datetime.now() if config_data["use_current_datetime"] else QDateTime.fromString(config_data["to_date"], Qt.DateFormat.ISODate).toPyDateTime()

        self._log(f"Searching emails from {from_date.strftime('%Y-%m-%d %H:%M')} to {to_date.strftime('%Y-%m-%d %H:%M')}")
        all_emails_data = []
        
        for folder_path in config_data["folders"]:
            folder = None
            try:
                # Resolve folder path using win32com's hierarchy
                path_parts = folder_path.strip('\\').split('\\')
                folder = outlook.Folders[path_parts[0]]
                for part in path_parts[1:]:
                    folder = folder.Folders[part]
                self._log(f"Processing folder: {folder.Name}")
            except Exception:
                self._log(f"Warning: Could not find folder path: {folder_path}. Skipping."); continue

            filter_str = f"[ReceivedTime] >= '{from_date.strftime('%m/%d/%Y %H:%M %p')}' AND [ReceivedTime] <= '{to_date.strftime('%m/%d/%Y %H:%M %p')}'"
            if config_data["status"] == "Unread Only": filter_str += " AND [Unread] = true"
            elif config_data["status"] == "Read Only": filter_str += " AND [Unread] = false"

            try:
                emails = folder.Items.Restrict(filter_str)
                emails.Sort("[ReceivedTime]", True)
            except Exception as e:
                self._log(f"Warning: Error filtering items in folder '{folder.Name}': {e}"); continue
            
            for email in emails:
                try:
                    sender = email.SenderName if email.SenderName else ""
                    subject = email.Subject if email.Subject else ""
                    body = email.Body if email.Body else ""

                    if config_data["sender"] and not fnmatch.fnmatch(sender.lower(), config_data["sender"].lower()): continue
                    if config_data["subject"] and not fnmatch.fnmatch(subject.lower(), config_data["subject"].lower()): continue
                    if config_data["body"] and not fnmatch.fnmatch(body.lower(), f'*{config_data["body"].lower()}*'): continue

                    has_matching_attachment, attachment_filenames = False, []
                    if email.Attachments.Count > 0:
                        for attachment in email.Attachments:
                            attachment_filenames.append(attachment.FileName)
                            if config_data["attachment_name"] and fnmatch.fnmatch(attachment.FileName.lower(), config_data["attachment_name"].lower()):
                                has_matching_attachment = True
                    
                    if config_data["attachment_name"] and not has_matching_attachment: continue
                    
                    saved_attachment_paths = []
                    if config_data["save_attachments"] and email.Attachments.Count > 0:
                        save_location = config_data["save_location"]
                        os.makedirs(save_location, exist_ok=True)
                        for attachment in email.Attachments:
                            if not config_data["attachment_name"] or fnmatch.fnmatch(attachment.FileName.lower(), config_data["attachment_name"].lower()):
                                file_path = os.path.join(save_location, attachment.FileName)
                                try:
                                    attachment.SaveAsFile(file_path)
                                    saved_attachment_paths.append(file_path)
                                except Exception as save_err:
                                    self._log(f"Warning: Could not save attachment '{attachment.FileName}'. Error: {save_err}")
                    
                    email_data = {
                        "Subject": subject, "Sender": sender,
                        "ReceivedTime": email.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S"),
                        "Body": body, "To": email.To, "CC": email.CC,
                        "Attachments": ", ".join(attachment_filenames),
                        "SavedAttachments": ", ".join(saved_attachment_paths)
                    }
                    all_emails_data.append(email_data)
                except Exception as e:
                    self._log(f"Warning: Could not process an email. Subject: '{getattr(email, 'Subject', 'N/A')}'. Error: {e}")

        self._log(f"Found {len(all_emails_data)} matching emails across all selected folders.")
        return pd.DataFrame(all_emails_data)

#
# --- HELPER: The GUI Dialog for Sending Outlook Mail ---
#
class _SendOutlookDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None, **kwargs):
        super().__init__(parent)
        self.setWindowTitle("Send Outlook Email")
        self.setMinimumSize(600, 500)
        self.global_variables = ["-- Select Variable --"] + global_variables
        
        main_layout = QVBoxLayout(self)

        # Helper function to create a row with a text input and a variable selector
        def create_input_row(label_text):
            layout = QHBoxLayout()
            line_edit = QLineEdit()
            combo_box = QComboBox()
            combo_box.addItems(self.global_variables)
            layout.addWidget(line_edit)
            layout.addWidget(QLabel("or from variable:"))
            layout.addWidget(combo_box)
            # When a variable is selected, disable the line edit
            combo_box.currentTextChanged.connect(
                lambda text, le=line_edit: le.setDisabled(text != "-- Select Variable --")
            )
            return layout, line_edit, combo_box

        form_layout = QFormLayout()
        
        # To, CC, BCC, Subject
        self.to_layout, self.to_edit, self.to_combo = create_input_row("To:")
        self.cc_layout, self.cc_edit, self.cc_combo = create_input_row("CC:")
        self.bcc_layout, self.bcc_edit, self.bcc_combo = create_input_row("BCC:")
        self.subject_layout, self.subject_edit, self.subject_combo = create_input_row("Subject:")
        
        form_layout.addRow("To:", self.to_layout)
        form_layout.addRow("CC (optional):", self.cc_layout)
        form_layout.addRow("BCC (optional):", self.bcc_layout)
        form_layout.addRow("Subject:", self.subject_layout)
        
        # Body
        body_layout = QVBoxLayout()
        self.body_edit = QPlainTextEdit()
        self.body_edit.setPlaceholderText("Type email body here...")
        body_var_layout = QHBoxLayout()
        body_var_layout.addWidget(QLabel("... or use body from variable:"))
        self.body_combo = QComboBox()
        self.body_combo.addItems(self.global_variables)
        body_var_layout.addWidget(self.body_combo)
        body_var_layout.addStretch()
        body_layout.addWidget(self.body_edit)
        body_layout.addLayout(body_var_layout)
        self.body_combo.currentTextChanged.connect(
            lambda text: self.body_edit.setDisabled(text != "-- Select Variable --")
        )
        form_layout.addRow("Body:", body_layout)
        
        # --- MODIFIED SECTION START ---
        # Attachments with a browse button
        self.attach_edit = QLineEdit()
        self.attach_edit.setPlaceholderText("C:\\path\\file.txt;C:\\path\\report.xlsx")
        self.attach_combo = QComboBox()
        self.attach_combo.addItems(self.global_variables)
        
        attach_browse_button = QPushButton("Browse...")
        attach_browse_button.clicked.connect(self._browse_for_attachments)

        attach_layout = QHBoxLayout()
        attach_layout.addWidget(self.attach_edit)
        attach_layout.addWidget(attach_browse_button)
        attach_layout.addWidget(QLabel("or from variable:"))
        attach_layout.addWidget(self.attach_combo)

        # Disable the text edit and browse button if a variable is selected
        self.attach_combo.currentTextChanged.connect(
            lambda text: (
                self.attach_edit.setDisabled(text != "-- Select Variable --"),
                attach_browse_button.setDisabled(text != "-- Select Variable --")
            )
        )
        
        form_layout.addRow("Attachments (optional, semi-colon separated):", attach_layout)
        # --- MODIFIED SECTION END ---
        
        main_layout.addLayout(form_layout)
        
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # Connections
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        
        # Populate from initial config if it exists
        self._populate_from_initial_config(initial_config)

    # --- NEW METHOD ---
    def _browse_for_attachments(self):
        """Opens a file dialog to select multiple attachment files."""
        files, _ = QFileDialog.getOpenFileNames(self, "Select Attachments")
        if files:
            # Join the file paths with a semicolon and update the line edit
            self.attach_edit.setText(";".join(files))
    # --- END NEW METHOD ---

    def _populate_from_initial_config(self, config: Optional[Dict[str, Any]]):
        if not config:
            return

        # Helper to populate a single field
        def populate_field(cfg_key, line_edit, combo):
            value = config.get(f"{cfg_key}_var")
            if value and value in self.global_variables:
                combo.setCurrentText(value)
            else:
                # Handle QPlainTextEdit vs QLineEdit
                if isinstance(line_edit, QPlainTextEdit):
                    line_edit.setPlainText(config.get(cfg_key, ""))
                else:
                    line_edit.setText(config.get(cfg_key, ""))
                combo.setCurrentText("-- Select Variable --")

        populate_field("to", self.to_edit, self.to_combo)
        populate_field("cc", self.cc_edit, self.cc_combo)
        populate_field("bcc", self.bcc_edit, self.bcc_combo)
        populate_field("subject", self.subject_edit, self.subject_combo)
        populate_field("body", self.body_edit, self.body_combo) # Now works correctly with helper
        populate_field("attachments", self.attach_edit, self.attach_combo)


    def get_executor_method_name(self) -> str:
        return "_send_outlook_email"

    def get_assignment_variable(self) -> Optional[str]:
        # This action does not produce a variable
        return None

    def get_config_data(self) -> Optional[Dict[str, Any]]:
        config = {}
        
        # Helper to extract data from a field
        def get_field_data(cfg_key, line_edit, combo):
            if combo.currentText() != "-- Select Variable --":
                config[f"{cfg_key}_var"] = combo.currentText()
                config[cfg_key] = "" # Ensure the hardcoded value is empty
            else:
                # Handle QPlainTextEdit vs QLineEdit
                text = line_edit.toPlainText().strip() if isinstance(line_edit, QPlainTextEdit) else line_edit.text().strip()
                config[cfg_key] = text
                config[f"{cfg_key}_var"] = "" # Ensure the var is empty
        
        get_field_data("to", self.to_edit, self.to_combo)
        get_field_data("cc", self.cc_edit, self.cc_combo)
        get_field_data("bcc", self.bcc_edit, self.bcc_combo)
        get_field_data("subject", self.subject_edit, self.subject_combo)
        get_field_data("body", self.body_edit, self.body_combo)
        get_field_data("attachments", self.attach_edit, self.attach_combo)

        # Validation: 'To' field must not be empty
        if not config["to"] and not config["to_var"]:
            QMessageBox.warning(self, "Input Error", "The 'To' field cannot be empty. Please provide an email address or select a variable.")
            return None
            
        return config

#
# --- The Public-Facing Module Class for Sending Outlook Mail ---
#
class Send_Outlook_Mail:
    def __init__(self, context: Optional[ExecutionContext] = None):
        self.context = context

    def _log(self, m: str):
        if self.context:
            self.context.add_log(m)
        else:
            print(m)

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        """Opens the configuration dialog for sending an email."""
        self._log("Opening Send Outlook Email configuration...")
        return _SendOutlookDialog(global_variables, parent_window, **kwargs)

    def _send_outlook_email(self, context: ExecutionContext, config_data: dict):
        """Creates and sends an email using Outlook."""
        self.context = context
        if not win32com:
            raise ImportError("'pywin32' library is not installed.")

        # Helper to resolve value from hardcode or variable
        def resolve_value(key: str) -> str:
            var_name = config_data.get(f"{key}_var")
            if var_name:
                value = self.context.get_variable(var_name)
                if value is None:
                    self._log(f"Warning: Global variable '{var_name}' for '{key}' not found or is None. Using empty string.")
                    return ""
                return str(value)
            return config_data.get(key, "")

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # 0: olMailItem

            mail.To = resolve_value("to")
            mail.CC = resolve_value("cc")
            mail.BCC = resolve_value("bcc")
            mail.Subject = resolve_value("subject")
            mail.Body = resolve_value("body")
            
            self._log(f"Preparing email to: {mail.To} with subject: '{mail.Subject}'")

            # Handle attachments
            attachments_str = resolve_value("attachments")
            if attachments_str:
                # Split by semicolon and strip whitespace from each path
                attachment_paths = [path.strip() for path in attachments_str.split(';') if path.strip()]
                for path in attachment_paths:
                    if os.path.exists(path):
                        mail.Attachments.Add(path)
                        self._log(f"Attaching file: {path}")
                    else:
                        self._log(f"Warning: Attachment path not found, skipping: {path}")

            mail.Send()
            self._log("Email sent successfully.")

        except Exception as e:
            error_message = f"Failed to send Outlook email. Error: {e}"
            self._log(error_message)
            # Re-raise the exception to halt execution and show the error in the main app
            raise RuntimeError(error_message) from e
