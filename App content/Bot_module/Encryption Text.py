# File: Bot_module/crypto_module.py

import sys
from typing import Optional, List, Dict, Any
import base64
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
import os

# --- PyQt6 Imports ---
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QDialogButtonBox,
    QComboBox, QWidget, QGroupBox, QMessageBox, QLabel,
    QCheckBox, QApplication, QFileDialog, QTableView, QListWidget, QListWidgetItem,
    QHBoxLayout, QRadioButton, QSpinBox, QProgressBar, QTextEdit, QPlainTextEdit
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
            return None # Fallback behavior

#
# --- PRIVATE: Text Cryptography Helper Class ---
#
class _TextCrypto:
    def __init__(self, method='fernet'):
        """
        Initialize the _TextCrypto class
        
        Args:
            method (str): Encryption method ('fernet', 'caesar', 'base64')
        """
        self.method = method.lower()
        self.key = None
        self.salt = None
        self.cipher_suite = None
        
        if self.method == 'fernet':
            # Don't auto-generate key, wait for password or explicit key
            pass
    
    def set_password(self, password):
        """Set a password for key derivation (used with Fernet method)"""
        if self.method == 'fernet':
            password_bytes = password.encode('utf-8')
            # Use a fixed salt for password-based encryption to ensure consistency
            # In production, you might want to store the salt with the encrypted data
            self.salt = b'salt_1234567890'  # Fixed 16-byte salt for consistency
            kdf = PBKDF2HMAC(
                algorithm=hashes.SHA256(),
                length=32,
                salt=self.salt,
                iterations=100000,
            )
            key = base64.urlsafe_b64encode(kdf.derive(password_bytes))
            self.key = key
            self.cipher_suite = Fernet(key)
    
    def set_key(self, key):
        """Set encryption key directly"""
        if self.method == 'fernet':
            self.key = key
            self.cipher_suite = Fernet(key)
    
    def generate_key(self):
        """Generate a new random key for Fernet"""
        if self.method == 'fernet':
            self.key = Fernet.generate_key()
            self.cipher_suite = Fernet(self.key)
            return self.key
        return None
    
    def encrypt(self, text, shift=3):
        """Encrypt text using the specified method"""
        if not isinstance(text, str):
            text = str(text)
            
        if self.method == 'fernet':
            return self._fernet_encrypt(text)
        elif self.method == 'caesar':
            return self._caesar_encrypt(text, shift)
        elif self.method == 'base64':
            return self._base64_encrypt(text)
        else:
            raise ValueError(f"Unsupported encryption method: {self.method}")
    
    def decrypt(self, encrypted_text, shift=3):
        """Decrypt text using the specified method"""
        if not isinstance(encrypted_text, str):
            encrypted_text = str(encrypted_text)
            
        if self.method == 'fernet':
            return self._fernet_decrypt(encrypted_text)
        elif self.method == 'caesar':
            return self._caesar_decrypt(encrypted_text, shift)
        elif self.method == 'base64':
            return self._base64_decrypt(encrypted_text)
        else:
            raise ValueError(f"Unsupported decryption method: {self.method}")
    
    def _fernet_encrypt(self, text):
        """Encrypt using Fernet (AES 128 in CBC mode)"""
        if not self.cipher_suite:
            raise ValueError("No key or password set for Fernet encryption. Call set_password() or set_key() first.")
        
        try:
            text_bytes = text.encode('utf-8')
            encrypted_bytes = self.cipher_suite.encrypt(text_bytes)
            # Return base64 encoded string for easy storage/transmission
            return base64.urlsafe_b64encode(encrypted_bytes).decode('utf-8')
        except Exception as e:
            raise ValueError(f"Fernet encryption failed: {str(e)}")
    
    def _fernet_decrypt(self, encrypted_text):
        """Decrypt using Fernet with improved error handling"""
        if not self.cipher_suite:
            raise ValueError("No key or password set for Fernet decryption. Call set_password() or set_key() first.")
        
        try:
            # Handle both direct Fernet tokens and base64-encoded tokens
            try:
                # First try: assume it's our base64-encoded format
                encrypted_bytes = base64.urlsafe_b64decode(encrypted_text.encode('utf-8'))
            except:
                # Second try: assume it's already a Fernet token
                encrypted_bytes = encrypted_text.encode('utf-8')
            
            decrypted_bytes = self.cipher_suite.decrypt(encrypted_bytes)
            return decrypted_bytes.decode('utf-8')
            
        except Exception as e:
            error_msg = str(e).lower()
            
            if "invalid" in error_msg or "token" in error_msg:
                raise ValueError("Invalid encrypted text format. The text may be corrupted or not encrypted with Fernet.")
            elif "decrypt" in error_msg:
                raise ValueError("Decryption failed. This could be due to wrong password, corrupted data, or different encryption key.")
            elif "padding" in error_msg:
                raise ValueError("Invalid padding in encrypted text. The data may be corrupted.")
            else:
                raise ValueError(f"Fernet decryption failed: {str(e)}")
    
    def _caesar_encrypt(self, text, shift):
        """Encrypt using Caesar cipher"""
        encrypted = ""
        for char in text:
            if char.isalpha():
                ascii_offset = ord('A') if char.isupper() else ord('a')
                encrypted += chr((ord(char) - ascii_offset + shift) % 26 + ascii_offset)
            else:
                encrypted += char
        return encrypted
    
    def _caesar_decrypt(self, encrypted_text, shift):
        """Decrypt using Caesar cipher"""
        return self._caesar_encrypt(encrypted_text, -shift)
    
    def _base64_encrypt(self, text):
        """Encrypt using Base64 encoding (not secure, just obfuscation)"""
        try:
            text_bytes = text.encode('utf-8')
            encrypted_bytes = base64.b64encode(text_bytes)
            return encrypted_bytes.decode('utf-8')
        except Exception as e:
            raise ValueError(f"Base64 encoding failed: {str(e)}")
    
    def _base64_decrypt(self, encrypted_text):
        """Decrypt using Base64 decoding with improved error handling"""
        try:
            # Remove any whitespace that might cause issues
            encrypted_text = encrypted_text.strip()
            encrypted_bytes = encrypted_text.encode('utf-8')
            decrypted_bytes = base64.b64decode(encrypted_bytes)
            return decrypted_bytes.decode('utf-8')
        except Exception as e:
            error_msg = str(e).lower()
            
            if "invalid" in error_msg or "padding" in error_msg:
                raise ValueError("Invalid Base64 format. The text may be corrupted or not properly Base64 encoded.")
            else:
                raise ValueError(f"Base64 decryption failed: {str(e)}")

#
# --- HELPER: The GUI Dialog for Text Encryption ---
#
class _TextEncryptDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None,
                 initial_variable: Optional[str] = None):
        super().__init__(parent)
        self.setWindowTitle("Text Encryption")
        self.setMinimumSize(700, 650)
        self.global_variables = global_variables
        
        main_layout = QVBoxLayout(self)

        # 1. Text Source
        source_group = QGroupBox("Text to Encrypt")
        source_layout = QFormLayout(source_group)
        
        self.text_source_combo = QComboBox()
        self.text_source_combo.addItems(['Direct Input', 'From Variable'])
        
        # Direct input
        self.direct_text_edit = QPlainTextEdit()
        self.direct_text_edit.setPlaceholderText("Enter text to encrypt here...")
        self.direct_text_edit.setMaximumHeight(100)
        
        # Variable input
        self.variable_combo = QComboBox()
        self.variable_combo.addItems(["-- Select Variable --"] + global_variables)
        
        source_layout.addRow("Text Source:", self.text_source_combo)
        source_layout.addRow("Direct Input:", self.direct_text_edit)
        source_layout.addRow("From Variable:", self.variable_combo)
        main_layout.addWidget(source_group)

        # 2. Encryption Method
        method_group = QGroupBox("Encryption Method")
        method_layout = QFormLayout(method_group)
        
        self.encryption_method_combo = QComboBox()
        self.encryption_method_combo.addItems(['fernet', 'caesar', 'base64'])
        
        # Password/Key settings
        self.password_source_combo = QComboBox()
        self.password_source_combo.addItems(['Hardcoded', 'From Global Variable'])
        
        self.hardcoded_password_edit = QLineEdit()
        self.hardcoded_password_edit.setPlaceholderText("Enter password for encryption")
        self.hardcoded_password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        
        self.password_variable_combo = QComboBox()
        self.password_variable_combo.addItems(["-- Select Variable --"] + global_variables)
        
        # Caesar cipher shift
        self.caesar_shift_spin = QSpinBox()
        self.caesar_shift_spin.setRange(1, 25)
        self.caesar_shift_spin.setValue(3)
        self.caesar_shift_label = QLabel("Caesar Shift:")
        
        method_layout.addRow("Encryption Method:", self.encryption_method_combo)
        method_layout.addRow("Password Source:", self.password_source_combo)
        method_layout.addRow("Hardcoded Password:", self.hardcoded_password_edit)
        method_layout.addRow("Password Variable:", self.password_variable_combo)
        method_layout.addRow(self.caesar_shift_label, self.caesar_shift_spin)
        main_layout.addWidget(method_group)

        # 3. Output Assignment
        output_group = QGroupBox("Assign Encrypted Result")
        output_layout = QFormLayout(output_group)
        
        self.new_var_radio = QRadioButton("New Variable Name:")
        self.new_var_input = QLineEdit("encrypted_text")
        
        self.existing_var_radio = QRadioButton("Existing Variable:")
        self.existing_var_combo = QComboBox()
        self.existing_var_combo.addItems(["-- Select --"] + global_variables)
        
        output_layout.addRow(self.new_var_radio, self.new_var_input)
        output_layout.addRow(self.existing_var_radio, self.existing_var_combo)
        main_layout.addWidget(output_group)
        
        # 4. Preview
        preview_group = QGroupBox("Preview")
        preview_layout = QVBoxLayout(preview_group)
        self.preview_button = QPushButton("Preview Encryption")
        self.preview_text = QPlainTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setMaximumHeight(80)
        preview_layout.addWidget(self.preview_button)
        preview_layout.addWidget(self.preview_text)
        main_layout.addWidget(preview_group)

        # Dialog Buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # Connections
        self.text_source_combo.currentTextChanged.connect(self._on_source_changed)
        self.encryption_method_combo.currentTextChanged.connect(self._on_method_changed)
        self.password_source_combo.currentTextChanged.connect(self._on_password_source_changed)
        self.preview_button.clicked.connect(self._preview_encryption)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        # Initial setup
        self.new_var_radio.setChecked(True)
        self._on_source_changed()
        self._on_method_changed()
        self._on_password_source_changed()
        
        if initial_config:
            self._populate_from_initial_config(initial_config, initial_variable)

    def _on_source_changed(self):
        is_direct = self.text_source_combo.currentText() == 'Direct Input'
        self.direct_text_edit.setVisible(is_direct)
        self.variable_combo.setVisible(not is_direct)

    def _on_method_changed(self):
        method = self.encryption_method_combo.currentText()
        is_caesar = method == 'caesar'
        is_fernet = method == 'fernet'
        
        self.caesar_shift_label.setVisible(is_caesar)
        self.caesar_shift_spin.setVisible(is_caesar)
        
        self.password_source_combo.setVisible(is_fernet)
        self.hardcoded_password_edit.setVisible(is_fernet)
        self.password_variable_combo.setVisible(is_fernet)

    def _on_password_source_changed(self):
        if self.encryption_method_combo.currentText() != 'fernet':
            return
            
        is_hardcoded = self.password_source_combo.currentText() == 'Hardcoded'
        self.hardcoded_password_edit.setVisible(is_hardcoded)
        self.password_variable_combo.setVisible(not is_hardcoded)

    def _preview_encryption(self):
        try:
            config = self._get_preview_config()
            if not config:
                return
                
            # Create crypto instance
            crypto = _TextCrypto(config['method'])
            
            # Set password if needed
            if config['method'] == 'fernet' and config.get('password'):
                crypto.set_password(config['password'])
            
            # Encrypt text
            encrypted = crypto.encrypt(config['text'], config.get('shift', 3))
            self.preview_text.setPlainText(f"Encrypted: {encrypted[:200]}{'...' if len(encrypted) > 200 else ''}")
            
        except Exception as e:
            self.preview_text.setPlainText(f"Preview Error: {str(e)}")

    def _get_preview_config(self):
        # Get text
        if self.text_source_combo.currentText() == 'Direct Input':
            text = self.direct_text_edit.toPlainText()
            if not text:
                QMessageBox.warning(self, "Input Error", "Please enter text to encrypt.")
                return None
        else:
            QMessageBox.information(self, "Preview", "Preview with variable input will be available during execution.")
            return None
        
        # Get method and settings
        method = self.encryption_method_combo.currentText()
        config = {'text': text, 'method': method}
        
        if method == 'fernet':
            if self.password_source_combo.currentText() == 'Hardcoded':
                password = self.hardcoded_password_edit.text()
                if not password:
                    QMessageBox.warning(self, "Input Error", "Please enter a password for Fernet encryption.")
                    return None
                config['password'] = password
            else:
                QMessageBox.information(self, "Preview", "Preview with variable password will be available during execution.")
                return None
        elif method == 'caesar':
            config['shift'] = self.caesar_shift_spin.value()
        
        return config

    def _populate_from_initial_config(self, config, variable):
        self.text_source_combo.setCurrentText(config.get("text_source", "Direct Input"))
        self.direct_text_edit.setPlainText(config.get("direct_text", ""))
        
        if config.get("source_variable") in self.global_variables:
            self.variable_combo.setCurrentText(config.get("source_variable"))
        
        self.encryption_method_combo.setCurrentText(config.get("method", "fernet"))
        self.password_source_combo.setCurrentText(config.get("password_source", "Hardcoded"))
        self.hardcoded_password_edit.setText(config.get("hardcoded_password", ""))
        
        if config.get("password_variable") in self.global_variables:
            self.password_variable_combo.setCurrentText(config.get("password_variable"))
        
        self.caesar_shift_spin.setValue(config.get("caesar_shift", 3))
        
        if variable:
            if variable in self.global_variables:
                self.existing_var_radio.setChecked(True)
                self.existing_var_combo.setCurrentText(variable)
            else:
                self.new_var_radio.setChecked(True)
                self.new_var_input.setText(variable)

    def get_executor_method_name(self) -> str:
        return "_encrypt_text"

    def get_assignment_variable(self) -> Optional[str]:
        if self.new_var_radio.isChecked():
            var_name = self.new_var_input.text().strip()
            if not var_name:
                QMessageBox.warning(self, "Input Error", "New variable name cannot be empty.")
                return None
            return var_name
        else:
            var_name = self.existing_var_combo.currentText()
            if var_name == "-- Select --":
                QMessageBox.warning(self, "Input Error", "Please select an existing variable.")
                return None
            return var_name

    def get_config_data(self) -> Optional[Dict[str, Any]]:
        config = {}
        
        # Text source
        config['text_source'] = self.text_source_combo.currentText()
        if config['text_source'] == 'Direct Input':
            text = self.direct_text_edit.toPlainText()
            if not text:
                QMessageBox.warning(self, "Input Error", "Please enter text to encrypt.")
                return None
            config['direct_text'] = text
        else:
            source_var = self.variable_combo.currentText()
            if source_var == "-- Select Variable --":
                QMessageBox.warning(self, "Input Error", "Please select a source variable.")
                return None
            config['source_variable'] = source_var
        
        # Encryption method
        config['method'] = self.encryption_method_combo.currentText()
        
        if config['method'] == 'fernet':
            config['password_source'] = self.password_source_combo.currentText()
            if config['password_source'] == 'Hardcoded':
                password = self.hardcoded_password_edit.text()
                if not password:
                    QMessageBox.warning(self, "Input Error", "Please enter a password for Fernet encryption.")
                    return None
                config['hardcoded_password'] = password
            else:
                password_var = self.password_variable_combo.currentText()
                if password_var == "-- Select Variable --":
                    QMessageBox.warning(self, "Input Error", "Please select a password variable.")
                    return None
                config['password_variable'] = password_var
        elif config['method'] == 'caesar':
            config['caesar_shift'] = self.caesar_shift_spin.value()
        
        return config

#
# --- HELPER: The GUI Dialog for Text Decryption ---
#
class _TextDecryptDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None,
                 initial_variable: Optional[str] = None):
        super().__init__(parent)
        self.setWindowTitle("Text Decryption")
        self.setMinimumSize(700, 650)
        self.global_variables = global_variables
        
        main_layout = QVBoxLayout(self)

        # 1. Encrypted Text Source
        source_group = QGroupBox("Encrypted Text to Decrypt")
        source_layout = QFormLayout(source_group)
        
        self.text_source_combo = QComboBox()
        self.text_source_combo.addItems(['Direct Input', 'From Variable'])
        
        # Direct input
        self.direct_text_edit = QPlainTextEdit()
        self.direct_text_edit.setPlaceholderText("Enter encrypted text to decrypt here...")
        self.direct_text_edit.setMaximumHeight(100)
        
        # Variable input
        self.variable_combo = QComboBox()
        self.variable_combo.addItems(["-- Select Variable --"] + global_variables)
        
        source_layout.addRow("Encrypted Text Source:", self.text_source_combo)
        source_layout.addRow("Direct Input:", self.direct_text_edit)
        source_layout.addRow("From Variable:", self.variable_combo)
        main_layout.addWidget(source_group)

        # 2. Decryption Method
        method_group = QGroupBox("Decryption Method")
        method_layout = QFormLayout(method_group)
        
        self.decryption_method_combo = QComboBox()
        self.decryption_method_combo.addItems(['fernet', 'caesar', 'base64'])
        
        # Password/Key settings
        self.password_source_combo = QComboBox()
        self.password_source_combo.addItems(['Hardcoded', 'From Global Variable'])
        
        self.hardcoded_password_edit = QLineEdit()
        self.hardcoded_password_edit.setPlaceholderText("Enter password for decryption")
        self.hardcoded_password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        
        self.password_variable_combo = QComboBox()
        self.password_variable_combo.addItems(["-- Select Variable --"] + global_variables)
        
        # Caesar cipher shift
        self.caesar_shift_spin = QSpinBox()
        self.caesar_shift_spin.setRange(1, 25)
        self.caesar_shift_spin.setValue(3)
        self.caesar_shift_label = QLabel("Caesar Shift:")
        
        method_layout.addRow("Decryption Method:", self.decryption_method_combo)
        method_layout.addRow("Password Source:", self.password_source_combo)
        method_layout.addRow("Hardcoded Password:", self.hardcoded_password_edit)
        method_layout.addRow("Password Variable:", self.password_variable_combo)
        method_layout.addRow(self.caesar_shift_label, self.caesar_shift_spin)
        main_layout.addWidget(method_group)

        # 3. Output Assignment
        output_group = QGroupBox("Assign Decrypted Result")
        output_layout = QFormLayout(output_group)
        
        self.new_var_radio = QRadioButton("New Variable Name:")
        self.new_var_input = QLineEdit("decrypted_text")
        
        self.existing_var_radio = QRadioButton("Existing Variable:")
        self.existing_var_combo = QComboBox()
        self.existing_var_combo.addItems(["-- Select --"] + global_variables)
        
        output_layout.addRow(self.new_var_radio, self.new_var_input)
        output_layout.addRow(self.existing_var_radio, self.existing_var_combo)
        main_layout.addWidget(output_group)

        # 4. Preview/Test
        preview_group = QGroupBox("Test Decryption")
        preview_layout = QVBoxLayout(preview_group)
        self.test_button = QPushButton("Test Decryption")
        self.test_result_text = QPlainTextEdit()
        self.test_result_text.setReadOnly(True)
        self.test_result_text.setMaximumHeight(80)
        preview_layout.addWidget(self.test_button)
        preview_layout.addWidget(self.test_result_text)
        main_layout.addWidget(preview_group)

        # Dialog Buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # Connections
        self.text_source_combo.currentTextChanged.connect(self._on_source_changed)
        self.decryption_method_combo.currentTextChanged.connect(self._on_method_changed)
        self.password_source_combo.currentTextChanged.connect(self._on_password_source_changed)
        self.test_button.clicked.connect(self._test_decryption)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        # Initial setup
        self.new_var_radio.setChecked(True)
        self._on_source_changed()
        self._on_method_changed()
        self._on_password_source_changed()
        
        if initial_config:
            self._populate_from_initial_config(initial_config, initial_variable)

    def _on_source_changed(self):
        is_direct = self.text_source_combo.currentText() == 'Direct Input'
        self.direct_text_edit.setVisible(is_direct)
        self.variable_combo.setVisible(not is_direct)

    def _on_method_changed(self):
        method = self.decryption_method_combo.currentText()
        is_caesar = method == 'caesar'
        is_fernet = method == 'fernet'
        
        self.caesar_shift_label.setVisible(is_caesar)
        self.caesar_shift_spin.setVisible(is_caesar)
        
        self.password_source_combo.setVisible(is_fernet)
        self.hardcoded_password_edit.setVisible(is_fernet)
        self.password_variable_combo.setVisible(is_fernet)

    def _on_password_source_changed(self):
        if self.decryption_method_combo.currentText() != 'fernet':
            return
            
        is_hardcoded = self.password_source_combo.currentText() == 'Hardcoded'
        self.hardcoded_password_edit.setVisible(is_hardcoded)
        self.password_variable_combo.setVisible(not is_hardcoded)

    def _test_decryption(self):
        try:
            config = self._get_test_config()
            if not config:
                return
                
            # Create crypto instance
            crypto = _TextCrypto(config['method'])
            
            # Set password if needed
            if config['method'] == 'fernet' and config.get('password'):
                crypto.set_password(config['password'])
            
            # Test decrypt
            decrypted = crypto.decrypt(config['encrypted_text'], config.get('shift', 3))
            self.test_result_text.setPlainText(f"Decrypted successfully: {decrypted[:200]}{'...' if len(decrypted) > 200 else ''}")
            
        except Exception as e:
            self.test_result_text.setPlainText(f"Decryption Error: {str(e)}")

    def _get_test_config(self):
        # Get encrypted text
        if self.text_source_combo.currentText() == 'Direct Input':
            encrypted_text = self.direct_text_edit.toPlainText()
            if not encrypted_text:
                QMessageBox.warning(self, "Input Error", "Please enter encrypted text to test.")
                return None
        else:
            QMessageBox.information(self, "Test", "Testing with variable input will be available during execution.")
            return None
        
        # Get method and settings
        method = self.decryption_method_combo.currentText()
        config = {'encrypted_text': encrypted_text, 'method': method}
        
        if method == 'fernet':
            if self.password_source_combo.currentText() == 'Hardcoded':
                password = self.hardcoded_password_edit.text()
                if not password:
                    QMessageBox.warning(self, "Input Error", "Please enter a password for Fernet decryption.")
                    return None
                config['password'] = password
            else:
                QMessageBox.information(self, "Test", "Testing with variable password will be available during execution.")
                return None
        elif method == 'caesar':
            config['shift'] = self.caesar_shift_spin.value()
        
        return config

    def _populate_from_initial_config(self, config, variable):
        self.text_source_combo.setCurrentText(config.get("text_source", "Direct Input"))
        self.direct_text_edit.setPlainText(config.get("direct_text", ""))
        
        if config.get("source_variable") in self.global_variables:
            self.variable_combo.setCurrentText(config.get("source_variable"))
        
        self.decryption_method_combo.setCurrentText(config.get("method", "fernet"))
        self.password_source_combo.setCurrentText(config.get("password_source", "Hardcoded"))
        self.hardcoded_password_edit.setText(config.get("hardcoded_password", ""))
        
        if config.get("password_variable") in self.global_variables:
            self.password_variable_combo.setCurrentText(config.get("password_variable"))
        
        self.caesar_shift_spin.setValue(config.get("caesar_shift", 3))
        
        if variable:
            if variable in self.global_variables:
                self.existing_var_radio.setChecked(True)
                self.existing_var_combo.setCurrentText(variable)
            else:
                self.new_var_radio.setChecked(True)
                self.new_var_input.setText(variable)

    def get_executor_method_name(self) -> str:
        return "_decrypt_text"

    def get_assignment_variable(self) -> Optional[str]:
        if self.new_var_radio.isChecked():
            var_name = self.new_var_input.text().strip()
            if not var_name:
                QMessageBox.warning(self, "Input Error", "New variable name cannot be empty.")
                return None
            return var_name
        else:
            var_name = self.existing_var_combo.currentText()
            if var_name == "-- Select --":
                QMessageBox.warning(self, "Input Error", "Please select an existing variable.")
                return None
            return var_name

    def get_config_data(self) -> Optional[Dict[str, Any]]:
        config = {}
        
        # Text source
        config['text_source'] = self.text_source_combo.currentText()
        if config['text_source'] == 'Direct Input':
            text = self.direct_text_edit.toPlainText()
            if not text:
                QMessageBox.warning(self, "Input Error", "Please enter encrypted text to decrypt.")
                return None
            config['direct_text'] = text
        else:
            source_var = self.variable_combo.currentText()
            if source_var == "-- Select Variable --":
                QMessageBox.warning(self, "Input Error", "Please select a source variable.")
                return None
            config['source_variable'] = source_var
        
        # Decryption method
        config['method'] = self.decryption_method_combo.currentText()
        
        if config['method'] == 'fernet':
            config['password_source'] = self.password_source_combo.currentText()
            if config['password_source'] == 'Hardcoded':
                password = self.hardcoded_password_edit.text()
                if not password:
                    QMessageBox.warning(self, "Input Error", "Please enter a password for Fernet decryption.")
                    return None
                config['hardcoded_password'] = password
            else:
                password_var = self.password_variable_combo.currentText()
                if password_var == "-- Select Variable --":
                    QMessageBox.warning(self, "Input Error", "Please select a password variable.")
                    return None
                config['password_variable'] = password_var
        elif config['method'] == 'caesar':
            config['caesar_shift'] = self.caesar_shift_spin.value()
        
        return config

#
# --- The Public-Facing Module Class for Text Encryption ---
#
class encrypt_text:
    def __init__(self, context: Optional[ExecutionContext] = None):
        self.context = context

    def _log(self, m: str):
        if self.context:
            self.context.add_log(m)
        else:
            print(m)

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        self._log("Opening Text Encryption configuration...")
        return _TextEncryptDialog(global_variables, parent_window, **kwargs)

    def _encrypt_text(self, context: ExecutionContext, config_data: dict) -> str:
        self.context = context
        
        try:
            # Get text to encrypt
            if config_data['text_source'] == 'Direct Input':
                text_to_encrypt = config_data['direct_text']
            else:
                source_var = config_data['source_variable']
                text_to_encrypt = self.context.get_variable(source_var)
                if text_to_encrypt is None:
                    raise ValueError(f"Variable '{source_var}' not found or is None.")
                text_to_encrypt = str(text_to_encrypt)
            
            self._log(f"Encrypting text using {config_data['method']} method...")
            
            # Create crypto instance
            crypto = _TextCrypto(config_data['method'])
            
            # Set password/key if needed
            if config_data['method'] == 'fernet':
                if config_data['password_source'] == 'Hardcoded':
                    password = config_data['hardcoded_password']
                else:
                    password_var = config_data['password_variable']
                    password = self.context.get_variable(password_var)
                    if password is None:
                        raise ValueError(f"Password variable '{password_var}' not found or is None.")
                    password = str(password)
                
                crypto.set_password(password)
                encrypted_text = crypto.encrypt(text_to_encrypt)
            elif config_data['method'] == 'caesar':
                shift = config_data['caesar_shift']
                encrypted_text = crypto.encrypt(text_to_encrypt, shift)
            else:  # base64
                encrypted_text = crypto.encrypt(text_to_encrypt)
            
            self._log(f"Successfully encrypted text. Result length: {len(encrypted_text)} characters.")
            return encrypted_text
            
        except Exception as e:
            self._log(f"Encryption failed: {str(e)}")
            raise

#
# --- The Public-Facing Module Class for Text Decryption ---
#
class decrypt_text:
    def __init__(self, context: Optional[ExecutionContext] = None):
        self.context = context

    def _log(self, m: str):
        if self.context:
            self.context.add_log(m)
        else:
            print(m)

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        self._log("Opening Text Decryption configuration...")
        return _TextDecryptDialog(global_variables, parent_window, **kwargs)

    def _decrypt_text(self, context: ExecutionContext, config_data: dict) -> str:
        self.context = context
        
        try:
            # Get encrypted text to decrypt
            if config_data['text_source'] == 'Direct Input':
                encrypted_text = config_data['direct_text']
            else:
                source_var = config_data['source_variable']
                encrypted_text = self.context.get_variable(source_var)
                if encrypted_text is None:
                    raise ValueError(f"Variable '{source_var}' not found or is None.")
                encrypted_text = str(encrypted_text).strip()
            
            if not encrypted_text:
                raise ValueError("Encrypted text is empty.")
            
            self._log(f"Decrypting text using {config_data['method']} method...")
            self._log(f"Encrypted text length: {len(encrypted_text)} characters")
            
            # Create crypto instance
            crypto = _TextCrypto(config_data['method'])
            
            # Set password/key if needed
            if config_data['method'] == 'fernet':
                if config_data['password_source'] == 'Hardcoded':
                    password = config_data['hardcoded_password']
                else:
                    password_var = config_data['password_variable']
                    password = self.context.get_variable(password_var)
                    if password is None:
                        raise ValueError(f"Password variable '{password_var}' not found or is None.")
                    password = str(password)
                
                self._log("Setting password for Fernet decryption...")
                crypto.set_password(password)
                decrypted_text = crypto.decrypt(encrypted_text)
            elif config_data['method'] == 'caesar':
                shift = config_data['caesar_shift']
                decrypted_text = crypto.decrypt(encrypted_text, shift)
            else:  # base64
                decrypted_text = crypto.decrypt(encrypted_text)
            
            self._log(f"Successfully decrypted text. Result length: {len(decrypted_text)} characters.")
            return decrypted_text
            
        except ValueError as e:
            # Re-raise ValueError with context
            error_msg = f"Decryption failed: {str(e)}"
            self._log(error_msg)
            raise ValueError(error_msg)
        except Exception as e:
            # Handle other exceptions
            error_msg = f"Unexpected error during decryption: {str(e)}"
            self._log(error_msg)
            raise Exception(error_msg)

#
# --- Example Usage and Testing ---
#
if __name__ == "__main__":
    import sys
    from PyQt6.QtWidgets import QApplication
    
    app = QApplication(sys.argv)
    
    # Test the encryption/decryption flow
    def test_crypto_flow():
        # Test with hardcoded values
        crypto = _TextCrypto('fernet')
        crypto.set_password('test_password')
        
        original_text = "Hello, this is a secret message!"
        print(f"Original: {original_text}")
        
        # Encrypt
        encrypted = crypto.encrypt(original_text)
        print(f"Encrypted: {encrypted}")
        
        # Decrypt with same password
        crypto2 = _TextCrypto('fernet')
        crypto2.set_password('test_password')
        decrypted = crypto2.decrypt(encrypted)
        print(f"Decrypted: {decrypted}")
        
        # Test with wrong password
        try:
            crypto3 = _TextCrypto('fernet')
            crypto3.set_password('wrong_password')
            bad_decrypt = crypto3.decrypt(encrypted)
        except Exception as e:
            print(f"Expected error with wrong password: {e}")
    
    test_crypto_flow()
    
    # Test GUI dialogs
    test_variables = ["var1", "var2", "secret_password", "encrypted_data"]
    
    # Test encryption dialog
    encrypt_module = encrypt_text()
    encrypt_dialog = encrypt_module.configure_data_hub(None, test_variables)
    
    if encrypt_dialog.exec() == QDialog.DialogCode.Accepted:
        config = encrypt_dialog.get_config_data()
        variable = encrypt_dialog.get_assignment_variable()
        print("Encryption Config:", config)
        print("Assign to variable:", variable)
    
    sys.exit()