import sys
import os
import inspect
import importlib
import time
import csv
import io
from datetime import datetime
import json
import base64
import re # <--- ADD THIS LINE
import ast # <--- ADD THIS LINE
from PIL import ImageGrab
import PIL.Image
from PIL.ImageQt import ImageQt
from PyQt6.QtGui import QPixmap, QColor, QFont, QPainter, QPen
from PyQt6 import QtWidgets, QtGui
from typing import Optional, List, Dict, Any, Tuple, Union

# Ensure my_lib is in the Python path
script_dir = os.path.dirname(os.path.abspath(__file__))
my_lib_dir = os.path.join(script_dir, "my_lib")
if my_lib_dir not in sys.path:
    sys.path.insert(0, my_lib_dir)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QVariant, QObject, QSize, QPoint, QRegularExpression,QRect,QDateTime, QTimer
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QComboBox, QListWidget, QLabel, QPushButton, QListWidgetItem,
    QMessageBox, QProgressBar, QFileDialog, QDialog,
    QLineEdit, QVBoxLayout as QVBoxLayoutDialog, QFormLayout,
    QDialogButtonBox,
    QRadioButton, QGroupBox, QCheckBox, QTextEdit,
    QTreeWidget, QTreeWidgetItem, QGridLayout, QHeaderView, QSplitter, QInputDialog,
    QStackedLayout, QBoxLayout,QMenu,QPlainTextEdit,QSizePolicy, QTextBrowser,QDateTimeEdit
)

from PyQt6.QtGui import QIntValidator
from PyQt6.QtGui import QSyntaxHighlighter, QTextCharFormat, QColor, QFont, QPainter, QPen, QTextCursor,QTextFormat,QKeyEvent,QTextDocument

# Use the actual libraries from the my_lib folder
from my_lib.shared_context import ExecutionContext, GuiCommunicator
from my_lib.BOT_take_image import MainWindow as BotTakeImageWindow



class SecondWindow(QtWidgets.QDialog): # Or QtWidgets.QMainWindow if you prefer a full window

    screenshot_saved = pyqtSignal(str)

    def __init__(self, image: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Take and Manage Screenshots")
        self.setMinimumSize(679, 248) # Set a minimum size to match the original GUI

        self.bot_take_image_ui = BotTakeImageWindow(image)

        self.bot_take_image_ui.screenshotSaved.connect(self._handle_screenshot_saved)

        self.bot_take_image_ui.exit_BOT_butt.clicked.connect(self.accept) # Accept the dialog when exit is clicked

        layout = QVBoxLayout(self)
        layout.addWidget(self.bot_take_image_ui.centralwidget)
        self.setLayout(layout)

        self.finished.connect(self._on_dialog_closed)

    def _handle_screenshot_saved(self, filename: str):
        """Pass through the signal from the embedded UI."""
        self.screenshot_saved.emit(filename)

    def _on_dialog_closed(self, result: int):
        """Handles dialog closure, including via 'X' button."""

        if result == QDialog.DialogCode.Rejected: # Closed via 'X' or Escape
            self.screenshot_saved.emit("") # No specific file was saved/selected

# --- ExecutionStepCard ---
class ExecutionStepCard(QWidget):
    edit_requested = pyqtSignal(dict)
    delete_requested = pyqtSignal(dict)
    move_up_requested = pyqtSignal(dict)
    move_down_requested = pyqtSignal(dict)
    save_as_template_requested = pyqtSignal(dict)
    execute_this_requested = pyqtSignal(dict)


    def __init__(self, step_data: Dict[str, Any], step_number: int, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.step_data = step_data
        self.step_number = step_number
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(5)
        self.setObjectName("ExecutionStepCard")
        self.set_status("#D3D3D3")

        top_row_layout = QHBoxLayout()
        step_label_text = self._get_formatted_title()
        self.step_label = QLabel(step_label_text)
        
        step_type = self.step_data.get("type")
        if step_type == 'group_start':
            self.step_label.setStyleSheet("font-weight: bold; background-color: #D6EAF8; padding: 4px; border-radius: 3px;")
        else:
            self.step_label.setStyleSheet("font-weight: bold; background-color: #EAEAEA; padding: 4px; border-radius: 3px;")

        top_row_layout.addWidget(self.step_label)
        top_row_layout.addStretch()

        self.up_button = QPushButton("↑")
        self.down_button = QPushButton("↓")
        self.edit_button = QPushButton("Edit")
        self.delete_button = QPushButton("Delete")
        self.save_template_button = QPushButton("Save Template")
        self.execute_this_button = QPushButton("Execute This")

        button_font = self.up_button.font()
        button_font.setBold(True)
        self.up_button.setFont(button_font)
        self.down_button.setFont(button_font)

        self.up_button.setFixedSize(25, 25)
        self.down_button.setFixedSize(25, 25)
        self.edit_button.setFixedSize(60, 25)
        self.delete_button.setFixedSize(60, 25)
        self.save_template_button.setFixedSize(100, 25)
        self.execute_this_button.setFixedSize(100, 25)

        self.up_button.clicked.connect(lambda: self.move_up_requested.emit(self.step_data))
        self.down_button.clicked.connect(lambda: self.move_down_requested.emit(self.step_data))
        self.edit_button.clicked.connect(lambda: self.edit_requested.emit(self.step_data))
        self.delete_button.clicked.connect(lambda: self.delete_requested.emit(self.step_data))
        self.save_template_button.clicked.connect(lambda: self.save_as_template_requested.emit(self.step_data))
        self.execute_this_button.clicked.connect(lambda: self.execute_this_requested.emit(self.step_data))

        if self.step_data.get("type") not in ["loop_start", "IF_START", "group_start"]:
            self.save_template_button.hide()
        
        if self.step_data.get("type") in ["loop_end", "ELSE", "IF_END", "group_end"]:
            self.up_button.hide()
            self.down_button.hide()
            self.edit_button.hide()
            self.delete_button.hide()
            self.execute_this_button.hide()

        top_row_layout.addWidget(self.up_button)
        top_row_layout.addWidget(self.down_button)
        top_row_layout.addWidget(self.edit_button)
        top_row_layout.addWidget(self.delete_button)
        top_row_layout.addWidget(self.save_template_button)
        top_row_layout.addWidget(self.execute_this_button)

        main_layout.addLayout(top_row_layout)

        # NEW: Store original method text and create the label
        self._original_method_text = self._get_formatted_method_name()
        if self._original_method_text:
            self.method_label = QLabel(self._original_method_text)
            self.method_label.setStyleSheet("font-size: 10pt; padding: 5px; background-color: white; border: 1px solid #E0E0E0; border-radius: 3px;")
            self.method_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            main_layout.addWidget(self.method_label)
        else:
            self.method_label = None # Ensure it exists for other methods to check

        if self.step_data.get("type") == "step":
            parameters_config = self.step_data.get("parameters_config", {})
            if parameters_config:
                params_group = QGroupBox("Parameters")
                params_layout = QFormLayout()
                params_layout.setContentsMargins(8, 5, 8, 5)
                for param_name, config in parameters_config.items():
                    if param_name == "original_listbox_row_index": continue
                    value_str = ""
                    if config.get('type') == 'hardcoded': value_str = repr(config['value'])
                    elif config.get('type') == 'hardcoded_file': value_str = f"File: '{config['value']}'"
                    elif config.get('type') == 'variable': value_str = f"Variable: @{config['value']}"
                    param_label = QLabel(f"{param_name}:")
                    value_label = QLineEdit(value_str)
                    value_label.setReadOnly(True)
                    # NEW: Adjusted style for readability
                    value_label.setStyleSheet("background-color: #FFFFFF; font-size: 9pt; padding: 2px; border: 1px solid #D3D3D3;")
                    params_layout.addRow(param_label, value_label)
                params_group.setLayout(params_layout)
                main_layout.addWidget(params_group)

    def _get_formatted_title(self) -> str:
        step_type = self.step_data.get("type", "Unknown")
        if step_type == "group_start":
            return f"Group: {self.step_data.get('group_name', 'Unnamed')}"
        
        # MODIFIED: Only show method name for standard steps
        if step_type == "step":
            method_name = self.step_data.get("method_name", "UnknownMethod")
            return f"Step {self.step_number}: {method_name}"
            
        step_type_display = step_type.replace("_", " ").title()
        return f"Step {self.step_number}: {step_type_display}"

    def _get_formatted_method_name(self) -> str:
        step_type = self.step_data["type"]
        if step_type == "step":
            class_name, method_name = self.step_data["class_name"], self.step_data["method_name"]
            assign_to_variable_name = self.step_data["assign_to_variable_name"]
            display_text = f"{class_name}.{method_name}"
            if assign_to_variable_name:
                display_text = f"@{assign_to_variable_name} = " + display_text
            return display_text
        elif step_type == "loop_start":
            loop_config = self.step_data["loop_config"]
            custom_loop_name = loop_config.get("loop_name")
            name_display = f"'{custom_loop_name}'" if custom_loop_name else f"(ID: {self.step_data['loop_id']})"
            count_config = loop_config["iteration_count_config"]
            loop_info = f"@{count_config['value']}" if count_config["type"] == "variable" else f"{count_config['value']} times"
            assign_var = loop_config.get("assign_iteration_to_variable")
            if assign_var: loop_info += f", assign iter to @{assign_var}"
            return f"{name_display} - {loop_info}"
        elif step_type in ["loop_end", "ELSE", "IF_END", "group_start", "group_end"]:
            return ""
        elif step_type == "IF_START":
            condition_config = self.step_data["condition_config"]
            block_name = condition_config.get("block_name")
            name_display = f"'{block_name}'" if block_name else f"(ID: {self.step_data['if_id']})"
            left_op, right_op, op = condition_config["condition"]["left_operand"], condition_config["condition"]["right_operand"], condition_config["condition"]["operator"]
            left_str = f"@{left_op['value']}" if left_op['type'] == 'variable' else repr(left_op['value'])
            right_str = f"@{right_op['value']}" if right_op['type'] == 'variable' else repr(right_op['value'])
            return f"{name_display} ({left_str} {op} {right_str})"
        return ""

    def set_status(self, border_color: str, is_running: bool = False):
        background_color = "#D4EDDA" if is_running else "#F8F8F8"
        self.setStyleSheet(f"#ExecutionStepCard {{ background-color: {background_color}; border: 2px solid {border_color}; border-radius: 4px; }}")

    def set_result_text(self, result_message: str):
        """Displays the execution result on the card by updating the method label."""
        if not self.method_label:
            return

        # Truncate long results for display
        if len(result_message) > 300:
             result_message = result_message[:297] + "..."

        assign_to_var = self.step_data.get("assign_to_variable_name")
        
        # Check if this is a standard step with a variable assignment and a valid result message
        if self.step_data.get("type") == "step" and assign_to_var and "Result: " in result_message:
            try:
                # Extract the actual result value (it's between "Result: " and " (Assigned to")
                result_val_str = result_message.split("Result: ")[1].split(" (Assigned to")[0]
            except IndexError:
                result_val_str = "Error parsing result"
            
            # NEW FORMAT: @variable = result
            display_text = f"@{assign_to_var} = {result_val_str}"
            self.method_label.setText(display_text)
            self.method_label.setStyleSheet("font-size: 10pt; font-style: italic; color: #155724; padding: 5px; background-color: #d4edda; border: 1px solid #c3e6cb; border-radius: 3px;")
        
        else:
            # For other step types (loops, ifs) or steps without assignment, just show the worker message
            self.method_label.setText(f"✓ {result_message}")
            self.method_label.setStyleSheet("font-size: 10pt; font-style: italic; color: #155724; padding: 5px; background-color: #d4edda; border: 1px solid #c3e6cb; border-radius: 3px;")


    def clear_result(self):
        """Resets the method label to its pre-execution state."""
        if self.method_label:
            self.method_label.setText(self._original_method_text)
            # Reset to original style
            self.method_label.setStyleSheet("font-size: 10pt; padding: 5px; background-color: white; border: 1px solid #E0E0E0; border-radius: 3px;")

class LoopConfigDialog(QDialog):
    def __init__(self, global_variables: Dict[str, Any], parent: Optional[QWidget] = None, initial_config: Optional[Dict[str, Any]] = None):
        super().__init__(parent)
        self.setWindowTitle("Configure Loop Group")
        self.setMinimumWidth(300)
        self.global_variables = global_variables
        main_layout, form_layout = QVBoxLayout(), QFormLayout()
        self.loop_name_editor = QLineEdit()
        self.loop_name_editor.setPlaceholderText("Optional: Enter a name for this loop")
        form_layout.addRow("Loop Name:", self.loop_name_editor)
        self.repeat_count_editor = QLineEdit("1")
        self.repeat_count_editor.setValidator(QIntValidator(1, 999999, self))
        form_layout.addRow("Loop Count:", self.repeat_count_editor)
        self.use_var_checkbox = QCheckBox("Use Global Variable for Loop Count")
        self.use_var_checkbox.stateChanged.connect(self._toggle_count_var_input)
        form_layout.addRow("", self.use_var_checkbox)
        self.global_var_combo_count = QComboBox()
        self.global_var_combo_count.addItem("-- Select Global Variable --")
        self.global_var_combo_count.addItems(sorted(global_variables.keys()))
        self.global_var_combo_count.setEnabled(False)
        form_layout.addRow("Variable for Loop Count:", self.global_var_combo_count)
        self.assign_iter_checkbox = QCheckBox("Assign Current Iteration to Global Variable")
        self.assign_iter_checkbox.stateChanged.connect(self._toggle_assign_iter_input)
        form_layout.addRow("", self.assign_iter_checkbox)
        self.global_var_combo_assign_iter = QComboBox()
        self.global_var_combo_assign_iter.addItem("-- Select Global Variable --")
        self.global_var_combo_assign_iter.addItems(sorted(global_variables.keys()))
        self.global_var_combo_assign_iter.setEnabled(False)
        form_layout.addRow("Assign Iteration To:", self.global_var_combo_assign_iter)
        self.new_var_iter_editor = QLineEdit()
        self.new_var_iter_editor.setPlaceholderText("Enter new variable name")
        self.new_var_iter_editor.setEnabled(False)
        form_layout.addRow("New Var Name for Iter:", self.new_var_iter_editor)
        main_layout.addLayout(form_layout)
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        main_layout.addWidget(button_box)
        self.setLayout(main_layout)
        if initial_config: self.set_config(initial_config)
        else: self._toggle_count_var_input(); self._toggle_assign_iter_input()

    def _toggle_count_var_input(self) -> None:
        is_using_var = self.use_var_checkbox.isChecked()
        self.repeat_count_editor.setEnabled(not is_using_var)
        self.repeat_count_editor.setVisible(not is_using_var)
        self.global_var_combo_count.setEnabled(is_using_var)
        self.global_var_combo_count.setVisible(is_using_var)
        if is_using_var: self.repeat_count_editor.clear()
        else: self.global_var_combo_count.setCurrentIndex(0)

    def _toggle_assign_iter_input(self) -> None:
        is_assigning_iter = self.assign_iter_checkbox.isChecked()
        self.global_var_combo_assign_iter.setEnabled(is_assigning_iter)
        self.global_var_combo_assign_iter.setVisible(is_assigning_iter)
        self.new_var_iter_editor.setEnabled(is_assigning_iter and self.global_var_combo_assign_iter.currentIndex() == 0)
        self.new_var_iter_editor.setVisible(is_assigning_iter)
        self.global_var_combo_assign_iter.currentIndexChanged.connect(self._update_new_var_iter_editor_state)
        self._update_new_var_iter_editor_state()

    def _update_new_var_iter_editor_state(self) -> None:
        is_assigning_iter = self.assign_iter_checkbox.isChecked()
        is_selecting_new = self.global_var_combo_assign_iter.currentIndex() == 0
        self.new_var_iter_editor.setEnabled(is_assigning_iter and is_selecting_new)
        if not (is_assigning_iter and is_selecting_new): self.new_var_iter_editor.clear()

    def get_config(self) -> Optional[Dict[str, Any]]:
        loop_name = self.loop_name_editor.text().strip()
        count_config = {}
        if self.use_var_checkbox.isChecked():
            global_var_name = self.global_var_combo_count.currentText()
            if global_var_name == "-- Select Global Variable --": QMessageBox.warning(self, "Input Error", "Please select a global variable for loop count."); return None
            count_config = {"type": "variable", "value": global_var_name}
        else:
            try:
                repeat_count_str = self.repeat_count_editor.text()
                if not repeat_count_str: QMessageBox.warning(self, "Input Error", "Please enter a value for loop count."); return None
                repeat_count = int(repeat_count_str)
                if repeat_count < 1: raise ValueError("Loop count must be at least 1.")
            except ValueError: QMessageBox.warning(self, "Input Error", "Please enter a valid positive integer for loop count."); return None
            count_config = {"type": "hardcoded", "value": repeat_count}
        assign_iter_var_name: Optional[str] = None
        if self.assign_iter_checkbox.isChecked():
            if self.global_var_combo_assign_iter.currentIndex() == 0:
                new_var_name = self.new_var_iter_editor.text().strip()
                if not new_var_name: QMessageBox.warning(self, "Input Error", "Please enter a new variable name to assign the iteration count to."); return None
                assign_iter_var_name = new_var_name
            else: assign_iter_var_name = self.global_var_combo_assign_iter.currentText()
            if count_config["type"] == "variable" and count_config["value"] == assign_iter_var_name: QMessageBox.warning(self, "Input Error", "The variable for Loop Count cannot be the same as the variable for assigning Current Iteration."); return None
        return {"loop_name": loop_name if loop_name else None, "iteration_count_config": count_config, "assign_iteration_to_variable": assign_iter_var_name}

    def set_config(self, config: Dict[str, Any]) -> None:
        self.loop_name_editor.setText(config.get("loop_name", "") or "")
        count_config = config.get("iteration_count_config", {})
        if count_config.get("type") == "variable":
            self.use_var_checkbox.setChecked(True)
            idx = self.global_var_combo_count.findText(count_config["value"])
            if idx != -1: self.global_var_combo_count.setCurrentIndex(idx)
        else:
            self.use_var_checkbox.setChecked(False)
            self.repeat_count_editor.setText(str(count_config.get("value", 1)))
        self._toggle_count_var_input()
        assign_iter_var_name = config.get("assign_iteration_to_variable")
        if assign_iter_var_name:
            self.assign_iter_checkbox.setChecked(True)
            idx = self.global_var_combo_assign_iter.findText(assign_iter_var_name)
            if idx != -1: self.global_var_combo_assign_iter.setCurrentIndex(idx); self.new_var_iter_editor.clear()
            else: self.global_var_combo_assign_iter.setCurrentIndex(0); self.new_var_iter_editor.setText(assign_iter_var_name)
        else: self.assign_iter_checkbox.setChecked(False)
        self._toggle_assign_iter_input()

class ConditionalConfigDialog(QDialog):
    def __init__(self, global_variables: Dict[str, Any], parent: Optional[QWidget] = None, initial_config: Optional[Dict[str, Any]] = None):
        super().__init__(parent)
        self.setWindowTitle("Configure Conditional Block (IF-ELSE)")
        self.setMinimumWidth(400)
        self.global_variables = global_variables
        main_layout, form_layout = QVBoxLayout(), QFormLayout()
        self.block_name_editor = QLineEdit()
        self.block_name_editor.setPlaceholderText("Optional: Enter a name for this conditional block")
        form_layout.addRow("Block Name:", self.block_name_editor)
        condition_group, condition_layout = QGroupBox("Condition"), QGridLayout()
        self.left_operand_source_combo = QComboBox(); self.left_operand_source_combo.addItems(["Hardcoded Value", "Global Variable"])
        self.left_operand_editor = QLineEdit()
        self.left_operand_var_combo = QComboBox(); self.left_operand_var_combo.addItem("-- Select Variable --"); self.left_operand_var_combo.addItems(sorted(global_variables.keys()))
        self.left_operand_source_combo.currentIndexChanged.connect(self._toggle_left_operand_input)
        condition_layout.addWidget(QLabel("Left Operand:"), 0, 0); condition_layout.addWidget(self.left_operand_source_combo, 0, 1); condition_layout.addWidget(self.left_operand_editor, 0, 2); condition_layout.addWidget(self.left_operand_var_combo, 0, 2)
        self.operator_combo = QComboBox(); self.operator_combo.addItems(['==', '!=', '<', '>', '<=', '>=', 'in', 'not in', 'is', 'is not'])
        condition_layout.addWidget(QLabel("Operator:"), 1, 0); condition_layout.addWidget(self.operator_combo, 1, 1, 1, 2)
        self.right_operand_source_combo = QComboBox(); self.right_operand_source_combo.addItems(["Hardcoded Value", "Global Variable"])
        self.right_operand_editor = QLineEdit()
        self.right_operand_var_combo = QComboBox(); self.right_operand_var_combo.addItem("-- Select Variable --"); self.right_operand_var_combo.addItems(sorted(global_variables.keys()))
        self.right_operand_source_combo.currentIndexChanged.connect(self._toggle_right_operand_input)
        condition_layout.addWidget(QLabel("Right Operand:"), 2, 0); condition_layout.addWidget(self.right_operand_source_combo, 2, 1); condition_layout.addWidget(self.right_operand_editor, 2, 2); condition_layout.addWidget(self.right_operand_var_combo, 2, 2)
        condition_group.setLayout(condition_layout)
        main_layout.addWidget(condition_group)
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept); button_box.rejected.connect(self.reject)
        main_layout.addLayout(form_layout); main_layout.addWidget(button_box)
        self.setLayout(main_layout)
        self._toggle_left_operand_input(); self._toggle_right_operand_input()
        if initial_config: self.set_config(initial_config)

    def _toggle_left_operand_input(self) -> None:
        is_using_var = (self.left_operand_source_combo.currentIndex() == 1)
        self.left_operand_editor.setVisible(not is_using_var)
        self.left_operand_var_combo.setVisible(is_using_var)
        if not is_using_var: self.left_operand_var_combo.setCurrentIndex(0)
        else: self.left_operand_editor.clear()

    def _toggle_right_operand_input(self) -> None:
        is_using_var = (self.right_operand_source_combo.currentIndex() == 1)
        self.right_operand_editor.setVisible(not is_using_var)
        self.right_operand_var_combo.setVisible(is_using_var)
        if not is_using_var: self.right_operand_var_combo.setCurrentIndex(0)
        else: self.right_operand_editor.clear()

    def _parse_value_from_string(self, value_str: str) -> Any:
        try:
            if value_str.lower() == 'none': return None
            if value_str.lower() == 'true': return True
            if value_str.lower() == 'false': return False
            if value_str.isdigit(): return int(value_str)
            if value_str.replace('.', '', 1).isdigit(): return float(value_str)
            if (value_str.startswith('{') and value_str.endswith('}')) or (value_str.startswith('[') and value_str.endswith(']')): return json.loads(value_str)
            return value_str
        except (ValueError, json.JSONDecodeError): return value_str

    def get_config(self) -> Optional[Dict[str, Any]]:
        block_name = self.block_name_editor.text().strip()
        left_op_config: Dict[str, Any] = {}
        if self.left_operand_source_combo.currentIndex() == 1:
            var_name = self.left_operand_var_combo.currentText()
            if var_name == "-- Select Variable --": QMessageBox.warning(self, "Input Error", "Please select a global variable for the left operand."); return None
            left_op_config = {"type": "variable", "value": var_name}
        else:
            value_str = self.left_operand_editor.text().strip()
            if not value_str and self.operator_combo.currentText() not in ['is', 'is not']: QMessageBox.information(self, "Input Hint", "Left operand is empty. It will be treated as an empty string.", QMessageBox.StandardButton.Ok)
            left_op_config = {"type": "hardcoded", "value": self._parse_value_from_string(value_str)}
        right_op_config: Dict[str, Any] = {}
        if self.right_operand_source_combo.currentIndex() == 1:
            var_name = self.right_operand_var_combo.currentText()
            if var_name == "-- Select Variable --": QMessageBox.warning(self, "Input Error", "Please select a global variable for the right operand."); return None
            right_op_config = {"type": "variable", "value": var_name}
        else:
            value_str = self.right_operand_editor.text().strip()
            if not value_str and self.operator_combo.currentText() not in ['is', 'is not']: QMessageBox.information(self, "Input Hint", "Right operand is empty. It will be treated as an empty string.", QMessageBox.StandardButton.Ok)
            right_op_config = {"type": "hardcoded", "value": self._parse_value_from_string(value_str)}
        operator = self.operator_combo.currentText()
        return {"block_name": block_name if block_name else None, "condition": {"left_operand": left_op_config, "operator": operator, "right_operand": right_op_config}}

    def set_config(self, config: Dict[str, Any]) -> None:
        self.block_name_editor.setText(config.get("block_name", "") or "")
        condition = config.get("condition", {})
        if not condition: return
        left_op = condition.get("left_operand", {})
        if left_op.get("type") == "variable":
            self.left_operand_source_combo.setCurrentIndex(1)
            idx = self.left_operand_var_combo.findText(left_op.get("value", ""));
            if idx != -1: self.left_operand_var_combo.setCurrentIndex(idx)
        else:
            self.left_operand_source_combo.setCurrentIndex(0)
            val_to_display = left_op.get("value", "")
            if not isinstance(val_to_display, str): val_to_display = repr(val_to_display)
            self.left_operand_editor.setText(str(val_to_display))
        idx = self.operator_combo.findText(condition.get("operator", "=="))
        if idx != -1: self.operator_combo.setCurrentIndex(idx)
        right_op = condition.get("right_operand", {})
        if right_op.get("type") == "variable":
            self.right_operand_source_combo.setCurrentIndex(1)
            idx = self.right_operand_var_combo.findText(right_op.get("value", ""))
            if idx != -1: self.right_operand_var_combo.setCurrentIndex(idx)
        else:
            self.right_operand_source_combo.setCurrentIndex(0)
            val_to_display = right_op.get("value", "")
            if not isinstance(val_to_display, str): val_to_display = repr(val_to_display)
            self.right_operand_editor.setText(str(val_to_display))
        self._toggle_left_operand_input(); self._toggle_right_operand_input()

class GlobalVariableDialog(QDialog):
    def __init__(self, variable_name: str = "", variable_value: Any = "", parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Add/Edit Global Variable")
        self.layout = QFormLayout(self)

        self.name_input = QLineEdit(variable_name)
        self.value_input = QLineEdit(str(variable_value))
        self.browse_button = QPushButton("Browse...")
        self.browse_button.clicked.connect(self._open_file_dialog)

        self.layout.addRow("Name:", self.name_input)
        self.value_layout = QHBoxLayout()
        self.value_layout.addWidget(self.value_input)
        self.value_layout.addWidget(self.browse_button)
        self.layout.addRow("Value:", self.value_layout)

        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        self.layout.addRow(self.buttons)

        self.name_input.editingFinished.connect(self._toggle_browse_button)
        self._toggle_browse_button()

    def _toggle_browse_button(self):
        """Shows or hides the browse button based on the variable name."""
        name = self.name_input.text()
        if name:
            self.browse_button.setVisible("link" in name.lower())
        else:
            self.browse_button.setVisible(False)

    def _open_file_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "All Files (*)")
        if file_path:
            self.value_input.setText(file_path)

    def get_variable_data(self) -> Optional[Tuple[str, str]]:
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "Input Error", "Variable name cannot be empty.")
            return None
        return name, self.value_input.text()

class ParameterInputDialog(QDialog):
    request_screenshot = pyqtSignal()
    update_image_filenames = pyqtSignal(list, str)

    def __init__(self, method_name: str, parameters_to_configure: Dict[str, Tuple[Any, Any]],
                 current_global_var_names: List[str], image_filenames: List[str],
                 gui_communicator: GuiCommunicator,
                 initial_parameters_config: Optional[Dict[str, Any]] = None, initial_assign_to_variable_name: Optional[str] = None, parent: Optional[QWidget] = None):

        super().__init__(parent)
        self.setWindowTitle(f"Configure Parameters for '{method_name}'")
        self.parameters_config: Dict[str, Any] = {}
        self.assign_to_variable_name: Optional[str] = None
        self.gui_communicator = gui_communicator
        self._parent_main_window = parent
        self._image_filenames = image_filenames
        self.param_editors: Dict[str, QLineEdit] = {}
        self.param_var_selectors: Dict[str, QComboBox] = {}
        self.param_value_source_combos: Dict[str, QComboBox] = {}
        self.file_selector_combos: Dict[str, QComboBox] = {}
        self.param_kinds: Dict[str, Any] = {}

        main_layout, form_layout = QVBoxLayoutDialog(), QFormLayout()
        initial_parameters_config = initial_parameters_config if initial_parameters_config is not None else {}

        for param_name, (default_value, param_kind) in parameters_to_configure.items():
            param_h_layout = QHBoxLayout()
            label = QLabel(f"{param_name}:")
            self.param_kinds[param_name] = param_kind

            val_to_display = ""
            value_source_combo = QComboBox()
            value_source_combo.addItems(["Hardcoded Value", "Global Variable"])
            self.param_value_source_combos[param_name] = value_source_combo

            hardcoded_editor = QLineEdit()
            self.param_editors[param_name] = hardcoded_editor

            variable_select_combo = QComboBox()
            variable_select_combo.addItem("-- Select Variable --")
            variable_select_combo.addItems(current_global_var_names)
            self.param_var_selectors[param_name] = variable_select_combo

            if "file_link" in param_name:
                browse_button = QPushButton("Browse...")
                browse_button.clicked.connect(lambda _, p_name=param_name: self._open_file_dialog_for_param(p_name))
                param_h_layout.addWidget(browse_button)

            if param_name == "current_file" and image_filenames is not None:
                file_source_combo = QComboBox()
                file_source_combo.addItems(["Select from Files", "Global Variable"])
                self.param_value_source_combos[param_name] = file_source_combo

                file_selector_combo = QComboBox()
                file_selector_combo.addItem("-- Select File --")
                file_selector_combo.addItems(sorted(self._image_filenames))
                file_selector_combo.addItem("--- Add new image ---")
                self.file_selector_combos[param_name] = file_selector_combo

                variable_select_combo = QComboBox()
                variable_select_combo.addItem("-- Select Variable --")
                variable_select_combo.addItems(current_global_var_names)
                self.param_var_selectors[param_name] = variable_select_combo

                file_source_combo.currentIndexChanged.connect(lambda index, f_sel=file_selector_combo, v_sel=variable_select_combo: self._toggle_file_or_var_input(index, f_sel, v_sel))
                file_selector_combo.currentIndexChanged.connect(lambda index, p_name=param_name: self._on_file_selection_changed(p_name, index))
                self.update_image_filenames.connect(self._refresh_file_selector_combo)

                param_h_layout.addWidget(file_source_combo, 0)
                param_h_layout.addWidget(file_selector_combo, 2)
                param_h_layout.addWidget(variable_select_combo, 1)

                if param_name in initial_parameters_config:
                    config = initial_parameters_config[param_name]
                    if config.get('type') == 'hardcoded_file':
                        file_source_combo.setCurrentIndex(0)
                        idx = file_selector_combo.findText(config['value'])
                        if idx != -1:
                            file_selector_combo.setCurrentIndex(idx)
                            self._on_file_selection_changed(param_name, idx)
                    elif config.get('type') == 'variable':
                        file_source_combo.setCurrentIndex(1)
                        idx = variable_select_combo.findText(config['value'])
                        if idx != -1:
                            variable_select_combo.setCurrentIndex(idx)
                self._toggle_file_or_var_input(file_source_combo.currentIndex(), file_selector_combo, variable_select_combo)
            else:
                if param_name in initial_parameters_config:
                    config = initial_parameters_config[param_name]
                    if config.get('type') == 'hardcoded':
                        value_source_combo.setCurrentIndex(0)
                        val_to_display = config['value']
                        if not isinstance(val_to_display, str):
                            val_to_display = repr(val_to_display)
                        hardcoded_editor.setText(str(val_to_display))
                    elif config.get('type') == 'variable':
                        value_source_combo.setCurrentIndex(1)
                        idx = variable_select_combo.findText(config['value'])
                        if idx != -1:
                            variable_select_combo.setCurrentIndex(idx)
                        else:
                            QMessageBox.warning(self, "Warning", f"Global variable '{config['value']}' not found for parameter '{param_name}'.")
                elif default_value is not inspect.Parameter.empty:
                    val_to_display = default_value
                    if not isinstance(val_to_display, str):
                        val_to_display = repr(val_to_display)
                    hardcoded_editor.setText(str(val_to_display))
                elif param_kind == inspect.Parameter.VAR_POSITIONAL:
                    hardcoded_editor.setPlaceholderText("comma, separated, values")
                elif param_kind == inspect.Parameter.VAR_KEYWORD:
                    hardcoded_editor.setPlaceholderText('{"key": "value", "another": 123}')
                else:
                    hardcoded_editor.setPlaceholderText(f"Enter value for {param_name}")

                param_h_layout.insertWidget(0, value_source_combo)
                param_h_layout.insertWidget(1, hardcoded_editor, 2)
                param_h_layout.insertWidget(2, variable_select_combo, 1)

                value_source_combo.currentIndexChanged.connect(lambda index, editor=hardcoded_editor, selector=variable_select_combo: self._toggle_param_input_type(index, editor, selector))
                self._toggle_param_input_type(value_source_combo.currentIndex(), hardcoded_editor, variable_select_combo)

            form_layout.addRow(label, param_h_layout)

        main_layout.addLayout(form_layout)
        assignment_group_box = QGroupBox("Assign Method Result to Variable")
        assignment_layout = QVBoxLayoutDialog()

        self.assign_checkbox = QCheckBox("Assign result")
        self.assign_checkbox.stateChanged.connect(self._toggle_assignment_widgets)
        assignment_layout.addWidget(self.assign_checkbox)

        self.new_var_radio = QRadioButton("New Variable Name:")
        self.new_var_input = QLineEdit()
        self.new_var_input.setPlaceholderText("Enter new variable name")

        self.existing_var_radio = QRadioButton("Existing Variable:")
        self.existing_var_combo = QComboBox()
        self.existing_var_combo.addItem("-- Select Variable --")
        self.existing_var_combo.addItems(current_global_var_names)

        if initial_assign_to_variable_name:
            if initial_assign_to_variable_name in current_global_var_names:
                self.existing_var_radio.setChecked(True)
                idx = self.existing_var_combo.findText(initial_assign_to_variable_name)
                if idx != -1:
                    self.existing_var_combo.setCurrentIndex(idx)
            else:
                self.new_var_radio.setChecked(True)
                self.new_var_input.setText(initial_assign_to_variable_name)
            self.assign_checkbox.setChecked(True)
        else:
            self.new_var_radio.setChecked(True)

        self.new_var_radio.toggled.connect(self._toggle_assignment_widgets)
        self.existing_var_radio.toggled.connect(self._toggle_assignment_widgets)

        assignment_form_layout = QFormLayout()
        assignment_form_layout.addRow(self.new_var_radio, self.new_var_input)
        assignment_form_layout.addRow(self.existing_var_radio, self.existing_var_combo)
        assignment_layout.addLayout(assignment_form_layout)
        assignment_group_box.setLayout(assignment_layout)
        main_layout.addWidget(assignment_group_box)

        self._toggle_assignment_widgets()

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        main_layout.addWidget(button_box)
        self.setLayout(main_layout)

    def _open_file_dialog_for_param(self, param_name: str):
        editor = self.param_editors.get(param_name)
        if editor:
            file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "All Files (*)")
            if file_path:
                editor.setText(file_path)

    def _toggle_param_input_type(self, index: int, hardcoded_editor: QLineEdit, variable_select_combo: QComboBox) -> None:
        is_hardcoded = (index == 0)
        hardcoded_editor.setVisible(is_hardcoded)
        variable_select_combo.setVisible(not is_hardcoded)
        if not is_hardcoded:
            hardcoded_editor.clear()

    def _toggle_file_or_var_input(self, index: int, file_selector_combo: QComboBox, variable_select_combo: QComboBox) -> None:
        is_file_selection = (index == 0)
        file_selector_combo.setVisible(is_file_selection)
        variable_select_combo.setVisible(not is_file_selection)
        if not is_file_selection:
            file_selector_combo.setCurrentIndex(0)

    def _on_file_selection_changed(self, param_name: str, index: int) -> None:
        if param_name == "current_file" and param_name in self.file_selector_combos:
            file_selector_combo = self.file_selector_combos[param_name]
            selected_text = file_selector_combo.currentText()
            if selected_text == "--- Add new image ---":
                file_selector_combo.blockSignals(True)
                file_selector_combo.setCurrentIndex(0)
                file_selector_combo.blockSignals(False)
                self.request_screenshot.emit()
            elif selected_text != "-- Select File --":
                self.gui_communicator.update_module_info_signal.emit(selected_text)
            else:
                self.gui_communicator.update_module_info_signal.emit("")

    def _refresh_file_selector_combo(self, new_filenames: List[str], saved_filename: str = "") -> None:
        for param_name, combo_box in self.file_selector_combos.items():
            if param_name == "current_file":
                combo_box.blockSignals(True)
                combo_box.clear()
                combo_box.addItem("-- Select File --")
                combo_box.addItems(sorted(new_filenames))
                combo_box.addItem("--- Add new image ---")
                if saved_filename:
                    idx = combo_box.findText(saved_filename)
                    if idx != -1:
                        combo_box.setCurrentIndex(idx)
                        self.gui_communicator.update_module_info_signal.emit(saved_filename)
                combo_box.blockSignals(False)
                if combo_box.currentIndex() == 0:
                    self.gui_communicator.update_module_info_signal.emit("")
                else:
                    self.gui_communicator.update_module_info_signal.emit(combo_box.currentText())

    def _toggle_assignment_widgets(self) -> None:
        is_assign_enabled = self.assign_checkbox.isChecked()
        self.new_var_radio.setVisible(is_assign_enabled)
        self.new_var_input.setVisible(is_assign_enabled and self.new_var_radio.isChecked())
        self.existing_var_radio.setVisible(is_assign_enabled)
        self.existing_var_combo.setVisible(is_assign_enabled and self.existing_var_radio.isChecked())
        if not is_assign_enabled:
            self.new_var_input.clear()
            self.existing_var_combo.setCurrentIndex(0)

    def get_parameters_config(self) -> Optional[Dict[str, Any]]:
        config_data: Dict[str, Any] = {}
        for param_name, source_combo in self.param_value_source_combos.items():
            if param_name == "current_file" and param_name in self.file_selector_combos:
                value_source_index = source_combo.currentIndex()
                if value_source_index == 0:
                    file_name = self.file_selector_combos[param_name].currentText()
                    if file_name in ["-- Select File --", "--- Add new image ---"]:
                        QMessageBox.warning(self, "Input Error", f"Please select a file for parameter '{param_name}'.")
                        return None
                    config_data[param_name] = {'type': 'hardcoded_file', 'value': file_name}
                else:
                    var_name = self.param_var_selectors[param_name].currentText()
                    if var_name == "-- Select Variable --":
                        QMessageBox.warning(self, "Input Error", f"Please select a global variable for parameter '{param_name}'.")
                        return None
                    config_data[param_name] = {'type': 'variable', 'value': var_name}
                continue

            value_source_index = source_combo.currentIndex()
            if value_source_index == 0:
                value_str = self.param_editors[param_name].text().strip()
                param_kind = self.param_kinds[param_name]
                try:
                    parsed_value = None
                    if value_str.lower() == 'none':
                        parsed_value = None
                    elif param_kind == inspect.Parameter.VAR_POSITIONAL:
                        parsed_value = [item.strip() for item in value_str.split(',') if item.strip()]
                    elif param_kind == inspect.Parameter.VAR_KEYWORD:
                        parsed_value = json.loads(value_str)
                    elif value_str.lower() == 'true':
                        parsed_value = True
                    elif value_str.lower() == 'false':
                        parsed_value = False
                    elif value_str.isdigit():
                        parsed_value = int(value_str)
                    elif value_str.replace('.', '', 1).isdigit():
                        parsed_value = float(value_str)
                    else:
                        if (value_str.startswith('{') and value_str.endswith('}')) or (value_str.startswith('[') and value_str.endswith(']')):
                            try:
                                parsed_value = json.loads(value_str)
                            except json.JSONDecodeError:
                                parsed_value = value_str
                        else:
                            parsed_value = value_str
                except (ValueError, json.JSONDecodeError) as e:
                    QMessageBox.warning(self, "Input Error", f"Could not parse value '{value_str}' for '{param_name}' ({param_kind}): {e}. Storing as string.")
                    parsed_value = value_str
                config_data[param_name] = {'type': 'hardcoded', 'value': parsed_value}
            else:
                var_name = self.param_var_selectors[param_name].currentText()
                if var_name == "-- Select Variable --":
                    QMessageBox.warning(self, "Input Error", f"Please select a global variable for parameter '{param_name}'.")
                    return None
                config_data[param_name] = {'type': 'variable', 'value': var_name}

        self.parameters_config = config_data
        return config_data

    def get_assignment_variable(self) -> Optional[str]:
        if not self.assign_checkbox.isChecked():
            return None
        if self.new_var_radio.isChecked():
            var_name = self.new_var_input.text().strip()
            if not var_name:
                QMessageBox.warning(self, "Input Error", "New variable name cannot be empty for assignment.")
                return None
            return var_name
        else:
            var_name = self.existing_var_combo.currentText()
            if var_name == "-- Select Variable --":
                QMessageBox.warning(self, "Input Error", "Please select an existing variable to assign the result to.")
                return None
            return var_name

class ExecutionWorker(QThread):
    execution_started = pyqtSignal(str)
    execution_progress = pyqtSignal(int)
    execution_item_started = pyqtSignal(dict, int)
    execution_item_finished = pyqtSignal(dict, str, int)
    execution_error = pyqtSignal(dict, str, int)
    execution_finished_all = pyqtSignal(ExecutionContext, bool, int)

    def __init__(self, steps_to_execute: List[Dict[str, Any]], module_directory: str, gui_communicator: GuiCommunicator,
                 global_variables_ref: Dict[str, Any], parent: Optional[QWidget] = None,
                 single_step_mode: bool = False, selected_start_index: int = 0):
        super().__init__(parent)
        self.steps_to_execute = steps_to_execute
        self.module_directory = module_directory
        self.click_image_dir = os.path.normpath(os.path.join(module_directory, "..", "Click_image"))
        self.instantiated_objects: Dict[Tuple[str, str], Any] = {}
        self.context = ExecutionContext()
        self.context.set_gui_communicator(gui_communicator)
        self.global_variables = global_variables_ref
        self._is_stopped = False
        self.loop_stack: List[Dict[str, Any]] = []
        self.conditional_stack: List[Dict[str, Any]] = []
        self.single_step_mode = single_step_mode
        self.selected_start_index = selected_start_index
        self.next_step_index_to_select: int = -1

    def _resolve_loop_count(self, loop_config: Dict[str, Any]) -> int:
        count_config = loop_config["iteration_count_config"]
        if count_config["type"] == "variable":
            var_name = count_config["value"]
            var_value = self.global_variables.get(var_name)
            if isinstance(var_value, int) and var_value >= 1: return var_value
            else: self.context.add_log(f"Warning: Global variable '{var_name}' for loop count is not a valid positive integer (value: {var_value}). Defaulting to 1 iteration."); return 1
        else: return count_config.get("value", 1)

    def _resolve_operand_value(self, operand_config: Dict[str, Any]) -> Any:
        if operand_config["type"] == "variable":
            var_name = operand_config["value"]
            if var_name not in self.global_variables: raise ValueError(f"Global variable '{var_name}' not found for condition operand.")
            return self.global_variables[var_name]
        else: return operand_config["value"]

    def _evaluate_condition(self, condition_config: Dict[str, Any]) -> bool:
        left_val = self._resolve_operand_value(condition_config["left_operand"])
        right_val = self._resolve_operand_value(condition_config["right_operand"])
        operator = condition_config["operator"]
        try:
            if operator == '==': return left_val == right_val
            elif operator == '!=': return left_val != right_val
            elif operator == '<': return left_val < right_val
            elif operator == '>': return left_val > right_val
            elif operator == '<=': return left_val <= right_val
            elif operator == '>=': return left_val >= right_val
            elif operator == 'in': return left_val in right_val
            elif operator == 'not in': return left_val not in right_val
            elif operator == 'is': return left_val is right_val
            elif operator == 'is not': return left_val is not right_val
            else: raise ValueError(f"Unknown operator: {operator}")
        except Exception as e: raise ValueError(f"Error evaluating condition '{left_val} {operator} {right_val}': {e}")

    def run(self) -> None:
        self.execution_started.emit("Starting execution...")
        if not self.steps_to_execute: self.execution_finished_all.emit(self.context, False, -1); return
        total_steps_for_progress = len(self.steps_to_execute) * 2 if not self.single_step_mode else 1
        if total_steps_for_progress == 0: total_steps_for_progress = 1
        current_execution_item_count = 0
        original_sys_path = sys.path[:]
        if self.module_directory not in sys.path: sys.path.insert(0, self.module_directory)
        self.loop_stack = []
        self.conditional_stack = []
        step_index = self.selected_start_index if self.single_step_mode else 0
        original_listbox_row_index = 0
        try:
            while step_index < len(self.steps_to_execute):
                if self._is_stopped: break
                if self.single_step_mode and current_execution_item_count >= 1: break
                step_data = self.steps_to_execute[step_index]; step_type = step_data["type"]
                original_listbox_row_index = step_data.get("original_listbox_row_index", step_index)
                current_execution_item_count += 1
                progress_percentage = int((current_execution_item_count / total_steps_for_progress) * 100)
                self.execution_progress.emit(min(progress_percentage, 100))
                is_skipping = False
                if self.conditional_stack:
                    current_if = self.conditional_stack[-1]
                    # Use .get() to safely access keys that might not exist in a skipped block
                    if not current_if.get('condition_result', True) and not current_if.get('else_taken', False):
                        if not (step_type == "ELSE" and step_data.get("if_id") == current_if["if_id"]): is_skipping = True
                    elif current_if.get('condition_result', False) and current_if.get('else_taken', False):
                        if not (step_type == "IF_END" and step_data.get("if_id") == current_if["if_id"]): is_skipping = True
                if is_skipping:
                    self.execution_item_started.emit(step_data, original_listbox_row_index)
                    if step_type == "IF_START": self.conditional_stack.append({'if_id': step_data['if_id'], 'skipped_marker': True})
                    elif step_type == "loop_start": self.loop_stack.append({'loop_id': step_data['loop_id'], 'skipped_marker': True})
                    elif step_type == "IF_END":
                        if self.conditional_stack and self.conditional_stack[-1].get('skipped_marker'): self.conditional_stack.pop()
                    elif step_type == "loop_end":
                        if self.loop_stack and self.loop_stack[-1].get('skipped_marker'): self.loop_stack.pop()
                    self.execution_item_finished.emit(step_data, "SKIPPED", original_listbox_row_index)
                    step_index += 1
                    continue
                self.execution_item_started.emit(step_data, original_listbox_row_index)
                QThread.msleep(50)

                if step_type in ["group_start", "group_end"]:
                    self.execution_item_finished.emit(step_data, "Organizational Step", original_listbox_row_index)
                    step_index += 1
                    continue

                if step_type == "step":
                    class_name, method_name, module_name = step_data["class_name"], step_data["method_name"], step_data["module_name"]
                    parameters_config = step_data["parameters_config"]; assign_to_variable_name = step_data["assign_to_variable_name"]
                    resolved_parameters, params_str_debug = {}, []
                    for param_name, config in parameters_config.items():
                        if param_name == "original_listbox_row_index": continue
                        if config['type'] == 'hardcoded': resolved_parameters[param_name] = config['value']; params_str_debug.append(f"{param_name}={repr(config['value'])}")
                        elif config['type'] == 'hardcoded_file': resolved_parameters[param_name] = config['value']; params_str_debug.append(f"{param_name}=FILE('{config['value']}')")
                        elif config['type'] == 'variable':
                            var_name = config['value']
                            if var_name in self.global_variables: resolved_parameters[param_name] = self.global_variables[var_name]; params_str_debug.append(f"{param_name}=@{var_name}({repr(self.global_variables[var_name])})")
                            else: raise ValueError(f"Global variable '{var_name}' not found for parameter '{param_name}'.")
                    self.context.add_log(f"Executing: {class_name}.{method_name}({', '.join(params_str_debug)})")
                    try:
                        module = importlib.import_module(module_name); importlib.reload(module)
                        class_obj = getattr(module, class_name); instance_key = (class_name, module_name)
                        if instance_key not in self.instantiated_objects:
                            init_kwargs = {};
                            if 'context' in inspect.signature(class_obj.__init__).parameters: init_kwargs['context'] = self.context
                            self.instantiated_objects[instance_key] = class_obj(**init_kwargs)
                        instance = self.instantiated_objects[instance_key]; method_func = getattr(instance, method_name)
                        method_kwargs = {k:v for k,v in resolved_parameters.items()}
                        if 'context' in inspect.signature(method_func).parameters: method_kwargs['context'] = self.context
                        result = method_func(**method_kwargs)
                        if assign_to_variable_name: self.global_variables[assign_to_variable_name] = result; result_msg = f"Result: {result} (Assigned to @{assign_to_variable_name})"
                        else: result_msg = f"Result: {result}"
                        self.execution_item_finished.emit(step_data, result_msg, original_listbox_row_index)
                    except Exception as e: raise e
                elif step_type == "loop_start":
                    loop_id, loop_config = step_data["loop_id"], step_data["loop_config"]
                    is_new_loop = not (self.loop_stack and self.loop_stack[-1].get('loop_id') == loop_id)
                    if is_new_loop:
                        loop_end_index, nesting_level = -1, 0
                        for i in range(step_index + 1, len(self.steps_to_execute)):
                            s, s_type = self.steps_to_execute[i], self.steps_to_execute[i].get("type")
                            if s_type in ["loop_start", "IF_START"]: nesting_level += 1
                            elif s_type == "loop_end" and s.get("loop_id") == loop_id and nesting_level == 0: loop_end_index = i; break
                            elif s_type in ["loop_end", "IF_END"] and nesting_level > 0: nesting_level -= 1
                        if loop_end_index == -1: raise ValueError(f"Loop '{loop_id}' has no matching 'loop_end' marker.")
                        total_iterations = self._resolve_loop_count(loop_config)
                        current_loop_info = {'loop_id': loop_id, 'start_index': step_index, 'end_index': loop_end_index, 'current_iteration': 1, 'total_iterations': total_iterations, 'loop_config': loop_config}
                        self.loop_stack.append(current_loop_info)
                    else:
                        current_loop_info = self.loop_stack[-1]; current_loop_info['current_iteration'] += 1
                        current_loop_info['total_iterations'] = self._resolve_loop_count(loop_config)
                    if current_loop_info['current_iteration'] > current_loop_info['total_iterations']: self.loop_stack.pop(); step_index = current_loop_info['end_index']; self.execution_item_finished.emit(step_data, "Loop Finished", original_listbox_row_index)
                    else:
                        assign_var = loop_config.get("assign_iteration_to_variable")
                        if assign_var: self.global_variables[assign_var] = current_loop_info['current_iteration']; self.context.add_log(f"Assigned iteration {current_loop_info['current_iteration']} to @{assign_var}")
                        self.execution_item_finished.emit(step_data, f"Iter {current_loop_info['current_iteration']}/{current_loop_info['total_iterations']}", original_listbox_row_index)
                elif step_type == "loop_end":
                    if not self.loop_stack or self.loop_stack[-1].get('loop_id') != step_data['loop_id']: raise ValueError(f"Mismatched loop_end for ID: {step_data['loop_id']}")
                    step_index = self.loop_stack[-1]['start_index'] - 1; self.execution_item_finished.emit(step_data, "Looping...", original_listbox_row_index)
                elif step_type == "IF_START":
                    if_id, condition_config = step_data["if_id"], step_data["condition_config"]
                    condition_result = self._evaluate_condition(condition_config["condition"]); self.context.add_log(f"IF '{if_id}' evaluated: {condition_result}")
                    self.conditional_stack.append({'if_id': if_id, 'condition_result': condition_result, 'else_taken': False}); self.execution_item_finished.emit(step_data, f"Condition: {condition_result}", original_listbox_row_index)
                elif step_type == "ELSE":
                    if not self.conditional_stack or self.conditional_stack[-1].get('if_id') != step_data['if_id']: raise ValueError(f"Mismatched ELSE for ID: {step_data['if_id']}")
                    current_if = self.conditional_stack[-1]; current_if['else_taken'] = True
                    self.execution_item_finished.emit(step_data, "Branching", original_listbox_row_index)
                    if current_if.get('condition_result', False):
                        nesting_level = 0
                        for i in range(step_index + 1, len(self.steps_to_execute)):
                            s, s_type = self.steps_to_execute[i], self.steps_to_execute[i].get("type")
                            if s_type in ["loop_start", "IF_START"]: nesting_level += 1
                            elif s_type == "IF_END" and s.get("if_id") == current_if["if_id"] and nesting_level == 0: step_index = i-1; break
                            elif s_type in ["loop_end", "IF_END"] and nesting_level > 0: nesting_level -= 1
                elif step_type == "IF_END":
                    if not self.conditional_stack or self.conditional_stack[-1].get('if_id') != step_data['if_id']: raise ValueError(f"Mismatched IF_END for ID: {step_data['if_id']}")
                    self.conditional_stack.pop(); self.execution_item_finished.emit(step_data, "End of Conditional", original_listbox_row_index)
                step_index += 1
        except Exception as e:
            error_msg = f"Error at step {original_listbox_row_index+1}: {type(e).__name__}: {e}"; self.context.add_log(error_msg)
            self.execution_error.emit(self.steps_to_execute[step_index] if step_index < len(self.steps_to_execute) else {}, error_msg, original_listbox_row_index); self._is_stopped = True
        finally:
            sys.path = original_sys_path
            if self.single_step_mode:
                next_index = -1
                if not self._is_stopped and step_index < len(self.steps_to_execute): next_index = self.steps_to_execute[step_index].get("original_listbox_row_index", -1)
                self.next_step_index_to_select = next_index
            self.execution_finished_all.emit(self.context, self._is_stopped, self.next_step_index_to_select)

class StepInsertionDialog(QDialog):
    def __init__(self, execution_tree: QTreeWidget, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Insert Step At...")
        self.layout = QVBoxLayout(self)
        self.execution_tree_view = QTreeWidget()
        self.execution_tree_view.setHeaderHidden(True)
        self.execution_tree_view.setSelectionMode(QTreeWidget.SelectionMode.SingleSelection)
        self.insertion_mode_group = QGroupBox("Insertion Mode")
        self.insertion_mode_layout = QHBoxLayout()
        self.insert_as_child_radio = QRadioButton("Insert as Child")
        self.insert_before_radio = QRadioButton("Insert Before Selected")
        self.insert_after_radio = QRadioButton("Insert After Selected")
        self.insert_as_child_radio.setChecked(True)
        self.insertion_mode_layout.addWidget(self.insert_as_child_radio)
        self.insertion_mode_layout.addWidget(self.insert_before_radio)
        self.insertion_mode_layout.addWidget(self.insert_after_radio)
        self.insertion_mode_group.setLayout(self.insertion_mode_layout)
        self.insert_as_child_radio.toggled.connect(self._update_insertion_options)
        self.insert_before_radio.toggled.connect(self._update_insertion_options)
        self.insert_after_radio.toggled.connect(self._update_insertion_options)
        self.layout.addWidget(QLabel("Select Parent or Sibling for new step:"))
        self.layout.addWidget(self.execution_tree_view)
        self.layout.addWidget(self.insertion_mode_group)
        self._populate_tree_view(execution_tree)
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.layout.addWidget(self.button_box)
        self.selected_parent_item: Optional[QTreeWidgetItem] = None
        self.insert_mode: str = "child"
        self.execution_tree_view.currentItemChanged.connect(self._update_insertion_options)
        self._update_insertion_options()

    def _get_item_data(self, item: QTreeWidgetItem) -> Optional[Dict[str, Any]]:
        if not item: return None
        data = item.data(0, Qt.ItemDataRole.UserRole)
        return data.value() if isinstance(data, QVariant) else data

    def _populate_tree_view(self, source_tree: QTreeWidget):
        self.execution_tree_view.clear()
        self._copy_children(source_tree.invisibleRootItem(), self.execution_tree_view.invisibleRootItem())
        self.execution_tree_view.expandAll()

    def _copy_children(self, source_item: QTreeWidgetItem, dest_item: QTreeWidgetItem):
        for i in range(source_item.childCount()):
            child_source = source_item.child(i)
            card_widget = self.parent().execution_tree.itemWidget(child_source, 0)
            display_text = f"{card_widget.step_label.text()}" if card_widget else ""
            child_dest = QTreeWidgetItem(dest_item, [display_text])
            child_dest.setData(0, Qt.ItemDataRole.UserRole, self._get_item_data(child_source))
            self._copy_children(child_source, child_dest)

    def _update_insertion_options(self) -> None:
        selected_item = self.execution_tree_view.currentItem()
        if selected_item is None:
            self.insert_as_child_radio.setEnabled(True); self.insert_before_radio.setEnabled(False); self.insert_after_radio.setEnabled(False); self.insert_as_child_radio.setChecked(True)
        else:
            step_data_dict = self._get_item_data(selected_item)
            is_block_parent = step_data_dict and step_data_dict.get("type") in ["loop_start", "IF_START", "ELSE", "group_start"]
            self.insert_as_child_radio.setEnabled(is_block_parent); self.insert_before_radio.setEnabled(True); self.insert_after_radio.setEnabled(True)
            if self.insert_as_child_radio.isChecked() and not self.insert_as_child_radio.isEnabled():
                if self.insert_after_radio.isEnabled(): self.insert_after_radio.setChecked(True)
                elif self.insert_before_radio.isEnabled(): self.insert_before_radio.setChecked(True)
            if (self.insert_before_radio.isChecked() or self.insert_after_radio.isChecked()) and not (self.insert_before_radio.isEnabled() and not self.insert_after_radio.isEnabled()):
                if self.insert_as_child_radio.isEnabled(): self.insert_as_child_radio.setChecked(True)
        if not (self.insert_as_child_radio.isChecked() or self.insert_before_radio.isChecked() or self.insert_after_radio.isChecked()):
            if self.insert_as_child_radio.isEnabled(): self.insert_as_child_radio.setChecked(True)
            elif self.insert_after_radio.isEnabled(): self.insert_after_radio.setChecked(True)
            elif self.insert_before_radio.isEnabled(): self.insert_before_radio.setChecked(True)

    def get_insertion_point(self) -> Tuple[Optional[QTreeWidgetItem], str]:
        self.selected_parent_item = self.execution_tree_view.currentItem()
        if self.insert_before_radio.isChecked(): self.insert_mode = "before"
        elif self.insert_after_radio.isChecked(): self.insert_mode = "after"
        else: self.insert_mode = "child"
        return self.selected_parent_item, self.insert_mode

# --- REPLACEMENT CLASS: TemplateVariableMappingDialog (with intelligent default) ---
class TemplateVariableMappingDialog(QDialog):
    """A dialog to map variables from a template to global variables."""
    def __init__(self, template_variables: set, global_variables: list, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Map Template Variables")
        self.setMinimumWidth(500)
        self.template_variables = sorted(list(template_variables))
        self.global_variables = global_variables
        self.mapping_widgets = {}

        main_layout = QVBoxLayout(self)
        form_layout = QFormLayout()

        for var_name in self.template_variables:
            row_layout = QHBoxLayout()
            
            action_combo = QComboBox()
            action_combo.addItems(["Map to Existing", "Create New", "Keep Original Name"])
            
            existing_var_combo = QComboBox()
            existing_var_combo.addItem("-- Select Existing --")
            existing_var_combo.addItems(self.global_variables)
            
            # --- INTELLIGENT DEFAULT LOGIC ---
            # If a global variable already exists with the same name as the template variable,
            # default to "Map to Existing" and select that variable.
            if var_name in self.global_variables:
                action_combo.setCurrentText("Map to Existing")
                existing_var_combo.setCurrentText(var_name)
            else:
                # Otherwise, default to "Keep Original Name"
                action_combo.setCurrentText("Keep Original Name")

            new_var_editor = QLineEdit(var_name)

            row_layout.addWidget(action_combo, 1)
            row_layout.addWidget(existing_var_combo, 2)
            row_layout.addWidget(new_var_editor, 2)
            
            form_layout.addRow(QLabel(f"'{var_name}':"), row_layout)

            self.mapping_widgets[var_name] = {
                'action': action_combo,
                'existing': existing_var_combo,
                'new': new_var_editor
            }
            
            action_combo.currentIndexChanged.connect(
                lambda index, v=var_name: self._toggle_inputs(v, index)
            )
            self._toggle_inputs(var_name, action_combo.currentIndex())

        main_layout.addLayout(form_layout)
        
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        main_layout.addWidget(button_box)

    def _toggle_inputs(self, var_name: str, index: int):
        widgets = self.mapping_widgets[var_name]
        is_mapping = (index == 0)
        widgets['existing'].setVisible(is_mapping)
        widgets['new'].setVisible(not is_mapping)

    def get_mapping(self) -> Optional[Tuple[Dict[str, str], Dict[str, Any]]]:
        mapping = {}
        new_variables = {}
        for var_name, widgets in self.mapping_widgets.items():
            action_index = widgets['action'].currentIndex()
            
            if action_index == 0: # Map to Existing
                target_var = widgets['existing'].currentText()
                if target_var == "-- Select Existing --":
                    QMessageBox.warning(self, "Input Error", f"Please select an existing variable to map '{var_name}' to.")
                    return None
                mapping[var_name] = target_var
            else: # Create New or Keep Original
                target_var = widgets['new'].text().strip()
                if not target_var:
                    QMessageBox.warning(self, "Input Error", f"The new variable name for '{var_name}' cannot be empty.")
                    return None
                mapping[var_name] = target_var
                if target_var not in self.global_variables:
                    new_variables[target_var] = None

        return mapping, new_variables

# --- NEW CLASS: SaveTemplateDialog ---
class SaveTemplateDialog(QDialog):
    def __init__(self, existing_templates: List[str], parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Save Step Template")
        self.setMinimumWidth(350)
        
        main_layout = QVBoxLayout(self)
        form_layout = QFormLayout()

        self.info_label = QLabel("Select an existing template to overwrite, or type a new name.")
        
        self.templates_combo = QComboBox()
        self.templates_combo.addItem("-- Create New Template --")
        self.templates_combo.addItems(sorted(existing_templates))

        self.name_editor = QLineEdit()
        self.name_editor.setPlaceholderText("Enter new template name")

        self.templates_combo.currentTextChanged.connect(self._selection_changed)

        form_layout.addRow("Action:", self.templates_combo)
        form_layout.addRow("Template Name:", self.name_editor)

        main_layout.addWidget(self.info_label)
        main_layout.addLayout(form_layout)
        
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        main_layout.addWidget(button_box)

        self._selection_changed(self.templates_combo.currentText())

    def _selection_changed(self, text: str) -> None:
        if text == "-- Create New Template --":
            self.name_editor.clear()
            self.name_editor.setEnabled(True)
            self.name_editor.setPlaceholderText("Enter new template name")
        else:
            self.name_editor.setText(text)
            self.name_editor.setEnabled(True)

    def get_template_name(self) -> Optional[str]:
        name = self.name_editor.text().strip()
        # Sanitize name
        sanitized_name = "".join(c for c in name if c.isalnum() or c in (' ', '_')).rstrip()
        if not sanitized_name:
            QMessageBox.warning(self, "Invalid Name", "Template name cannot be empty or contain only invalid characters.")
            return None
        return sanitized_name

# --- NEW CLASS: SaveBotDialog ---
class SaveBotDialog(QDialog):
    def __init__(self, existing_bots: List[str], parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Save Bot Steps")
        self.setMinimumWidth(350)
        
        main_layout = QVBoxLayout(self)
        form_layout = QFormLayout()

        self.info_label = QLabel("Select an existing bot to overwrite, or type a new name.")
        
        self.bots_combo = QComboBox()
        self.bots_combo.addItem("-- Create New Bot --")
        self.bots_combo.addItems(sorted(existing_bots))

        self.name_editor = QLineEdit()
        self.name_editor.setPlaceholderText("Enter new bot name")

        self.bots_combo.currentTextChanged.connect(self._selection_changed)

        form_layout.addRow("Action:", self.bots_combo)
        form_layout.addRow("Bot Name:", self.name_editor)

        main_layout.addWidget(self.info_label)
        main_layout.addLayout(form_layout)
        
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        main_layout.addWidget(button_box)

        self._selection_changed(self.bots_combo.currentText())

    def _selection_changed(self, text: str) -> None:
        if text == "-- Create New Bot --":
            self.name_editor.clear()
            self.name_editor.setEnabled(True)
            self.name_editor.setPlaceholderText("Enter new bot name")
        else:
            self.name_editor.setText(text)
            self.name_editor.setEnabled(True)

    def get_bot_name(self) -> Optional[str]:
        name = self.name_editor.text().strip()
        sanitized_name = "".join(c for c in name if c.isalnum() or c in (' ', '_')).rstrip()
        if not sanitized_name:
            QMessageBox.warning(self, "Invalid Name", "Bot name cannot be empty or contain only invalid characters.")
            return None
        return sanitized_name

# In main_app.py, REPLACE your existing GroupedTreeWidget class with this one

class GroupedTreeWidget(QTreeWidget):
    """
    A custom QTreeWidget that draws vertical lines to indicate the scope
    of expanded groups, loops, and conditional blocks.
    """
    def __init__(self, main_window: 'MainWindow', parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.main_window = main_window

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:
        # First, run the default paint event to draw all the items.
        super().paintEvent(event)
        
        # Now, draw our custom lines on top.
        painter = QPainter(self.viewport())
        pen = QPen(QColor("#C0392B")) # A shade of red
        pen.setWidth(2)
        painter.setPen(pen)

        self._draw_group_lines_recursive(self.invisibleRootItem(), painter)

    def _draw_group_lines_recursive(self, parent_item: QTreeWidgetItem, painter: QPainter):
        for i in range(parent_item.childCount()):
            child_item = parent_item.child(i)
            
            # Get the data associated with this tree item.
            item_data = self.main_window._get_item_data(child_item)
            
            # --- THIS IS THE CORRECTED LINE ---
            # Check if this item is a block start and is currently expanded.
            if (item_data and item_data.get("type") in ["group_start", "loop_start", "IF_START"] and child_item.isExpanded()):
                
                start_index = item_data.get("original_listbox_row_index")
                # Use MainWindow's helper to find the index of the matching 'end' block.
                _ , end_index = self.main_window._find_block_indices(start_index)
                
                # Find the corresponding end item widget using the map from MainWindow.
                end_item = self.main_window.data_to_item_map.get(end_index)
                
                start_rect = self.visualItemRect(child_item)
                
                # Proceed only if we found the end item and the start item is visible.
                if end_item and not start_rect.isEmpty():
                    end_rect = self.visualItemRect(end_item)
                    
                    # Define line geometry.
                    x_offset = 8    # How far from the left edge.
                    tick_width = 5  # The width of the small horizontal ticks.
                    start_y = start_rect.center().y()
                    
                    # If the end item is scrolled out of view, draw the line to the
                    # bottom of the visible area. Otherwise, draw to the end item's center.
                    end_y = end_rect.center().y() if not end_rect.isEmpty() else self.viewport().height()

                    # Don't draw if the line would be pointing upwards.
                    if end_y < start_y:
                        continue
                        
                    # --- Draw the lines ---
                    # 1. The main vertical line.
                    painter.drawLine(x_offset, start_y, x_offset, end_y)
                    # 2. The top horizontal tick.
                    painter.drawLine(x_offset, start_y, x_offset + tick_width, start_y)
                    # 3. The bottom horizontal tick (only if the end item is visible).
                    if not end_rect.isEmpty():
                        painter.drawLine(x_offset, end_y, x_offset + tick_width, end_y)

            # Recurse into the children of the current item to draw nested block lines.
            if child_item.isExpanded():
                self._draw_group_lines_recursive(child_item, painter)

# --- NEW WIDGET: FindReplaceWidget ---
class FindReplaceWidget(QWidget):
    """A widget for find and replace functionality."""
    def __init__(self, editor):
        super().__init__()
        self.editor = editor
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 5, 0, 5)

        find_layout = QHBoxLayout()
        self.find_input = QLineEdit()
        self.find_input.setPlaceholderText("Find")
        self.find_input.textChanged.connect(self.update_editor_selection)
        
        self.next_button = QPushButton("Next")
        self.prev_button = QPushButton("Previous")
        
        find_layout.addWidget(QLabel("Find:"))
        find_layout.addWidget(self.find_input)
        find_layout.addWidget(self.next_button)
        find_layout.addWidget(self.prev_button)
        
        replace_layout = QHBoxLayout()
        self.replace_input = QLineEdit()
        self.replace_input.setPlaceholderText("Replace with")
        
        self.replace_button = QPushButton("Replace")
        self.replace_all_button = QPushButton("Replace All")

        replace_layout.addWidget(QLabel("Replace:"))
        replace_layout.addWidget(self.replace_input)
        replace_layout.addWidget(self.replace_button)
        replace_layout.addWidget(self.replace_all_button)

        options_layout = QHBoxLayout()
        self.case_sensitive_check = QCheckBox("Case Sensitive")
        options_layout.addWidget(self.case_sensitive_check)
        options_layout.addStretch()

        main_layout.addLayout(find_layout)
        main_layout.addLayout(replace_layout)
        main_layout.addLayout(options_layout)

    def update_editor_selection(self):
        """Highlights all occurrences of the find text in the editor."""
        text = self.find_input.text()
        if not text:
            self.editor.setExtraSelections([])
            return

        selections = []
        cursor = self.editor.document().find(text, 0, self.get_find_flags())
        while not cursor.isNull():
            selection = QTextEdit.ExtraSelection()
            selection.format.setBackground(QColor("cyan"))
            selection.cursor = cursor
            selections.append(selection)
            cursor = self.editor.document().find(text, cursor, self.get_find_flags())
        
        self.editor.setExtraSelections(selections)

    def get_find_flags(self):
        """Returns the appropriate find flags based on UI options."""
        flags = QTextDocument.FindFlag(0)
        if self.case_sensitive_check.isChecked():
            flags |= QTextDocument.FindFlag.FindCaseSensitively
        return flags

# --- NEW CLASS: HtmlViewerDialog ---
class HtmlViewerDialog(QDialog):
    """A dialog to display HTML content."""
    def __init__(self, file_path: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle(f"Documentation: {os.path.basename(file_path)}")
        self.setGeometry(200, 200, 800, 600)

        main_layout = QVBoxLayout(self)

        self.text_browser = QTextBrowser()
        self.text_browser.setOpenExternalLinks(True) # Allow opening links in a web browser
        main_layout.addWidget(self.text_browser)

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        button_box.rejected.connect(self.reject)
        main_layout.addWidget(button_box)

        self.load_html(file_path)

    def load_html(self, file_path: str):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                html_content = f.read()
                self.text_browser.setHtml(html_content)
        except Exception as e:
            self.text_browser.setPlainText(f"Error loading documentation file:\n{e}")
            
# --- NEW CLASS: PythonHighlighter (No changes from before) ---
class PythonHighlighter(QSyntaxHighlighter):
    """A syntax highlighter for Python code."""
    def __init__(self, parent):
        super().__init__(parent)
        self.highlighting_rules = []

        # Keywords
        keyword_format = QTextCharFormat()
        keyword_format.setForeground(QColor("#B57627")) # Orange-brown
        keyword_format.setFontWeight(QFont.Weight.Bold)
        keywords = ["and", "as", "assert", "break", "class", "continue",
                    "def", "del", "elif", "else", "except", "False",
                    "finally", "for", "from", "global", "if", "import",
                    "in", "is", "lambda", "None", "nonlocal", "not",
                    "or", "pass", "raise", "return", "True", "try",
                    "while", "with", "yield"]
        for word in keywords:
            pattern = QRegularExpression(rf"\b{word}\b")
            self.highlighting_rules.append((pattern, keyword_format))

        # self
        self_format = QTextCharFormat()
        self_format.setForeground(QColor("#AF609F")) # Purple
        self.highlighting_rules.append((QRegularExpression(r"\bself\b"), self_format))

        # Decorators
        decorator_format = QTextCharFormat()
        decorator_format.setForeground(QColor("#7E9F43")) # Green
        self.highlighting_rules.append((QRegularExpression(r"@[A-Za-z0-9_]+"), decorator_format))

        # Strings
        string_format = QTextCharFormat()
        string_format.setForeground(QColor("#5D9248")) # Dark Green
        self.highlighting_rules.append((QRegularExpression(r"\"[^\"\\]*(\\.[^\"\\]*)*\""), string_format))
        self.highlighting_rules.append((QRegularExpression(r"'[^'\\]*(\\.[^'\\]*)*'"), string_format))

        # Comments
        comment_format = QTextCharFormat()
        comment_format.setForeground(QColor("#A0A0A0")) # Gray
        comment_format.setFontItalic(True)
        self.highlighting_rules.append((QRegularExpression(r"#[^\n]*"), comment_format))

    def highlightBlock(self, text):
        for pattern, format in self.highlighting_rules:
            match_iterator = pattern.globalMatch(text)
            while match_iterator.hasNext():
                match = match_iterator.next()
                self.setFormat(match.capturedStart(), match.capturedLength(), format)

# --- NEW WIDGET: CodeEditorWithLineNumbers ---
# --- REPLACEMENT CLASS: CodeEditorWithLineNumbers (Complete Version) ---
# This version includes all features: line numbers, zoom, tab handling,
# indentation guides, and syntax error highlighting.
# --- REPLACEMENT CLASS: CodeEditorWithLineNumbers (Final Corrected Version) ---

class CodeEditorWithLineNumbers(QPlainTextEdit):
    class LineNumberArea(QWidget):
        def __init__(self, editor):
            super().__init__(editor)
            self.codeEditor = editor

        def sizeHint(self):
            return QSize(self.codeEditor.lineNumberAreaWidth(), 0)

        def paintEvent(self, event):
            self.codeEditor.lineNumberAreaPaintEvent(event)

    def __init__(self):
        super().__init__()
        self.lineNumberArea = self.LineNumberArea(self)
        self.blockCountChanged.connect(self.updateLineNumberAreaWidth)
        self.updateRequest.connect(self.updateLineNumberArea)
        self.cursorPositionChanged.connect(self.highlightCurrentLine)
        
        self.font_metrics = self.fontMetrics()
        self.setTabStopDistance(self.font_metrics.horizontalAdvance(' ') * 4)

        self.updateLineNumberAreaWidth(0)
        self.highlightCurrentLine()

    def lineNumberAreaWidth(self):
        digits = 1
        max_num = max(1, self.blockCount())
        while max_num >= 10:
            max_num //= 10
            digits += 1
        space = 3 + self.fontMetrics().horizontalAdvance('9') * digits + 15
        return space

    def updateLineNumberAreaWidth(self, _):
        self.setViewportMargins(self.lineNumberAreaWidth(), 0, 0, 0)

    def updateLineNumberArea(self, rect, dy):
        if dy:
            self.lineNumberArea.scroll(0, dy)
        else:
            self.lineNumberArea.update(0, rect.y(), self.lineNumberArea.width(), rect.height())
        if rect.contains(self.viewport().rect()):
            self.updateLineNumberAreaWidth(0)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        cr = self.contentsRect()
        self.lineNumberArea.setGeometry(QRect(cr.left(), cr.top(), self.lineNumberAreaWidth(), cr.height()))

    def lineNumberAreaPaintEvent(self, event):
        painter = QPainter(self.lineNumberArea)
        painter.fillRect(event.rect(), QColor("#F0F0F0"))
        block = self.firstVisibleBlock()
        blockNumber = block.blockNumber()
        top = self.blockBoundingGeometry(block).translated(self.contentOffset()).top()
        bottom = top + self.blockBoundingRect(block).height()
        while block.isValid() and top <= event.rect().bottom():
            if block.isVisible() and bottom >= event.rect().top():
                number = f"- {blockNumber + 1}"
                painter.setPen(QColor("#A0A0A0"))
                painter.drawText(0, int(top), self.lineNumberArea.width() - 5, self.fontMetrics().height(),
                                 Qt.AlignmentFlag.AlignRight, number)
            block = block.next()
            top = bottom
            bottom = top + self.blockBoundingRect(block).height()
            blockNumber += 1

    # --- CORRECTED METHOD ---
    def highlightCurrentLine(self):
        extraSelections = []
        if not self.isReadOnly():
            selection = QTextEdit.ExtraSelection()
            lineColor = QColor(Qt.GlobalColor.yellow).lighter(160)
            selection.format.setBackground(lineColor)
            selection.format.setProperty(QTextFormat.Property.FullWidthSelection, True)
            selection.cursor = self.textCursor()
            selection.cursor.clearSelection()
            extraSelections.append(selection)
        
        # Keep existing error highlights when updating the current line highlight
        current_selections = self.extraSelections()
        for sel in current_selections:
            # FIX IS HERE: .color().name()
            if sel.format.background().color().name() == "#ffcccc":
                extraSelections.append(sel)
        
        self.setExtraSelections(extraSelections)

    def keyPressEvent(self, event: QKeyEvent):
        if event.key() == Qt.Key.Key_Tab:
            cursor = self.textCursor()
            cursor.insertText(" " * 4)
            return
        super().keyPressEvent(event)

    def paintEvent(self, event: QtGui.QPaintEvent):
        super().paintEvent(event)
        painter = QPainter(self.viewport())
        color = QColor("#D3D3D3")
        pen = QPen(color)
        pen.setStyle(Qt.PenStyle.DotLine)
        painter.setPen(pen)
        tab_width = self.tabStopDistance()
        left_margin = self.contentsRect().left()
        block = self.firstVisibleBlock()
        while block.isValid() and block.isVisible():
            geom = self.blockBoundingGeometry(block).translated(self.contentOffset())
            text = block.text()
            leading_spaces = len(text) - len(text.lstrip(' '))
            for i in range(1, (leading_spaces // 4) + 2):
                x = left_margin + (i * tab_width)
                if x > self.viewport().width(): break
                painter.drawLine(int(x), int(geom.top()), int(x), int(geom.bottom()))
            block = block.next()

    def wheelEvent(self, event: QtGui.QWheelEvent):
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            delta = event.angleDelta().y()
            font = self.font()
            current_size = font.pointSize()
            if delta > 0 and current_size < 30:
                font.setPointSize(current_size + 1)
            elif delta < 0 and current_size > 6:
                font.setPointSize(current_size - 1)
            self.setFont(font)
        else:
            super().wheelEvent(event)

    def highlight_error_line(self, line_number):
        self.clear_error_highlight()
        selection = QTextEdit.ExtraSelection()
        lineColor = QColor("#FFCCCC")
        selection.format.setBackground(lineColor)
        selection.format.setProperty(QTextFormat.Property.FullWidthSelection, True)
        block = self.document().findBlockByNumber(line_number - 1)
        if block.isValid():
            cursor = self.textCursor()
            cursor.setPosition(block.position())
            selection.cursor = cursor
            self.setExtraSelections(self.extraSelections() + [selection])
            self.setTextCursor(cursor)

    # --- CORRECTED METHOD ---
    def clear_error_highlight(self):
        current_selections = self.extraSelections()
        # FIX IS HERE: .color().name()
        new_selections = [sel for sel in current_selections if sel.format.background().color().name() != "#ffcccc"]
        self.setExtraSelections(new_selections)

# --- REPLACEMENT CLASS: CodeEditorDialog (with Syntax Checking) ---
class CodeEditorDialog(QDialog):
    def __init__(self, module_path: str, class_name: str, method_name: str,
                 all_method_data: dict, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.module_path = module_path
        self.class_name = class_name
        self.method_name = method_name
        self.all_method_data = all_method_data
        self.uses_tabs = False

        self.setWindowTitle(f"Editing: {class_name} in {os.path.basename(module_path)}")
        self.setWindowState(Qt.WindowState.WindowMaximized)

        main_layout = QHBoxLayout(self)
        splitter = QSplitter(Qt.Orientation.Horizontal)

        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        self.nav_tree = QTreeWidget()
        self.nav_tree.setHeaderLabels(["Available Modules"])
        self.nav_tree.itemDoubleClicked.connect(self.insert_method_call)
        left_layout.addWidget(self.nav_tree)

        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        self.code_editor = CodeEditorWithLineNumbers()
        self.code_editor.setFont(QFont("Courier New", 10))
        # Clear error highlight whenever the user types
        self.code_editor.textChanged.connect(self.code_editor.clear_error_highlight)
        self.highlighter = PythonHighlighter(self.code_editor.document())

        self.find_replace_widget = FindReplaceWidget(self.code_editor)
        right_layout.addWidget(self.find_replace_widget)
        right_layout.addWidget(self.code_editor)

        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setSizes([300, 700])
        main_layout.addWidget(splitter)
        
        button_box = QDialogButtonBox()
        # --- ADD NEW SYNTAX CHECK BUTTON ---
        self.check_syntax_button = button_box.addButton("Check Syntax", QDialogButtonBox.ButtonRole.ActionRole)
        self.save_button = button_box.addButton("Save Changes", QDialogButtonBox.ButtonRole.AcceptRole)
        cancel_button = button_box.addButton(QDialogButtonBox.StandardButton.Cancel)
        right_layout.addWidget(button_box)
        
        self.check_syntax_button.clicked.connect(self.check_syntax)
        self.save_button.clicked.connect(self.save_changes)
        cancel_button.clicked.connect(self.reject)
        self.find_replace_widget.next_button.clicked.connect(self.find_next)
        self.find_replace_widget.prev_button.clicked.connect(self.find_previous)
        self.find_replace_widget.replace_button.clicked.connect(self.replace_one)
        self.find_replace_widget.replace_all_button.clicked.connect(self.replace_all)
        
        self.load_and_display_code()

    # --- NEW METHOD: Checks Python syntax ---
    def check_syntax(self):
        """Validates the Python syntax of the code in the editor."""
        code_to_check = self.code_editor.toPlainText()
        
        try:
            ast.parse(code_to_check)
            self.code_editor.clear_error_highlight()
            QMessageBox.information(self, "Syntax OK", "The Python syntax is valid.")
            return True
        except SyntaxError as e:
            # Highlight the line with the error
            self.code_editor.highlight_error_line(e.lineno)
            QMessageBox.critical(self, "Syntax Error", f"Error on line {e.lineno}:\n{e.msg}")
            return False
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred during syntax check:\n{e}")
            return False

    def insert_method_call(self, item: QTreeWidgetItem, column: int):
        method_data = item.data(0, Qt.ItemDataRole.UserRole)
        if not (method_data and isinstance(method_data, tuple)): return
        _, clicked_class_name, clicked_method_name, clicked_module_name, params_dict = method_data
        param_names = [name for name, (_, kind) in params_dict.items() if kind != inspect.Parameter.VAR_KEYWORD and kind != inspect.Parameter.VAR_POSITIONAL]
        param_string = ", ".join(param_names)
        if clicked_class_name == self.class_name:
            text_to_insert = f"self.{clicked_method_name}({param_string})"
        else:
            text_to_insert = f"{clicked_module_name}.{clicked_class_name}(self.context).{clicked_method_name}({param_string})"
        self.code_editor.textCursor().insertText(text_to_insert)

    def populate_nav_tree(self):
        self.nav_tree.clear()
        for module_name, classes in self.all_method_data.items():
            module_item = QTreeWidgetItem(self.nav_tree, [module_name])
            for class_name, methods in classes.items():
                class_item = QTreeWidgetItem(module_item, [class_name])
                for method_info in methods:
                    display_text, _, _, _, _ = method_info
                    method_item = QTreeWidgetItem(class_item, [display_text])
                    method_item.setData(0, Qt.ItemDataRole.UserRole, method_info)
        self.nav_tree.expandToDepth(0)
    
    def detect_indentation(self, code_lines):
        for line in code_lines[:20]:
            if len(line) > 0 and line[0] == '\t':
                self.uses_tabs = True; return
        self.uses_tabs = False
    
    def load_and_display_code(self):
        try:
            with open(self.module_path, 'r', encoding='utf-8') as f: original_code_lines = f.readlines()
            self.detect_indentation(original_code_lines)
            code_with_spaces = [line.replace('\t', ' ' * 4) for line in original_code_lines]
            self.code_editor.setPlainText("".join(code_with_spaces))
            self.populate_nav_tree()
            self.find_and_highlight_method()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not load source file:\n{e}")

    def find_and_highlight_method(self):
        cursor = self.code_editor.textCursor()
        regex = QRegularExpression(rf"^\s*def\s+{self.method_name}\s*\(")
        found_cursor = self.code_editor.document().find(regex, cursor)
        if not found_cursor.isNull():
            self.code_editor.setTextCursor(found_cursor)
            self.code_editor.ensureCursorVisible()
        else:
            QMessageBox.warning(self, "Warning", f"Could not find method '{self.method_name}'.")

    # --- UPDATED METHOD: Now checks syntax before saving ---
    def save_changes(self):
        # First, validate the syntax. If it's not valid, stop the save process.
        if not self.check_syntax():
            return

        reply = QMessageBox.question(self, "Confirm Save", "Syntax is OK. Overwrite the file?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            try:
                code_to_save = self.code_editor.toPlainText()
                if self.uses_tabs: code_to_save = code_to_save.replace(' ' * 4, '\t')
                with open(self.module_path, 'w', encoding='utf-8') as f: f.write(code_to_save)
                QMessageBox.information(self, "Success", "File saved.")
                self.accept()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not save file:\n{e}")

    def find_next(self):
        query = self.find_replace_widget.find_input.text()
        flags = self.find_replace_widget.get_find_flags()
        if not self.code_editor.find(query, flags):
            self.code_editor.moveCursor(QTextCursor.MoveOperation.Start)
            self.code_editor.find(query, flags)

    def find_previous(self):
        query = self.find_replace_widget.find_input.text()
        flags = self.find_replace_widget.get_find_flags()
        flags |= QTextDocument.FindFlag.FindBackward
        if not self.code_editor.find(query, flags):
            self.code_editor.moveCursor(QTextCursor.MoveOperation.End)
            self.code_editor.find(query, flags)
            
    def replace_one(self):
        if self.code_editor.textCursor().hasSelection():
            replace_text = self.find_replace_widget.replace_input.text()
            self.code_editor.textCursor().insertText(replace_text)
        self.find_next()

    def replace_all(self):
        find_text = self.find_replace_widget.find_input.text()
        replace_text = self.find_replace_widget.replace_input.text()
        if not find_text: return
        cursor = self.code_editor.textCursor()
        cursor.setPosition(0)
        self.code_editor.setTextCursor(cursor)
        count = 0
        while self.find_next() and self.code_editor.textCursor().hasSelection():
            self.code_editor.textCursor().insertText(replace_text)
            count += 1
        QMessageBox.information(self, "Finished", f"Replaced {count} occurrence(s).")

class ScheduleTaskDialog(QDialog):
    """A dialog for scheduling bot execution."""
    def __init__(self, bot_name: str, schedule_data: Optional[Dict[str, Any]] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle(f"Schedule Task for '{bot_name}'")
        self.setMinimumWidth(400)

        self.layout = QFormLayout(self)

        self.enable_checkbox = QCheckBox("Enable Schedule")
        self.layout.addRow(self.enable_checkbox)

        self.date_edit = QDateTimeEdit(QDateTime.currentDateTime())
        self.date_edit.setCalendarPopup(True)
        self.layout.addRow("Start Date and Time:", self.date_edit)

        self.repeat_combo = QComboBox()
        self.repeat_combo.addItems(["Do not repeat", "Hourly", "Daily", "Monthly"])
        self.layout.addRow("Repeat:", self.repeat_combo)

        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        self.layout.addRow(self.buttons)

        if schedule_data:
            self.enable_checkbox.setChecked(schedule_data.get("enabled", False))
            start_datetime_str = schedule_data.get("start_datetime")
            if start_datetime_str:
                self.date_edit.setDateTime(QDateTime.fromString(start_datetime_str, Qt.DateFormat.ISODate))
            repeat_str = schedule_data.get("repeat")
            if repeat_str:
                index = self.repeat_combo.findText(repeat_str, Qt.MatchFlag.MatchFixedString)
                if index >= 0:
                    self.repeat_combo.setCurrentIndex(index)

    def get_schedule_data(self) -> Dict[str, Any]:
        """Returns the schedule data from the dialog."""
        return {
            "enabled": self.enable_checkbox.isChecked(),
            "start_datetime": self.date_edit.dateTime().toString(Qt.DateFormat.ISODate),
            "repeat": self.repeat_combo.currentText()
        }
        
class MainWindow(QWidget):
# In main_app.py, replace the __init__ method in the MainWindow class

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Automate Your Task By simple Bot - Designed and Programmed by Phung Tuan Hung")
        # No longer setting geometry here, as we will maximize it.

        # --- 1. Initialize attributes that don't depend on the UI ---
        self.gui_communicator = GuiCommunicator()
        self.base_directory = os.path.dirname(os.path.abspath(__file__))
        self.module_subfolder = "Bot_module"
        self.module_directory = os.path.join(self.base_directory, self.module_subfolder)
        self.click_image_dir = os.path.join(self.base_directory, "Click_image")
        os.makedirs(self.click_image_dir, exist_ok=True)
        self.bot_steps_subfolder = "Bot_steps"
        self.bot_steps_directory = os.path.join(self.base_directory, self.bot_steps_subfolder)
        self.steps_template_subfolder = "Steps_template"
        self.steps_template_directory = os.path.join(self.base_directory, self.steps_template_subfolder)
        self.template_document_directory = os.path.join(self.base_directory, "template_document")
        self.added_steps_data: List[Dict[str, Any]] = []
        self.last_executed_context: Optional[ExecutionContext] = None
        self.global_variables: Dict[str, Any] = {}
        self.loop_id_counter: int = 0
        self.if_id_counter: int = 0
        self.group_id_counter: int = 0
        self.active_param_input_dialog: Optional[ParameterInputDialog] = None
        self.all_parsed_method_data: Dict[str, Dict[str, List[Tuple[str, str, str, str, Dict[str, Any]]]]] = {}
        self.data_to_item_map: Dict[int, QTreeWidgetItem] = {}
        self.minimized_for_execution = False
        self.original_geometry = None
        self.widget_homes = {}
        self.is_bot_running = False # Flag to prevent concurrent executions

        # --- 2. Create all the UI elements ---
        self.init_ui() # This method creates self.log_console

        # --- 3. Now that UI exists, connect signals and start logic that might use it ---
        self.gui_communicator.log_message_signal.connect(self._log_to_console)
        self.gui_communicator.update_module_info_signal.connect(self.update_label_info_from_module)

        # --- SCHEDULER SETUP ---
        self.schedule_timer = QTimer(self)
        self.schedule_timer.timeout.connect(self.check_schedules)
        self.schedule_timer.start(60000) # 60,000 milliseconds = 1 minute
        self._log_to_console("Scheduler started. Will check for due tasks every minute.")
        # --- END SCHEDULER SETUP ---

        # --- 4. Load data and finish the setup ---
        self.load_all_modules_to_tree()
        self.load_saved_steps_to_tree()
        self._update_variables_list_display()

        # --- 5. Maximize the window on startup ---
        self.showMaximized()
    def _get_item_data(self, item: QTreeWidgetItem) -> Optional[Dict[str, Any]]:
        if not item: return None
        data = item.data(0, Qt.ItemDataRole.UserRole)
        return data.value() if isinstance(data, QVariant) else data

    def _get_image_filenames(self) -> List[str]:
        filenames: List[str] = []
        if os.path.exists(self.click_image_dir):
            for filename in os.listdir(self.click_image_dir):
                if filename.lower().endswith(".txt"):
                    filenames.append(os.path.splitext(filename)[0])
        return sorted(filenames)

    def _log_to_console(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_console.append(f"[{timestamp}] {message}")
# In main_app.py, inside the MainWindow class

    # ... (other methods like _log_to_console)

    def select_bot_steps_folder(self):
        """Opens a dialog to allow the user to select a different folder for saved bots."""
        current_dir = self.bot_steps_directory
        new_dir = QFileDialog.getExistingDirectory(self, "Select Bot Steps Folder", current_dir)
        
        if new_dir and new_dir != current_dir:
            self.bot_steps_directory = new_dir
            self.load_saved_steps_to_tree() # Refresh the tree from the new location
            self._log_to_console(f"Changed bot steps folder to: {new_dir}")

    # ... (the rest of the MainWindow methods)
    def _handle_screenshot_request_from_param_dialog(self) -> None:
        if self.active_param_input_dialog:
            self.active_param_input_dialog.hide()
        self._log_to_console("ParameterInputDialog hidden, opening screenshot tool.")
        self.open_screenshot_tool()

    # In main_app.py, REPLACE your entire init_ui method with this one:
    # Add this new method to the MainWindow class
    
    def select_bot_steps_folder(self):
        """Opens a dialog to allow the user to select a different folder for saved bots."""
        current_dir = self.bot_steps_directory
        new_dir = QFileDialog.getExistingDirectory(self, "Select Bot Steps Folder", current_dir)
        
        if new_dir and new_dir != current_dir:
            self.bot_steps_directory = new_dir
            self.load_saved_steps_to_tree() # Refresh the tree from the new location    
    # In main_app.py, REPLACE your entire init_ui method with this one:
    
    # In main_app.py, REPLACE your entire init_ui method with this one:
    
# In main_app.py, inside the MainWindow class

    # REPLACE your existing init_ui method with this one
    def init_ui(self) -> None:
        os.makedirs(self.steps_template_directory, exist_ok=True)
        
        master_layout = QVBoxLayout(self)
        master_layout.setContentsMargins(5, 5, 5, 5) # Add some margin to the main window
        self.stacked_layout = QStackedLayout()
        
        self.full_view_container = QWidget()
        main_layout = QVBoxLayout(self.full_view_container)
        
        # 3. Use a QGridLayout for precise horizontal alignment of top-level items
        bottom_grid_layout = QGridLayout()
        bottom_grid_layout.setContentsMargins(0,0,0,0)

        # --- LEFT PANEL ---
        self.left_panel_widget = QWidget()
        left_panel_layout = QVBoxLayout(self.left_panel_widget)
        left_panel_layout.setContentsMargins(0,0,0,0) # Remove margins to align with grid

        # --- RIGHT PANEL ---
        self.right_panel_widget = QWidget()
        right_panel_layout = QVBoxLayout(self.right_panel_widget)
        right_panel_layout.setContentsMargins(0,0,0,0) # Remove margins to align with grid

        # --- Add Labels for alignment to the top of the grid (Row 0) ---
        saved_bots_label = QLabel("Saved Bots")
        saved_bots_label.setStyleSheet("font-weight: bold; margin-bottom: 2px;")
        execution_flow_label = QLabel("Execution Flow")
        execution_flow_label.setStyleSheet("font-weight: bold; margin-bottom: 2px;")
        
        bottom_grid_layout.addWidget(saved_bots_label, 0, 0)
        bottom_grid_layout.addWidget(execution_flow_label, 0, 1)

        # --- LEFT PANEL CONTENT (goes under its label in the grid) ---
        left_splitter = QSplitter(Qt.Orientation.Vertical)
        # 1. Visualize the splitter handle
        left_splitter.setStyleSheet("""
            QSplitter::handle:vertical {
                background-color: #B0BEC5;
                height: 5px;
                border-top: 1px solid #90A4AE;
                border-bottom: 1px solid #90A4AE;
                margin: 1px 0;
            }
        """)

        # -- Saved Bots Widget (Top of splitter, no more QGroupBox) --
        saved_bots_container = QWidget()
        saved_bots_layout = QVBoxLayout(saved_bots_container)
        saved_bots_layout.setContentsMargins(0,0,0,0)
        self.saved_steps_tree = QTreeWidget()
        self.saved_steps_tree.setHeaderLabels(["Bot Name", "Schedule", "Status"])
        self.saved_steps_tree.header().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.saved_steps_tree.itemDoubleClicked.connect(self.saved_step_tree_item_selected)
        self.saved_steps_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.saved_steps_tree.customContextMenuRequested.connect(self.show_saved_bot_context_menu)
        self.change_bot_folder_button = QPushButton("Change your working folder")
        self.change_bot_folder_button.clicked.connect(self.select_bot_steps_folder)
        saved_bots_layout.addWidget(self.saved_steps_tree)
        saved_bots_layout.addWidget(self.change_bot_folder_button)
        left_splitter.addWidget(saved_bots_container)

        # -- Module Browser container (Bottom of splitter) --
        module_browser_container = QWidget()
        module_browser_layout = QVBoxLayout(module_browser_container)
        module_browser_layout.setContentsMargins(0,5,0,0) # Add a little space above the filter
        
        filter_layout = QHBoxLayout()
        self.filter_label = QLabel("Filter Module:")
        self.module_filter_dropdown = QComboBox()
        self.module_filter_dropdown.addItem("-- Show All Modules --")
        self.module_filter_dropdown.currentIndexChanged.connect(self.filter_module_tree)
        filter_layout.addWidget(self.filter_label)
        filter_layout.addWidget(self.module_filter_dropdown)
        module_browser_layout.addLayout(filter_layout)
        
        self.tree_section_layout = QVBoxLayout()
        self.tree_label = QLabel("Available Modules, Classes, and Methods (Double-click to add):")
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search for methods or templates...")
        self.search_box.textChanged.connect(self.search_module_tree)
        self.tree_section_layout.addWidget(self.tree_label)
        self.tree_section_layout.addWidget(self.search_box)
        self.module_tree = QTreeWidget()
        self.module_tree.setHeaderLabels(["Module/Class/Method"])
        self.module_tree.itemDoubleClicked.connect(self.add_item_to_execution_tree)
        self.module_tree.itemClicked.connect(self.update_selected_method_info)
        self.module_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.module_tree.customContextMenuRequested.connect(self.show_context_menu)
        self.tree_section_layout.addWidget(self.module_tree)
        module_browser_layout.addLayout(self.tree_section_layout)
        left_splitter.addWidget(module_browser_container)
        
        left_splitter.setSizes([200, 400])
        left_panel_layout.addWidget(left_splitter) # This now contains both top and bottom sections

        # Add the rest of the left panel widgets below the splitter
        self.variables_group_box = QGroupBox("Global Variables")
        # ... (rest of the left panel setup remains the same, no changes needed here)
        variables_layout = QVBoxLayout()
        self.variables_list = QListWidget()
        self.variables_list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        variables_layout.addWidget(self.variables_list)
        var_buttons_layout = QHBoxLayout()
        self.add_var_button = QPushButton("Add Variable")
        self.add_var_button.clicked.connect(self.add_variable)
        var_buttons_layout.addWidget(self.add_var_button)
        self.edit_var_button = QPushButton("Edit Variable")
        self.edit_var_button.clicked.connect(self.edit_variable)
        var_buttons_layout.addWidget(self.edit_var_button)
        self.delete_var_button = QPushButton("Delete Variable")
        self.delete_var_button.clicked.connect(self.delete_variable)
        var_buttons_layout.addWidget(self.delete_var_button)
        self.clear_vars_button = QPushButton("Reset All Values to None")
        self.clear_vars_button.clicked.connect(self.reset_all_variable_values)
        var_buttons_layout.addWidget(self.clear_vars_button)
        variables_layout.addLayout(var_buttons_layout)
        self.variables_group_box.setLayout(variables_layout)
        left_panel_layout.addWidget(self.variables_group_box)
        left_panel_layout.addSpacing(10)
        
        execute_buttons_layout = QVBoxLayout()
        execute_row_layout = QHBoxLayout()
        self.execute_all_button = QPushButton("Execute All Steps")
        self.execute_all_button.clicked.connect(self.execute_all_steps)
        execute_row_layout.addWidget(self.execute_all_button)
        self.execute_one_step_button = QPushButton("Execute 1 Step")
        self.execute_one_step_button.clicked.connect(self.execute_one_step)
        self.execute_one_step_button.setEnabled(False)
        execute_row_layout.addWidget(self.execute_one_step_button)
        execute_buttons_layout.addLayout(execute_row_layout)
        block_buttons_layout = QHBoxLayout()
        self.add_loop_button = QPushButton("Add Loop")
        self.add_loop_button.clicked.connect(self.add_loop_block)
        block_buttons_layout.addWidget(self.add_loop_button)
        self.add_conditional_button = QPushButton("Add Conditional Block")
        self.add_conditional_button.clicked.connect(self.add_conditional_block)
        block_buttons_layout.addWidget(self.add_conditional_button)
        execute_buttons_layout.addLayout(block_buttons_layout)
        left_panel_layout.addLayout(execute_buttons_layout)
        
        button_row_layout_1 = QHBoxLayout()
        self.save_steps_button = QPushButton("Save Bot Steps")
        self.save_steps_button.clicked.connect(self.save_bot_steps_dialog)
        button_row_layout_1.addWidget(self.save_steps_button)
        self.group_steps_button = QPushButton("Group Selected")
        self.group_steps_button.clicked.connect(self.group_selected_steps)
        button_row_layout_1.addWidget(self.group_steps_button)
        self.clear_selected_button = QPushButton("Clear Selected Steps")
        self.clear_selected_button.clicked.connect(self.clear_selected_steps)
        button_row_layout_1.addWidget(self.clear_selected_button)
        self.remove_all_steps_button = QPushButton("Remove All Steps")
        self.remove_all_steps_button.clicked.connect(self.clear_all_steps)
        button_row_layout_1.addWidget(self.remove_all_steps_button)
        left_panel_layout.addLayout(button_row_layout_1)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.hide()
        left_panel_layout.addWidget(self.progress_bar)
        left_panel_layout.addSpacing(10)
        
        utility_buttons_layout = QHBoxLayout()
        self.utility_buttons_layout_2 = QHBoxLayout()
        self.always_on_top_button = QPushButton("Always On Top: Off")
        self.always_on_top_button.setCheckable(True)
        self.always_on_top_button.clicked.connect(self.toggle_always_on_top)
        utility_buttons_layout.addWidget(self.always_on_top_button)
        self.open_screenshot_tool_button = QPushButton("Open Screenshot Tool")
        self.open_screenshot_tool_button.clicked.connect(self.open_screenshot_tool)
        utility_buttons_layout.addWidget(self.open_screenshot_tool_button)
        self.toggle_log_checkbox = QCheckBox("Show Execution Log")
        self.toggle_log_checkbox.setChecked(False)
        self.utility_buttons_layout_2.addWidget(self.toggle_log_checkbox)
        self.exit_button = QPushButton("Exit GUI")
        self.exit_button.clicked.connect(QApplication.instance().quit)
        self.utility_buttons_layout_2.addWidget(self.exit_button)
        left_panel_layout.addLayout(utility_buttons_layout)
        left_panel_layout.addLayout(self.utility_buttons_layout_2)

        # --- RIGHT PANEL CONTENT (goes under its label in the grid) ---
        self.right_splitter = QSplitter(Qt.Orientation.Horizontal)
        self.execution_tree_widget = QWidget()
        self.execution_tree_layout = QVBoxLayout(self.execution_tree_widget)
        self.execution_tree_layout.setContentsMargins(0,0,0,0)
        
        # 2. Reduce space around the website link
        self.website_label = QLabel('<a href="http://www.AutomateTask.Click" style="color: blue; text-decoration: none; font-size: 14pt;">www.AutomateTask.Click</a>')
        self.website_label.setOpenExternalLinks(True)
        self.website_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.website_label.setContentsMargins(0, 2, 0, 2) # top, left, bottom, right margins
        
        self.execution_tree = GroupedTreeWidget(self)
        self.execution_tree.setHeaderHidden(True) # Hiding the header to align content
        self.execution_tree.setDragDropMode(QTreeWidget.DragDropMode.NoDragDrop)
        self.execution_tree.setSelectionMode(QTreeWidget.SelectionMode.ExtendedSelection)
        
        self.execution_tree_layout.addWidget(self.website_label)
        self.execution_tree_layout.addWidget(self.execution_tree)

        self.info_labels_layout = QHBoxLayout()
        self.label_info1 = QLabel("Module Info: None")
        # ... (rest of the info labels setup is fine)
        self.label_info1.setStyleSheet("font-style: italic; color: gray;")
        self.label_info2 = QLabel("Image Preview")
        self.label_info2.setStyleSheet("font-style: italic; color: gray;")
        self.label_info2.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label_info3 = QLabel("Image Name")
        self.label_info3.setStyleSheet("font-style: italic; color: blue;")
        self.info_labels_layout.addWidget(self.label_info1, 1, Qt.AlignmentFlag.AlignLeft)
        self.info_labels_layout.addWidget(self.label_info2, 1, Qt.AlignmentFlag.AlignCenter)
        self.info_labels_layout.addWidget(self.label_info3, 1, Qt.AlignmentFlag.AlignRight)
        self.execution_tree_layout.addLayout(self.info_labels_layout)
    
        self.log_widget = QWidget()
        # ... (log widget setup is fine)
        log_layout = QVBoxLayout(self.log_widget)
        log_group_box = QGroupBox("Execution Log")
        log_group_layout = QVBoxLayout()
        self.log_console = QTextEdit()
        self.log_console.setReadOnly(True)
        self.log_console.setStyleSheet("background-color: #2E2E2E; color: #E0E0E0; font-family: 'Consolas', 'Courier New', monospace;")
        log_group_layout.addWidget(self.log_console)
        self.clear_log_button = QPushButton("Clear Log")
        self.clear_log_button.clicked.connect(self.log_console.clear)
        log_group_layout.addWidget(self.clear_log_button)
        log_group_box.setLayout(log_group_layout)
        log_layout.addWidget(log_group_box)
        
        self.toggle_log_checkbox.toggled.connect(self.log_widget.setVisible)
        self.log_widget.setVisible(False)
        
        self.right_splitter.addWidget(self.execution_tree_widget)
        self.right_splitter.addWidget(self.log_widget)
        self.right_splitter.setSizes([600, 400])
        
        # --- Add the main content widgets to the grid (Row 1) ---
        self.left_panel_widget.setLayout(left_panel_layout)
        self.right_panel_widget.setLayout(right_panel_layout)
        right_panel_layout.addWidget(self.right_splitter)

        bottom_grid_layout.addWidget(self.left_panel_widget, 1, 0)
        bottom_grid_layout.addWidget(self.right_panel_widget, 1, 1)

        # Set column stretch factors
        bottom_grid_layout.setColumnStretch(0, 1)
        bottom_grid_layout.setColumnStretch(1, 2)

        main_layout.addLayout(bottom_grid_layout)
        
        # --- MINI VIEW (no changes needed) ---
        self.mini_view_container = QWidget()
        # ...
        self.mini_view_layout = QVBoxLayout(self.mini_view_container)
        self.mini_view_layout.setContentsMargins(5,5,5,5)
    
        self.stacked_layout.addWidget(self.full_view_container)
        self.stacked_layout.addWidget(self.mini_view_container)
        master_layout.addLayout(self.stacked_layout)
        
        self.execution_tree.itemSelectionChanged.connect(self._toggle_execute_one_step_button)
    
        self.widget_homes = {
            self.execution_tree: (self.execution_tree_layout, 1),
            self.label_info2: (self.info_labels_layout, 1),
            self.progress_bar: (left_panel_layout, 4),
            self.exit_button: (self.utility_buttons_layout_2, 1)
        }
    def show_saved_bot_context_menu(self, position: QPoint):
        item = self.saved_steps_tree.itemAt(position)
        if not item or item.text(0) == "No saved bots found.":
            return

        bot_name = item.text(0)
        context_menu = QMenu(self)
        open_action = context_menu.addAction("Open Bot")
        schedule_action = context_menu.addAction("Schedule Task")
        delete_action = context_menu.addAction("Delete Bot")

        action = context_menu.exec(self.saved_steps_tree.mapToGlobal(position))

        if action == open_action:
            self.open_saved_bot(bot_name)
        elif action == schedule_action:
            self.schedule_bot(bot_name)
        elif action == delete_action:
            self.delete_saved_bot(bot_name)

    def open_saved_bot(self, bot_name: str):
        """Loads the selected bot into the Execution Flow."""
        file_path = os.path.join(self.bot_steps_directory, f"{bot_name}.csv")
        if os.path.exists(file_path):
            self.load_steps_from_file(file_path)
        else:
            QMessageBox.warning(self, "File Not Found", f"The bot file '{bot_name}.csv' was not found.")
            self.load_saved_steps_to_tree()

    def schedule_bot(self, bot_name: str):
        """Opens the scheduling dialog for the selected bot."""
        schedule_data = self.schedules.get(bot_name)
        dialog = ScheduleTaskDialog(bot_name, schedule_data, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.schedules[bot_name] = dialog.get_schedule_data()
            self.save_schedules()
            self.load_saved_steps_to_tree()

    def delete_saved_bot(self, bot_name: str):
        """Deletes the selected bot and its schedule."""
        reply = QMessageBox.question(self, "Confirm Delete",
                                     f"Are you sure you want to permanently delete the bot '{bot_name}' and its schedule?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            try:
                bot_path = os.path.join(self.bot_steps_directory, f"{bot_name}.csv")
                if os.path.exists(bot_path):
                    os.remove(bot_path)

                if bot_name in self.schedules:
                    del self.schedules[bot_name]
                    self.save_schedules()

                QMessageBox.information(self, "Success", f"Bot '{bot_name}' has been deleted.")
                self.load_saved_steps_to_tree()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"An error occurred while deleting the bot:\n{e}")

    def save_schedules(self):
        """Saves the current schedules to a JSON file."""
        os.makedirs(self.schedules_directory, exist_ok=True)
        schedule_file_path = os.path.join(self.schedules_directory, "schedules.json")
        try:
            with open(schedule_file_path, 'w', encoding='utf-8') as f:
                json.dump(self.schedules, f, indent=4)
        except Exception as e:
            self._log_to_console(f"Error saving schedules: {e}")


    def load_saved_steps_to_tree(self) -> None:
        """Loads saved bot step files and their schedules into the QTreeWidget."""
        self.saved_steps_tree.clear()
        self.load_schedules()
        try:
            os.makedirs(self.bot_steps_directory, exist_ok=True)
            step_files = sorted([f for f in os.listdir(self.bot_steps_directory) if f.endswith(".csv")], reverse=True)
            for file_name in step_files:
                bot_name = os.path.splitext(file_name)[0]
                schedule_info = self.schedules.get(bot_name)
                schedule_str = "Not Set"
                status_str = "Idle"
                if schedule_info:
                    schedule_str = f"{schedule_info.get('repeat', 'Once')} at {QDateTime.fromString(schedule_info.get('start_datetime'), Qt.DateFormat.ISODate).toString('yyyy-MM-dd hh:mm')}"
                    status_str = "Scheduled" if schedule_info.get("enabled") else "Disabled"

                tree_item = QTreeWidgetItem(self.saved_steps_tree, [bot_name, schedule_str, status_str])
            if not step_files:
                self.saved_steps_tree.addTopLevelItem(QTreeWidgetItem(["No saved bots found."]))
        except Exception as e:
            QMessageBox.critical(self, "Error Loading Saved Bots", f"Could not load bot files: {e}")

    def load_schedules(self):
        """Loads schedules from the JSON file."""
        schedule_file_path = os.path.join(self.schedules_directory, "schedules.json")
        if os.path.exists(schedule_file_path):
            try:
                with open(schedule_file_path, 'r', encoding='utf-8') as f:
                    self.schedules = json.load(f)
            except (json.JSONDecodeError, Exception) as e:
                self._log_to_console(f"Could not load schedules file: {e}")
                self.schedules = {}
        else:
            self.schedules = {}
    def _toggle_minimized_view(self, minimize: bool):
        if minimize:
            self.minimized_for_execution = True
            self.original_geometry = self.geometry()

            self.mini_view_layout.addWidget(self.execution_tree)
            self.mini_view_layout.addWidget(self.label_info2)
            self.mini_view_layout.addWidget(self.progress_bar)
            self.mini_view_layout.addWidget(self.exit_button)
            
            self.progress_bar.show()

            self.stacked_layout.setCurrentWidget(self.mini_view_container)

            new_width, new_height = 350, 400
            screen_geometry = QApplication.primaryScreen().geometry()
            self.move(screen_geometry.width() - new_width, 0)
            self.setFixedSize(new_width, new_height)

        else:
            if self.original_geometry:
                self.setFixedSize(QSize(16777215, 16777215))
                self.setGeometry(self.original_geometry)

            for widget, (layout, index) in self.widget_homes.items():
                if isinstance(layout, QBoxLayout):
                    layout.insertWidget(index, widget)
                else:
                    layout.addWidget(widget)

            self.progress_bar.hide()
            
            self.stacked_layout.setCurrentWidget(self.full_view_container)
            self.minimized_for_execution = False


    def _toggle_execute_one_step_button(self) -> None:
        self.execute_one_step_button.setEnabled(len(self.execution_tree.selectedItems()) > 0)

    def _rebuild_added_steps_data_from_tree(self):
        new_flat_data: List[Dict[str, Any]] = []
        self._flatten_tree_recursive(self.execution_tree.invisibleRootItem(), new_flat_data)
        self.added_steps_data = new_flat_data
        self._update_original_listbox_row_indices()
        selected_item_data = self._get_item_data(self.execution_tree.currentItem())
        self._rebuild_execution_tree(item_to_focus_data=selected_item_data)

    def _flatten_tree_recursive(self, parent_item: QTreeWidgetItem, flat_list: List[Dict[str, Any]]):
        for i in range(parent_item.childCount()):
            child_item = parent_item.child(i)
            step_data = self._get_item_data(child_item)
            if step_data and isinstance(step_data, dict):
                flat_list.append(step_data)
            self._flatten_tree_recursive(child_item, flat_list)

    def _update_original_listbox_row_indices(self):
        for i, step_data in enumerate(self.added_steps_data):
            step_data["original_listbox_row_index"] = i
            if step_data["type"] == "step" and "parameters_config" in step_data and step_data["parameters_config"] is not None:
                step_data["parameters_config"]["original_listbox_row_index"] = i

# Add these two new methods inside the MainWindow class

# In the MainWindow class, replace your existing method with this one

    def _get_expansion_state(self) -> set:
        """Recursively finds all expanded block items and returns their unique IDs."""
        expanded_ids = set()
        
        def traverse(parent_item):
            for i in range(parent_item.childCount()):
                child_item = parent_item.child(i)
                # --- THIS IS THE CORRECTED LINE ---
                if child_item.isExpanded():
                    item_data = self._get_item_data(child_item)
                    if item_data:
                        item_id = item_data.get("group_id") or item_data.get("loop_id") or item_data.get("if_id")
                        if item_id:
                            expanded_ids.add(item_id)
                # Recurse into children
                if child_item.childCount() > 0:
                    traverse(child_item)
    
        traverse(self.execution_tree.invisibleRootItem())
        return expanded_ids

    def _restore_expansion_state(self, expanded_ids: set):
        """Recursively traverses the tree and expands any items whose ID is in the provided set."""
        def traverse(parent_item):
            for i in range(parent_item.childCount()):
                child_item = parent_item.child(i)
                item_data = self._get_item_data(child_item)
                if item_data:
                    item_id = item_data.get("group_id") or item_data.get("loop_id") or item_data.get("if_id")
                    if item_id and item_id in expanded_ids:
                        self.execution_tree.expandItem(child_item)
                
                if child_item.childCount() > 0:
                    traverse(child_item)
        
        traverse(self.execution_tree.invisibleRootItem())
    
    # In the MainWindow class, REPLACE your existing method with this one
    
    def _rebuild_execution_tree(self, item_to_focus_data: Optional[Dict[str, Any]] = None) -> None:
        """
        Rebuilds the entire execution tree from the flat `added_steps_data` list.
        - Handles nesting of blocks correctly.
        - Selects and scrolls to a specified item.
        - PRESERVES the expansion state of blocks.
        - Populates the data_to_item_map for visual guides.
        """
        expanded_state = self._get_expansion_state()
    
        self.execution_tree.clear()
        # --- ADD THIS LINE to clear the map ---
        self.data_to_item_map.clear()
        
        current_parent_stack: List[QTreeWidgetItem] = [self.execution_tree.invisibleRootItem()]
        item_to_focus: Optional[QTreeWidgetItem] = None
    
        for i, step_data_dict in enumerate(self.added_steps_data):
            step_data_dict["original_listbox_row_index"] = i
            step_type = step_data_dict.get("type")
    
            # Stack Popping Logic
            if step_type == "group_end":
                if len(current_parent_stack) > 1:
                    last_parent_data = self._get_item_data(current_parent_stack[-1])
                    if last_parent_data and last_parent_data.get("type") == "group_start" and last_parent_data.get("group_id") == step_data_dict.get("group_id"):
                        current_parent_stack.pop()
            elif step_type == "loop_end":
                if len(current_parent_stack) > 1:
                    last_parent_data = self._get_item_data(current_parent_stack[-1])
                    if last_parent_data and last_parent_data.get("type") == "loop_start" and last_parent_data.get("loop_id") == step_data_dict.get("loop_id"):
                        current_parent_stack.pop()
            elif step_type == "IF_END":
                if len(current_parent_stack) > 1:
                    last_parent_data = self._get_item_data(current_parent_stack[-1])
                    if last_parent_data and last_parent_data.get("type") == "ELSE" and last_parent_data.get("if_id") == step_data_dict.get("if_id"):
                        current_parent_stack.pop()
                        if len(current_parent_stack) > 1:
                             if_start_parent_data = self._get_item_data(current_parent_stack[-1])
                             if if_start_parent_data and if_start_parent_data.get("type") == "IF_START" and if_start_parent_data.get("if_id") == step_data_dict.get("if_id"):
                                 current_parent_stack.pop()
                    elif last_parent_data and last_parent_data.get("type") == "IF_START" and last_parent_data.get("if_id") == step_data_dict.get("if_id"):
                        current_parent_stack.pop()
    
            # Item Creation
            parent_for_current_item = current_parent_stack[-1]
            tree_item = QTreeWidgetItem(parent_for_current_item)
            tree_item.setData(0, Qt.ItemDataRole.UserRole, QVariant(step_data_dict))
            
            # --- ADD THIS LINE to populate the map ---
            self.data_to_item_map[i] = tree_item
    
            card = ExecutionStepCard(step_data_dict, i + 1)
            card.edit_requested.connect(self._handle_edit_request)
            card.delete_requested.connect(self._handle_delete_request)
            card.move_up_requested.connect(self._handle_move_up_request)
            card.move_down_requested.connect(self._handle_move_down_request)
            card.save_as_template_requested.connect(self._handle_save_as_template_request)
            card.execute_this_requested.connect(self._handle_execute_this_request)
            tree_item.setSizeHint(0, card.sizeHint())
            self.execution_tree.setItemWidget(tree_item, 0, card)
    
            if item_to_focus_data and step_data_dict == item_to_focus_data:
                item_to_focus = tree_item
    
            # Stack Pushing Logic
            if step_type in ["loop_start", "IF_START", "ELSE", "group_start"]:
                current_parent_stack.append(tree_item)
    
        if item_to_focus:
            self.execution_tree.setCurrentItem(item_to_focus)
            self.execution_tree.scrollToItem(item_to_focus, QtWidgets.QAbstractItemView.ScrollHint.PositionAtCenter)
    
        self.update_status_column_for_all_items()
        self._restore_expansion_state(expanded_state)
    

    def _handle_edit_request(self, step_data_to_edit: Dict[str, Any]):
        item_to_edit = self._find_qtreewidget_item(step_data_to_edit)
        if item_to_edit:
            self.edit_step_in_execution_tree(item_to_edit, 0)
        else:
            QMessageBox.critical(self, "Error", "Could not find the step to edit.")

    def _handle_delete_request(self, step_to_delete: Dict[str, Any]):
        step_type = step_to_delete.get("type")
        start_index = -1
        try:
            start_index = self.added_steps_data.index(step_to_delete)
        except ValueError:
            QMessageBox.critical(self, "Error", "Could not find step to delete in data list.")
            return

        if step_type in ["loop_start", "IF_START", "group_start"]:
            start_idx, end_idx = self._find_block_indices(start_index)
            if start_idx != end_idx: 
                items_to_remove_data = self.added_steps_data[start_idx : end_idx + 1]
            else: 
                items_to_remove_data = [step_to_delete]
        else:
            items_to_remove_data = [step_to_delete]

        if QMessageBox.question(self, "Confirm Delete", f"Are you sure you want to delete {len(items_to_remove_data)} selected step(s)?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            self.added_steps_data = [s for s in self.added_steps_data if s not in items_to_remove_data]
            self._rebuild_execution_tree()
            
    def _find_block_indices(self, start_index: int) -> Tuple[int, int]:
        start_step = self.added_steps_data[start_index]
        step_type = start_step.get("type")

        if step_type not in ["loop_start", "IF_START", "group_start"]:
            return start_index, start_index

        if step_type == "group_start":
            block_id = start_step.get("group_id")
            end_type = "group_end"
        elif step_type == "loop_start":
            block_id = start_step.get("loop_id")
            end_type = "loop_end"
        else:
            block_id = start_step.get("if_id")
            end_type = "IF_END"

        nesting_level = 0
        for i in range(start_index + 1, len(self.added_steps_data)):
            current_step = self.added_steps_data[i]
            current_type = current_step.get("type")

            if current_type in ["loop_start", "IF_START", "group_start"]:
                nesting_level += 1
            elif current_type == end_type and (current_step.get("group_id") or current_step.get("loop_id") or current_step.get("if_id")) == block_id:
                if nesting_level == 0:
                    return start_index, i
                else:
                    nesting_level -= 1
            elif current_type in ["loop_end", "IF_END", "group_end"]:
                nesting_level -= 1
        return start_index, start_index

    def _handle_move_up_request(self, step_data: Dict[str, Any]):
        try:
            current_pos = self.added_steps_data.index(step_data)
        except ValueError:
            return
        if current_pos == 0:
            return
        start_idx, end_idx = self._find_block_indices(current_pos)
        block_to_move = self.added_steps_data[start_idx : end_idx + 1]
        
        prev_start_idx, prev_end_idx = self._find_block_indices(start_idx - 1) if (start_idx -1) >= 0 else (start_idx-1, start_idx-1)
        if prev_start_idx == -1:
            return
        
        part_before = self.added_steps_data[0:prev_start_idx]
        part_middle = self.added_steps_data[prev_start_idx : start_idx]
        part_after = self.added_steps_data[end_idx+1:]
        
        self.added_steps_data = part_before + block_to_move + part_middle + part_after
        self._rebuild_execution_tree(item_to_focus_data=block_to_move[0])

    def _handle_move_down_request(self, step_data: Dict[str, Any]):
        try:
            current_pos = self.added_steps_data.index(step_data)
        except ValueError:
            return
        start_idx, end_idx = self._find_block_indices(current_pos)
        if end_idx >= len(self.added_steps_data) - 1:
            return
        
        block_to_move = self.added_steps_data[start_idx : end_idx + 1]
        
        next_start_idx, next_end_idx = self._find_block_indices(end_idx+1)

        part_before = self.added_steps_data[0:start_idx]
        part_middle = self.added_steps_data[end_idx + 1 : next_end_idx + 1]
        part_after = self.added_steps_data[next_end_idx+1:]
        
        self.added_steps_data = part_before + part_middle + block_to_move + part_after
        self._rebuild_execution_tree(item_to_focus_data=block_to_move[0])

    # --- NEW HELPER METHOD ---
    def _get_template_names(self) -> List[str]:
        """Returns a sorted list of template names without extensions."""
        os.makedirs(self.steps_template_directory, exist_ok=True)
        return sorted([os.path.splitext(f)[0] for f in os.listdir(self.steps_template_directory) if f.endswith(".json")])

    # --- MODIFIED METHOD ---
    def _handle_save_as_template_request(self, step_data_to_save: Dict[str, Any]):
        try:
            start_index = self.added_steps_data.index(step_data_to_save)
        except ValueError:
            QMessageBox.critical(self, "Error", "Could not find the step to save as a template.")
            return

        start_idx, end_idx = self._find_block_indices(start_index)
        if start_idx == end_idx:
            QMessageBox.warning(self, "Save Error", "Cannot save a single step as a block template. The block is incomplete.")
            return
            
        block_to_save = self.added_steps_data[start_idx : end_idx + 1]

        existing_templates = self._get_template_names()
        dialog = SaveTemplateDialog(existing_templates, self)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            template_name = dialog.get_template_name()
            if not template_name:
                return # Dialog already showed an error message

            # Check for overwrite
            if template_name in existing_templates:
                reply = QMessageBox.question(self, "Confirm Overwrite", 
                                             f"A template named '{template_name}' already exists. Do you want to overwrite it?",
                                             QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if reply == QMessageBox.StandardButton.No:
                    return

            file_path = os.path.join(self.steps_template_directory, f"{template_name}.json")
            
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(block_to_save, f, indent=4)
                QMessageBox.information(self, "Success", f"Template '{template_name}' saved successfully.")
                self.load_all_modules_to_tree() # Refresh the tree to show the new/updated template
            except Exception as e:
                QMessageBox.critical(self, "Save Error", f"Failed to save the template:\n{e}")

    def _handle_execute_this_request(self, step_data: Dict[str, Any]):
        """Executes a single step requested from an ExecutionStepCard button."""
        if not (step_data and isinstance(step_data, dict)):
            return
        try:
            current_row = self.added_steps_data.index(step_data)
        except ValueError:
            QMessageBox.critical(self, "Error", "Could not find the selected step in the internal data model.")
            return

        self.set_ui_enabled_state(False)
        self.progress_bar.setValue(0)
        self.progress_bar.show()
        self.update_status_column_for_all_items()
        
        self.worker = ExecutionWorker(
            self.added_steps_data, 
            self.module_directory, 
            self.gui_communicator, 
            self.global_variables, 
            single_step_mode=True, 
            selected_start_index=current_row
        )
        self._connect_worker_signals()
        self.worker.start()

    def update_status_column_for_all_items(self):
        self._clear_status_recursive(self.execution_tree.invisibleRootItem())

    def _clear_status_recursive(self, parent_item: QTreeWidgetItem):
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            card = self.execution_tree.itemWidget(child, 0)
            if card:
                card.set_status("#D3D3D3")
                card.clear_result()
            self._clear_status_recursive(child)

    def toggle_always_on_top(self) -> None:
        if self.always_on_top_button.isChecked():
            self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
            self.always_on_top_button.setText("Always On Top: On")
        else:
            self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, False)
            self.always_on_top_button.setText("Always On Top: Off")
        self.show()

    def open_screenshot_tool(self, initial_image_name: str = "") -> None:
        self.hide()
        current_image_name_for_tool = self.label_info3.text() if self.label_info3.text() != "Image Name" else ""
        self.screenshot_window = SecondWindow(current_image_name_for_tool, parent=self)
        self.screenshot_window.screenshot_saved.connect(self._handle_screenshot_tool_closed)
        self.screenshot_window.show()

    def _handle_screenshot_tool_closed(self, saved_filename: str) -> None:
        self.show()
        self._log_to_console(f"Screenshot tool closed. Saved filename: '{saved_filename}'")
        if self.active_param_input_dialog and not self.active_param_input_dialog.isHidden():
            new_filenames = self._get_image_filenames()
            self.active_param_input_dialog.update_image_filenames.emit(new_filenames, saved_filename)
        elif self.active_param_input_dialog:
             self.active_param_input_dialog.show()
             new_filenames = self._get_image_filenames()
             self.active_param_input_dialog.update_image_filenames.emit(new_filenames, saved_filename)
    
    # --- MODIFIED METHOD ---
    def load_all_modules_to_tree(self) -> None:
        self.module_tree.clear()
        self.module_filter_dropdown.clear()
        self.module_filter_dropdown.addItem("-- Show All Modules --")
        self.label_info1.setText("Module Info: Loading modules...")
        self.all_parsed_method_data.clear()
        dummy_context = ExecutionContext()
        dummy_communicator = GuiCommunicator()
        dummy_context.set_gui_communicator(dummy_communicator)
        try:
            if self.module_directory not in sys.path:
                sys.path.insert(0, self.module_directory)
            if not os.path.exists(self.module_directory):
                QMessageBox.warning(self, "Module Folder Not Found", f"The '{self.module_subfolder}' folder was not found at: {self.module_directory}\nPlease create it and place your Python modules inside.")
                self.label_info1.setText("Error: Bot_module folder not found.")
                return
            module_files: List[str] = [filename[:-3] for filename in os.listdir(self.module_directory) if filename.endswith(".py") and filename != "__init__.py"]
            module_files.sort()
            if not module_files:
                self.label_info1.setText(f"No Python modules found in '{self.module_subfolder}' folder.")
                
            for module_name in module_files:
                self.all_parsed_method_data[module_name] = {}
                self.module_filter_dropdown.addItem(module_name)
                try:
                    module = importlib.import_module(module_name)
                    importlib.reload(module)
                    module_item = QTreeWidgetItem(self.module_tree)
                    module_item.setText(0, module_name)
                    module_item.setFlags(module_item.flags() & ~Qt.ItemFlag.ItemIsSelectable)
                    module_item.setToolTip(0, f"Module: {module_name}")
            
                    for name, obj in inspect.getmembers(module):
                        if inspect.isclass(obj) and obj.__module__ == module.__name__:
                            class_name = name
                            self.all_parsed_method_data[module_name][class_name] = []
                            class_item = QTreeWidgetItem(module_item)
                            class_item.setText(0, class_name)
                            class_item.setFlags(class_item.flags() & ~Qt.ItemFlag.ItemIsSelectable)
                            class_item.setToolTip(0, f"Class: {class_name} in {module_name}")
                            init_kwargs_for_inspection: Dict[str, Any] = {}
                            try:
                                init_signature = inspect.signature(obj.__init__)
                                if 'context' in init_signature.parameters:
                                    init_kwargs_for_inspection['context'] = dummy_context
                                temp_instance = obj(**init_kwargs_for_inspection)
                            except Exception as e:
                                self._log_to_console(f"Warning: Could not instantiate class '{class_name}' for inspection: {e}")
                                class_item.addChild(QTreeWidgetItem(["(Init error, cannot inspect methods)"]))
                                continue
                            for method_name, method_obj in inspect.getmembers(temp_instance):
                                if (not method_name.startswith('_') and callable(method_obj) and not inspect.isclass(method_obj) and not isinstance(method_obj, staticmethod) and not isinstance(method_obj, classmethod)):
                                    func_obj = method_obj.__func__ if inspect.ismethod(method_obj) else method_obj
                                    try:
                                        sig = inspect.signature(func_obj)
                                        params_for_dialog: Dict[str, Tuple[Any, Any]] = {p.name: (p.default, p.kind) for p in sig.parameters.values() if p.name not in ['self', 'context']}
                                        display_text = method_name
                                        
                                        method_info_tuple = (display_text, class_name, method_name, module_name, params_for_dialog)
                                        self.all_parsed_method_data[module_name][class_name].append(method_info_tuple)
                                        method_item = QTreeWidgetItem(class_item)
                                        method_item.setText(0, display_text)
                                        method_item.setToolTip(0, f"Method: {class_name}.{method_name}")
                                        method_item.setData(0, Qt.ItemDataRole.UserRole, QVariant(method_info_tuple))
                                    except Exception as e:
                                        self._log_to_console(f"Warning: Error inspecting '{class_name}.{method_name}': {e}")
                except Exception as e:
                    self._log_to_console(f"Error loading module '{module_name}': {e}")

            # --- NEW: Load templates into the tree ---
            template_root_item = QTreeWidgetItem(self.module_tree)
            template_root_item.setText(0, "Bot Templates")
            template_root_item.setFlags(template_root_item.flags() & ~Qt.ItemFlag.ItemIsSelectable)
            template_root_item.setToolTip(0, "Saved step templates")
            try:
                for template_name in self._get_template_names():
                    template_item = QTreeWidgetItem(template_root_item)
                    template_item.setText(0, template_name)
                    template_item.setToolTip(0, f"Template: {template_name}")
                    template_item.setData(0, Qt.ItemDataRole.UserRole, QVariant({'type': 'template', 'name': template_name}))
            except Exception as e:
                self._log_to_console(f"Error loading templates into tree: {e}")
                error_item = QTreeWidgetItem(template_root_item)
                error_item.setText(0, "(Error loading templates)")

            # --- MODIFIED: Collapse all items by default ---
            self.module_tree.collapseAll()
            self.label_info1.setText("Module Info: All modules loaded.")

        finally:
            if self.module_directory in sys.path:
                sys.path.remove(self.module_directory)
        self.module_tree.itemClicked.connect(self.update_selected_method_info)

    def filter_module_tree(self, index: int) -> None:
        selected_module_name = self.module_filter_dropdown.currentText()
        root = self.module_tree.invisibleRootItem()
        for i in range(root.childCount()):
            module_item = root.child(i)
            # Also hide/show the templates root item based on filter
            is_template_item = module_item.text(0) == "Bot Templates"
            if is_template_item:
                module_item.setHidden(not (selected_module_name == "-- Show All Modules --"))
            else:
                module_item.setHidden(not (selected_module_name == "-- Show All Modules --" or module_item.text(0) == selected_module_name))
        
        # --- MODIFIED: Collapse all instead of expanding all ---
        self.module_tree.collapseAll()

    # --- NEW METHOD: search_module_tree ---
    def search_module_tree(self, text: str) -> None:
        """Filters the module tree based on the search text."""
        search_text = text.lower()
        root = self.module_tree.invisibleRootItem()

        def filter_recursive(item: QTreeWidgetItem) -> bool:
            # Check if any child matches
            child_matches = False
            for i in range(item.childCount()):
                if filter_recursive(item.child(i)):
                    child_matches = True
            
            # Check if the item itself matches
            item_text = item.text(0).lower()
            item_matches = search_text in item_text
            
            # The item should be visible if it matches OR if any of its children match
            should_be_visible = item_matches or child_matches
            item.setHidden(not should_be_visible)

            # Expand parents of matching items, but not the matching item itself unless it has children
            if should_be_visible and not item_matches and child_matches:
                 item.setExpanded(True)
            elif not search_text: # Collapse all when search is cleared
                item.setExpanded(False)

            return should_be_visible

        # Iterate over top-level items (modules and templates category)
        for i in range(root.childCount()):
            filter_recursive(root.child(i))

    def saved_step_tree_item_selected(self, item: QTreeWidgetItem, column: int):
        """Loads a bot's steps when its item is double-clicked in the tree."""
        bot_name = item.text(0)
        if bot_name == "No saved bots found.":
            return
    
        file_path = os.path.join(self.bot_steps_directory, f"{bot_name}.csv")
        if os.path.exists(file_path):
            self.load_steps_from_file(file_path)
        else:
            QMessageBox.warning(self, "File Not Found", f"The selected bot file was not found:\n{file_path}")
            self.load_saved_steps_to_tree() # Refresh the tree if a file is missing


    # --- NEW HELPER: _extract_variables_from_steps ---
    def _extract_variables_from_steps(self, steps: List[Dict[str, Any]]) -> set:
        """Recursively finds all variable names used in a list of steps."""
        found_variables = set()

        def search_dict(d: Dict[str, Any]):
            # Check for variable assignment
            if "assign_to_variable_name" in d and d["assign_to_variable_name"]:
                found_variables.add(d["assign_to_variable_name"])
            if "assign_iteration_to_variable" in d and d.get("assign_iteration_to_variable"):
                 found_variables.add(d["assign_iteration_to_variable"])
            
            # Check for variable usage in standard structures
            if d.get("type") == "variable" and "value" in d:
                found_variables.add(d["value"])

            # Recurse into nested structures
            for key, value in d.items():
                if isinstance(value, dict):
                    search_dict(value)
                elif isinstance(value, list):
                    search_list(value)

        def search_list(lst: List[Any]):
            for item in lst:
                if isinstance(item, dict):
                    search_dict(item)
                elif isinstance(item, list):
                    search_list(item)
        
        search_list(steps)
        return found_variables
    
    # --- NEW HELPER: _apply_variable_mapping ---
    def _apply_variable_mapping(self, steps: List[Dict[str, Any]], mapping: Dict[str, str]) -> List[Dict[str, Any]]:
        """Recursively replaces variable names in steps based on the provided mapping."""
        
        # Create a deep copy to avoid modifying the original template data
        steps_copy = json.loads(json.dumps(steps))

        def replace_in_dict(d: Dict[str, Any]):
            # Replace assigned variables
            if "assign_to_variable_name" in d and d["assign_to_variable_name"] in mapping:
                d["assign_to_variable_name"] = mapping[d["assign_to_variable_name"]]
            if "assign_iteration_to_variable" in d and d.get("assign_iteration_to_variable") in mapping:
                d["assign_iteration_to_variable"] = mapping[d["assign_iteration_to_variable"]]

            # Replace used variables
            if d.get("type") == "variable" and d.get("value") in mapping:
                d["value"] = mapping[d["value"]]

            # Recurse
            for key, value in d.items():
                if isinstance(value, dict):
                    replace_in_dict(value)
                elif isinstance(value, list):
                    replace_in_list(value)
        
        def replace_in_list(lst: List[Any]):
            for item in lst:
                if isinstance(item, dict):
                    replace_in_dict(item)
                elif isinstance(item, list):
                    replace_in_list(item)
        
        replace_in_list(steps_copy)
        return steps_copy

    # --- REFACTORED/RENAMED METHOD ---
    def _load_template_by_name(self, template_name: str) -> None:
        """Loads a template file, handles variable mapping, and inserts it into the execution tree."""
        file_path = os.path.join(self.steps_template_directory, f"{template_name}.json")

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                template_steps = json.load(f)

            if not template_steps:
                QMessageBox.warning(self, "Empty Template", "The selected template is empty.")
                return

            # --- Variable Mapping Logic ---
            template_variables = self._extract_variables_from_steps(template_steps)
            mapped_steps = template_steps

            if template_variables:
                self._log_to_console(f"Template '{template_name}' contains variables: {template_variables}")
                var_dialog = TemplateVariableMappingDialog(template_variables, list(self.global_variables.keys()), self)
                if var_dialog.exec() == QDialog.DialogCode.Accepted:
                    mapping_result = var_dialog.get_mapping()
                    if mapping_result is None: # User error in dialog
                        return 
                
                    variable_map, new_vars = mapping_result
                    
                    # Apply the mapping to the steps
                    mapped_steps = self._apply_variable_mapping(template_steps, variable_map)
                    
                    # Add new variables to global context
                    if new_vars:
                        self.global_variables.update(new_vars)
                        self._update_variables_list_display()
                        self._log_to_console(f"Added new global variables: {list(new_vars.keys())}")
                else: # User cancelled the mapping
                    self._log_to_console("Template loading cancelled by user at variable mapping stage.")
                    return
            # --- END: Variable Logic ---

            re_id_d_steps = self._re_id_template_blocks(mapped_steps)

            insertion_dialog = StepInsertionDialog(self.execution_tree, parent=self)
            if insertion_dialog.exec() == QDialog.DialogCode.Accepted:
                selected_tree_item, insert_mode = insertion_dialog.get_insertion_point()
                insert_data_index = self._calculate_flat_insertion_index(selected_tree_item, insert_mode)
                
                self.added_steps_data[insert_data_index:insert_data_index] = re_id_d_steps
                
                self._rebuild_execution_tree(item_to_focus_data=re_id_d_steps[0])
                self._log_to_console(f"Loaded template '{template_name}' with {len(re_id_d_steps)} steps.")

        except FileNotFoundError:
            QMessageBox.critical(self, "Load Error", f"Template file not found:\n{file_path}")
        except json.JSONDecodeError:
            QMessageBox.critical(self, "Load Error", f"The template file '{template_name}.json' is corrupted.")
        except Exception as e:
            QMessageBox.critical(self, "Load Error", f"An unexpected error occurred while loading the template:\n{e}")
            
    def _re_id_template_blocks(self, template_steps: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        id_map = {}
        
        for step in template_steps:
            step_type = step.get("type")
            if step_type == "loop_start":
                old_id = step.get("loop_id")
                if old_id and old_id not in id_map:
                    self.loop_id_counter += 1
                    id_map[old_id] = f"loop_{self.loop_id_counter}"
            elif step_type == "IF_START":
                old_id = step.get("if_id")
                if old_id and old_id not in id_map:
                    self.if_id_counter += 1
                    id_map[old_id] = f"if_{self.if_id_counter}"
            elif step_type == "group_start":
                old_id = step.get("group_id")
                if old_id and old_id not in id_map:
                    self.group_id_counter += 1
                    id_map[old_id] = f"group_{self.group_id_counter}"
        
        for step in template_steps:
            if "loop_id" in step and step["loop_id"] in id_map:
                step["loop_id"] = id_map[step["loop_id"]]
            if "if_id" in step and step["if_id"] in id_map:
                step["if_id"] = id_map[step["if_id"]]
            if "group_id" in step and step["group_id"] in id_map:
                step["group_id"] = id_map[step["group_id"]]
                
        return template_steps

    def update_selected_method_info(self, item: QTreeWidgetItem, column: int) -> None:
        method_info_tuple = self._get_item_data(item)
        self.label_info2.clear()
        if method_info_tuple and isinstance(method_info_tuple, tuple) and len(method_info_tuple) == 5:
            display_text, class_name, method_name, module_name, params_for_dialog = method_info_tuple
            #self._log_to_console(f"Selected Method: {class_name}.{method_name} (from {module_name})")
            self.gui_communicator.update_module_info_signal.emit("")
        else:
            pass
            #self._log_to_console("Selected item is not a method or template.")
            #self.gui_communicator.update_module_info_signal.emit("")

    # --- MODIFIED METHOD ---
    def add_item_to_execution_tree(self, item: QTreeWidgetItem, column: int) -> None:
        item_data = self._get_item_data(item)

        # --- Handle template loading ---
        if isinstance(item_data, dict) and item_data.get('type') == 'template':
            template_name = item_data.get('name')
            if template_name:
                self._load_template_by_name(template_name)
            return

        # --- Handle method adding ---
        if not (item_data and isinstance(item_data, tuple) and len(item_data) == 5):
            return # Not a clickable method or template
            
        display_text, class_name, method_name, module_name, params_for_dialog = item_data
        
        self.active_param_input_dialog = ParameterInputDialog(f"{class_name}.{method_name}", params_for_dialog, list(self.global_variables.keys()), self._get_image_filenames(), self.gui_communicator, parent=self)
        self.active_param_input_dialog.request_screenshot.connect(self._handle_screenshot_request_from_param_dialog)
        if self.active_param_input_dialog.exec() == QDialog.DialogCode.Accepted:
            parameters_config = self.active_param_input_dialog.get_parameters_config()
            if parameters_config is None:
                self.gui_communicator.update_module_info_signal.emit("")
                return
            assign_to_variable_name = self.active_param_input_dialog.get_assignment_variable()
        else:
            self.gui_communicator.update_module_info_signal.emit("")
            return
        
        new_step_data_dict: Dict[str, Any] = {"type": "step", "class_name": class_name, "method_name": method_name, "module_name": module_name, "parameters_config": parameters_config, "assign_to_variable_name": assign_to_variable_name}
        insertion_dialog = StepInsertionDialog(self.execution_tree, parent=self)
        if insertion_dialog.exec() == QDialog.DialogCode.Accepted:
            selected_tree_item, insert_mode = insertion_dialog.get_insertion_point()
            insert_data_index = self._calculate_flat_insertion_index(selected_tree_item, insert_mode)
            self.added_steps_data.insert(insert_data_index, new_step_data_dict)
            self._rebuild_execution_tree(item_to_focus_data=new_step_data_dict)
        
        self.gui_communicator.update_module_info_signal.emit("")
        self.active_param_input_dialog = None

    def _calculate_flat_insertion_index(self, selected_tree_item: Optional[QTreeWidgetItem], insert_mode: str) -> int:
        if selected_tree_item is None:
            return len(self.added_steps_data)
        selected_item_data_in_dialog = self._get_item_data(selected_tree_item)
        if not selected_item_data_in_dialog:
            return len(self.added_steps_data)
        try:
            selected_flat_index = self.added_steps_data.index(selected_item_data_in_dialog)
        except ValueError:
            return len(self.added_steps_data)
        
        if insert_mode == "before":
            return selected_flat_index
        elif insert_mode == "after":
            _, end_idx = self._find_block_indices(selected_flat_index)
            return end_idx + 1
        elif insert_mode == "child":
            block_type = selected_item_data_in_dialog.get("type")
            if block_type in ["group_start", "loop_start", "IF_START", "ELSE"]:
                return selected_flat_index + 1
            else:
                return selected_flat_index + 1
        return len(self.added_steps_data)

    def add_loop_block(self) -> None:
        dialog = LoopConfigDialog(self.global_variables, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            loop_config = dialog.get_config()
            if loop_config is None:
                return
            self.loop_id_counter += 1
            loop_id = f"loop_{self.loop_id_counter}"
            new_loop_start_data: Dict[str, Any] = {"type": "loop_start", "loop_id": loop_id, "loop_config": loop_config}
            insertion_dialog = StepInsertionDialog(self.execution_tree, parent=self)
            if insertion_dialog.exec() == QDialog.DialogCode.Accepted:
                selected_tree_item, insert_mode = insertion_dialog.get_insertion_point()
                insert_data_index = self._calculate_flat_insertion_index(selected_tree_item, insert_mode)
                self.added_steps_data.insert(insert_data_index, new_loop_start_data)
                loop_end_data_dict = {"type": "loop_end", "loop_id": loop_id, "loop_config": loop_config}
                self.added_steps_data.insert(insert_data_index + 1, loop_end_data_dict)
                self._update_original_listbox_row_indices()
                self._rebuild_execution_tree(item_to_focus_data=new_loop_start_data)

    def add_conditional_block(self) -> None:
        dialog = ConditionalConfigDialog(self.global_variables, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            conditional_config = dialog.get_config()
            if conditional_config is None:
                return
            self.if_id_counter += 1
            if_id = f"if_{self.if_id_counter}"
            new_if_start_data: Dict[str, Any] = {"type": "IF_START", "if_id": if_id, "condition_config": conditional_config}
            new_else_data: Dict[str, Any] = {"type": "ELSE", "if_id": if_id, "condition_config": conditional_config}
            new_if_end_data: Dict[str, Any] = {"type": "IF_END", "if_id": if_id, "condition_config": conditional_config}
            insertion_dialog = StepInsertionDialog(self.execution_tree, parent=self)
            if insertion_dialog.exec() == QDialog.DialogCode.Accepted:
                selected_tree_item, insert_mode = insertion_dialog.get_insertion_point()
                insert_data_index = self._calculate_flat_insertion_index(selected_tree_item, insert_mode)
                self.added_steps_data.insert(insert_data_index, new_if_start_data)
                self.added_steps_data.insert(insert_data_index + 1, new_else_data)
                self.added_steps_data.insert(insert_data_index + 2, new_if_end_data)
                self._update_original_listbox_row_indices()
                self._rebuild_execution_tree(item_to_focus_data=new_if_start_data)

    def group_selected_steps(self) -> None:
        selected_items = self.execution_tree.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "No Selection", "Please select one or more steps to group.")
            return

        group_name, ok = QInputDialog.getText(self, "Create Group", "Enter a name for the group:")
        if ok and group_name:
            selected_data = [self._get_item_data(item) for item in selected_items]
            
            first_item_data = selected_data[0]
            try:
                start_index = self.added_steps_data.index(first_item_data)
            except ValueError:
                QMessageBox.critical(self, "Error", "Could not find the start of the selection.")
                return

            last_item_data = selected_data[-1]
            try:
                last_item_start_idx_in_flat_list = self.added_steps_data.index(last_item_data)
                _, end_index = self._find_block_indices(last_item_start_idx_in_flat_list)
            except ValueError:
                QMessageBox.critical(self, "Error", "Could not find the end of the selection.")
                return

            self.group_id_counter += 1
            group_id = f"group_{self.group_id_counter}"
            
            group_start_data = {"type": "group_start", "group_id": group_id, "group_name": group_name}
            group_end_data = {"type": "group_end", "group_id": group_id}

            self.added_steps_data.insert(start_index, group_start_data)
            self.added_steps_data.insert(end_index + 2, group_end_data)

            self._rebuild_execution_tree(item_to_focus_data=group_start_data)

    def edit_step_in_execution_tree(self, item: QTreeWidgetItem, column: int) -> None:
        step_data_dict = self._get_item_data(item)
        if not step_data_dict or not isinstance(step_data_dict, dict):
            QMessageBox.warning(self, "Invalid Item", "Cannot edit this item type or no data found.")
            return
        step_type = step_data_dict["type"]
        current_row = -1
        try:
            current_row = self.added_steps_data.index(step_data_dict)
        except ValueError:
            QMessageBox.critical(self, "Error", "Could not find selected item in internal data model for editing.")
            return
        if step_type == "step":
            class_name, method_name, module_name = step_data_dict["class_name"], step_data_dict["method_name"], step_data_dict["module_name"]
            parameters_config_with_index = step_data_dict["parameters_config"]
            assign_to_variable_name = step_data_dict["assign_to_variable_name"]
            dialog_parameters_config = {k: v for k, v in parameters_config_with_index.items() if k != "original_listbox_row_index"}
            try:
                if self.module_directory not in sys.path:
                    sys.path.insert(0, self.module_directory)
                module = importlib.import_module(module_name)
                importlib.reload(module)
                class_obj = getattr(module, class_name)
                init_kwargs_for_inspection: Dict[str, Any] = {}
                if 'context' in inspect.signature(class_obj.__init__).parameters:
                    init_kwargs_for_inspection['context'] = ExecutionContext()
                temp_instance = class_obj(**init_kwargs_for_inspection)
                method_func_obj = getattr(temp_instance, method_name)
                func_to_inspect = method_func_obj.__func__ if inspect.ismethod(method_func_obj) else method_func_obj
                sig = inspect.signature(func_to_inspect)
                params_for_dialog: Dict[str, Tuple[Any, Any]] = {p.name: (p.default, p.kind) for p in sig.parameters.values() if p.name not in ['self', 'context']}
            except Exception as e:
                QMessageBox.critical(self, "Error Editing Step", f"Could not re-inspect method for editing:\n{e}")
                return
            finally:
                if self.module_directory in sys.path:
                    sys.path.remove(self.module_directory)
            self.active_param_input_dialog = ParameterInputDialog(f"{class_name}.{method_name}", params_for_dialog, list(self.global_variables.keys()), self._get_image_filenames(), self.gui_communicator, initial_parameters_config=dialog_parameters_config, initial_assign_to_variable_name=assign_to_variable_name, parent=self)
            self.active_param_input_dialog.request_screenshot.connect(self._handle_screenshot_request_from_param_dialog)
            if self.active_param_input_dialog.exec() == QDialog.DialogCode.Accepted:
                new_parameters_config = self.active_param_input_dialog.get_parameters_config()
                if new_parameters_config is None:
                    self.gui_communicator.update_module_info_signal.emit("")
                    return
                new_assign_to_variable_name = self.active_param_input_dialog.get_assignment_variable()
                self.added_steps_data[current_row].update({"parameters_config": new_parameters_config, "assign_to_variable_name": new_assign_to_variable_name})
                self._rebuild_execution_tree(item_to_focus_data=self.added_steps_data[current_row])
            self.gui_communicator.update_module_info_signal.emit("")
            self.active_param_input_dialog = None
        elif step_type == "loop_start":
            loop_config = step_data_dict["loop_config"]
            dialog = LoopConfigDialog(self.global_variables, parent=self, initial_config=loop_config)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                new_loop_config = dialog.get_config()
                if new_loop_config is None:
                    return
                self.added_steps_data[current_row]["loop_config"] = new_loop_config
                loop_id, nesting_level = step_data_dict["loop_id"], 0
                for idx in range(current_row + 1, len(self.added_steps_data)):
                    step = self.added_steps_data[idx]
                    if step.get("type") in ["loop_start", "IF_START"]:
                        nesting_level += 1
                    elif step.get("type") == "loop_end" and nesting_level == 0 and step.get("loop_id") == loop_id:
                        self.added_steps_data[idx]["loop_config"] = new_loop_config
                        break
                    elif step.get("type") in ["loop_end", "IF_END"]:
                        nesting_level -= 1
                self._rebuild_execution_tree(item_to_focus_data=self.added_steps_data[current_row])
        elif step_type == "IF_START":
            condition_config = step_data_dict["condition_config"]
            dialog = ConditionalConfigDialog(self.global_variables, parent=self, initial_config=condition_config)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                new_conditional_config = dialog.get_config()
                if new_conditional_config is None:
                    return
                self.added_steps_data[current_row]["condition_config"] = new_conditional_config
                if_id, nesting_level = step_data_dict["if_id"], 0
                for idx in range(current_row + 1, len(self.added_steps_data)):
                    step = self.added_steps_data[idx]
                    if step.get("type") in ["loop_start", "IF_START"]:
                        nesting_level += 1
                    elif nesting_level == 0 and step.get("if_id") == if_id:
                        if step.get("type") == "ELSE":
                            self.added_steps_data[idx]["condition_config"] = new_conditional_config
                        elif step.get("type") == "IF_END":
                            self.added_steps_data[idx]["condition_config"] = new_conditional_config
                            break
                    elif step.get("type") in ["loop_end", "IF_END"]:
                        nesting_level -= 1
                self._rebuild_execution_tree(item_to_focus_data=self.added_steps_data[current_row])
        elif step_type == "group_start":
            group_name = step_data_dict.get("group_name", "")
            new_name, ok = QInputDialog.getText(self, "Rename Group", "Enter new group name:", text=group_name)
            if ok and new_name:
                self.added_steps_data[current_row]["group_name"] = new_name
                self._rebuild_execution_tree(item_to_focus_data=self.added_steps_data[current_row])
        elif step_type in ["loop_end", "ELSE", "IF_END", "group_end"]:
            QMessageBox.information(self, "Edit Block Marker", "To change parameters, edit the corresponding 'Start' block.")

    def clear_selected_steps(self) -> None:
        selected_items = self.execution_tree.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "No Selection", "Please select steps to clear.")
            return
        if QMessageBox.question(self, "Confirm Clear", f"Are you sure you want to remove {len(selected_items)} selected step(s)? This may break block structures.", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            selected_step_data = [self._get_item_data(item) for item in selected_items]
            self.added_steps_data = [s for s in self.added_steps_data if s not in selected_step_data]
            self._rebuild_execution_tree()

    def _internal_clear_all_steps(self):
        self.execution_tree.clear()
        self.added_steps_data.clear()
        self.global_variables.clear()
        self._update_variables_list_display()
        self.progress_bar.setValue(0)
        self.progress_bar.hide()
        self.set_ui_enabled_state(True)
        self.loop_id_counter = 0
        self.if_id_counter = 0
        self.group_id_counter = 0
        self._log_to_console("Internal clear all steps executed.")

    def clear_all_steps(self) -> None:
        if not self.added_steps_data and not self.global_variables:
            QMessageBox.information(self, "Info", "The execution queue and variables are already empty.")
            return
        if QMessageBox.question(self, "Confirm Remove All", "Are you sure you want to remove ALL steps and variables?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            self._internal_clear_all_steps()
            self._log_to_console("All steps cleared by user.")

# --- MODIFY execute_all_steps ---
    def execute_all_steps(self) -> None:
        # --- ADD THIS CHECK AT THE BEGINNING ---
        if self.is_bot_running:
            QMessageBox.warning(self, "Execution in Progress", "A bot is already running. Please wait for it to complete.")
            return
        # --- END ---

        if not self.added_steps_data:
            QMessageBox.information(self, "No Steps", "No steps have been added.")
            return
        if not self._validate_block_structure_on_execution():
            return
        
        # --- SET THE FLAG TO TRUE ---
        self.is_bot_running = True
        
        if self.always_on_top_button.isChecked():
            self._toggle_minimized_view(True)

        self.set_ui_enabled_state(False)
        self.progress_bar.setValue(0)
        self.progress_bar.show()
        self.update_status_column_for_all_items()
        self.worker = ExecutionWorker(self.added_steps_data, self.module_directory, self.gui_communicator, self.global_variables)
        self._connect_worker_signals()
        self.worker.start()

    def _validate_block_structure_on_execution(self) -> bool:
        open_blocks = []
        for step_data in self.added_steps_data:
            step_type = step_data["type"]
            if step_type == "group_start":
                open_blocks.append(("group", step_data["group_id"]))
            elif step_type == "loop_start":
                open_blocks.append(("loop", step_data["loop_id"]))
            elif step_type == "IF_START":
                open_blocks.append(("if", step_data["if_id"]))
            elif step_type == "ELSE":
                if not open_blocks or open_blocks[-1][0] != "if":
                    QMessageBox.warning(self, "Invalid Block Structure", "Mismatched ELSE block.")
                    return False
                if_id = open_blocks.pop()[1]
                open_blocks.append(("else", if_id))
            elif step_type == "group_end":
                if not open_blocks or open_blocks.pop() != ("group", step_data["group_id"]):
                    QMessageBox.warning(self, "Invalid Block Structure", "Mismatched GROUP block.")
                    return False
            elif step_type == "loop_end":
                if not open_blocks or open_blocks.pop() != ("loop", step_data["loop_id"]):
                    QMessageBox.warning(self, "Invalid Block Structure", "Mismatched LOOP block.")
                    return False
            elif step_type == "IF_END":
                if not open_blocks or open_blocks[-1][0] not in ["if", "else"] or open_blocks.pop()[1] != step_data["if_id"]:
                    QMessageBox.warning(self, "Invalid Block Structure", "Mismatched IF_END block.")
                    return False
        
        if open_blocks:
            QMessageBox.warning(self, "Invalid Block Structure", f"Unclosed blocks remain: {open_blocks}.")
            return False
        return True

    def execute_one_step(self) -> None:
        selected_items = self.execution_tree.selectedItems()
        if len(selected_items) != 1:
            QMessageBox.information(self, "Selection Error", "Please select exactly ONE step to execute.")
            return
        selected_step_data = self._get_item_data(selected_items[0])
        if not (selected_step_data and isinstance(selected_step_data, dict)):
            return
        try:
            current_row = self.added_steps_data.index(selected_step_data)
        except ValueError:
            QMessageBox.critical(self, "Error", "Could not find selected step in internal data model.")
            return
        self.set_ui_enabled_state(False)
        self.progress_bar.setValue(0)
        self.progress_bar.show()
        self.update_status_column_for_all_items()
        self.worker = ExecutionWorker(self.added_steps_data, self.module_directory, self.gui_communicator, self.global_variables, single_step_mode=True, selected_start_index=current_row)
        self._connect_worker_signals()
        self.worker.start()

    def _connect_worker_signals(self) -> None:
        try:
            self.worker.execution_started.disconnect()
            self.worker.execution_progress.disconnect()
            self.worker.execution_item_started.disconnect()
            self.worker.execution_item_finished.disconnect()
            self.worker.execution_error.disconnect()
            self.worker.execution_finished_all.disconnect()
        except TypeError:
            pass
        self.worker.execution_started.connect(lambda msg: self._log_to_console(f"GUI: {msg}"))
        self.worker.execution_progress.connect(self.progress_bar.setValue)
        self.worker.execution_item_started.connect(self.update_execution_tree_item_status_started)
        self.worker.execution_item_finished.connect(self.update_execution_tree_item_status_finished)
        self.worker.execution_error.connect(self.update_execution_tree_item_status_error)
        self.worker.execution_finished_all.connect(self.on_execution_finished)
        self.worker.execution_finished_all.connect(lambda: self._update_variables_list_display())

    def _find_qtreewidget_item(self, target_step_data_dict: Dict[str, Any], parent_item: Optional[QTreeWidgetItem] = None) -> Optional[QTreeWidgetItem]:
        if parent_item is None:
            parent_item = self.execution_tree.invisibleRootItem()
        for i in range(parent_item.childCount()):
            child_item = parent_item.child(i)
            item_data = self._get_item_data(child_item)
            if item_data == target_step_data_dict:
                return child_item
            found_in_children = self._find_qtreewidget_item(target_step_data_dict, child_item)
            if found_in_children:
                return found_in_children
        return None

    def update_execution_tree_item_status_started(self, step_data_dict: Dict[str, Any], original_listbox_row_index: int) -> None:
        item_widget = self._find_qtreewidget_item(step_data_dict)
        if item_widget:
            card = self.execution_tree.itemWidget(item_widget, 0)
            if card:
                card.set_status("blue", is_running=True)
                self._log_to_console(f"Executing: {card._get_formatted_title()}")
            self.execution_tree.setCurrentItem(item_widget)
            self.execution_tree.scrollToItem(item_widget, QtWidgets.QAbstractItemView.ScrollHint.PositionAtCenter)

    def update_execution_tree_item_status_finished(self, step_data_dict: Dict[str, Any], message: str, original_listbox_row_index: int) -> None:
        item_widget = self._find_qtreewidget_item(step_data_dict)
        if item_widget:
            card = self.execution_tree.itemWidget(item_widget, 0)
            if card:
                card.set_status("darkGreen", is_running=False)
                card.set_result_text(message)
            self._log_to_console(f"Finished: {card._get_formatted_title()} | {message}")
            self._update_variables_list_display()

    def update_execution_tree_item_status_error(self, step_data_dict: Dict[str, Any], error_message: str, original_listbox_row_index: int) -> None:
        item_widget = self._find_qtreewidget_item(step_data_dict)
        if item_widget:
            card = self.execution_tree.itemWidget(item_widget, 0)
            if card:
                card.set_status("red", is_running=False)
            self._log_to_console(f"ERROR on {card._get_formatted_title()}: {error_message}")
            self._update_variables_list_display()

# --- MODIFY on_execution_finished ---
    def on_execution_finished(self, context: ExecutionContext, stopped_by_error: bool, next_step_index_to_select: int) -> None:
        # --- SET THE FLAG BACK TO FALSE ---
        self.is_bot_running = False
        # --- END ---

        self.progress_bar.setValue(100)
        self.set_ui_enabled_state(True)
        self.last_executed_context = context
        
        if self.minimized_for_execution:
            self._toggle_minimized_view(False)

        if stopped_by_error:
            QMessageBox.critical(self, "Execution Halted", "Execution stopped due to an error.")
            self._log_to_console("Execution STOPPED due to an error.")
        else:
            # Making this message less intrusive for scheduled tasks
            if not self.minimized_for_execution:
                 QMessageBox.information(self, "Execution Complete", "All steps processed.")
            self._log_to_console("Execution finished successfully.")
        
        if next_step_index_to_select != -1 and 0 <= next_step_index_to_select < len(self.added_steps_data):
            target_item = self._find_qtreewidget_item(self.added_steps_data[next_step_index_to_select])
            if target_item:
                self.execution_tree.setCurrentItem(target_item)
                self.execution_tree.scrollToItem(target_item)
        
        # --- REFRESH THE TREE VIEW ---
        # This ensures schedule status is updated after a run
        self.load_saved_steps_to_tree()
        # --- END ---

    def update_label_info_from_module(self, message: str) -> None:
        if message and not message.startswith("Module Info:") and not message.startswith("Last Module Log:"):
            image_filepath = os.path.join(self.click_image_dir, f"{message}.txt")
            if os.path.exists(image_filepath):
                self.label_info3.setText(message)
                try:
                    with open(image_filepath, 'r') as json_file:
                        img_data = json.load(json_file)
                        base64_string = next(iter(img_data.values()), None)
                    if base64_string:
                        pic_png = self.base64_pgn(base64_string)
                        qimage = ImageQt(pic_png)
                        pixmap = self.resize_qimage_and_create_qpixmap(qimage)
                        self.label_info2.setPixmap(pixmap)
                    else:
                        self.label_info1.setText(f"No image data found for '{message}'.")
                except Exception as e:
                    self.label_info1.setText(f"Error loading image '{message}': {e}")
            else:
                self.label_info1.setText(f"Image file '{message}.txt' not found.")
        elif message == "":
            self.label_info2.clear()
        else:
            self.label_info1.setText(f"Last Module Log: {message}")
            if not self.module_filter_dropdown.currentText() or self.module_filter_dropdown.currentText() == "-- Show All Modules --":
                self.label_info2.clear()
            self.label_info2.setText("Image Preview")

    def base64_pgn(self,text):
        return PIL.Image.open(io.BytesIO(base64.b64decode(text)))

    def resize_qimage_and_create_qpixmap(self,qimage_input, percentage=98):
        if qimage_input.isNull():
            return QPixmap()
        new_width = int(qimage_input.width() * (percentage / 100))
        new_height = int(qimage_input.height() * (percentage / 100))
        return QPixmap.fromImage(qimage_input.scaled(QSize(new_width, new_height), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))

# In the MainWindow class, REPLACE your existing method with this one

    def set_ui_enabled_state(self, enabled: bool) -> None:
        widgets_to_toggle = [
            self.execute_all_button, self.add_loop_button, self.add_conditional_button, 
            self.save_steps_button, self.clear_selected_button, self.remove_all_steps_button,
            self.module_filter_dropdown, self.module_tree, self.saved_steps_tree, # Corrected line
            self.add_var_button, self.edit_var_button, self.delete_var_button, 
            self.clear_vars_button, self.open_screenshot_tool_button, 
            self.group_steps_button
        ]
        for widget in widgets_to_toggle:
            widget.setEnabled(enabled)
        self.execute_one_step_button.setEnabled(enabled and len(self.execution_tree.selectedItems()) > 0)
        
    def _update_variables_list_display(self) -> None:
        self.variables_list.clear()
        if not self.global_variables:
            self.variables_list.addItem("No global variables defined.")
            return
        for name, value in self.global_variables.items():
            value_str = repr(value)
            if len(value_str) > 60:
                value_str = value_str[:57] + "..."
            self.variables_list.addItem(f"{name} = {value_str}")

    def add_variable(self) -> None:
        dialog = GlobalVariableDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            name, value = dialog.get_variable_data()
            if name:
                if name in self.global_variables:
                    QMessageBox.warning(self, "Duplicate Variable", f"Variable '{name}' already exists.")
                    return
                self.global_variables[name] = value
                self._update_variables_list_display()

    def edit_variable(self) -> None:
        selected_item = self.variables_list.currentItem()
        if not selected_item or selected_item.text() == "No global variables defined.":
            QMessageBox.information(self, "No Selection", "Please select a variable to edit.")
            return
        var_name = selected_item.text().split(' = ')[0]
        if var_name not in self.global_variables:
            self._update_variables_list_display()
            return
        dialog = GlobalVariableDialog(variable_name=var_name, variable_value=self.global_variables[var_name], parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_name, new_value = dialog.get_variable_data()
            if new_name:
                self.global_variables[new_name] = new_value
            self._update_variables_list_display()

    def delete_variable(self) -> None:
        selected_item = self.variables_list.currentItem()
        if not selected_item or selected_item.text() == "No global variables defined.":
            QMessageBox.information(self, "No Selection", "Please select a variable to delete.")
            return
        var_name = selected_item.text().split(' = ')[0]
        if QMessageBox.question(self, "Confirm Delete", f"Are you sure you want to delete variable '{var_name}'?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            if var_name in self.global_variables:
                del self.global_variables[var_name]
            self._update_variables_list_display()

    def reset_all_variable_values(self) -> None:
        if not self.global_variables:
            QMessageBox.information(self, "Info", "No global variables to reset.")
            return
        if QMessageBox.question(self, "Confirm Reset", "Reset all global variable values to 'None'?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            for var_name in self.global_variables:
                self.global_variables[var_name] = None
            self._update_variables_list_display()

# In the MainWindow class, REPLACE your existing save_bot_steps_dialog method with this one

    def save_bot_steps_dialog(self) -> None:
        """Opens a custom dialog to get a name and save the bot steps,
        sanitizing any non-serializable global variables."""
        if not self.added_steps_data and not self.global_variables:
            QMessageBox.information(self, "Nothing to Save", "The execution queue and global variables are empty.")
            return
        
        os.makedirs(self.bot_steps_directory, exist_ok=True)
        existing_bots = [os.path.splitext(f)[0] for f in os.listdir(self.bot_steps_directory) if f.endswith(".csv")]
        
        dialog = SaveBotDialog(existing_bots, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            bot_name = dialog.get_bot_name()
            if not bot_name:
                return
    
            if bot_name in existing_bots:
                reply = QMessageBox.question(self, "Confirm Overwrite",
                                             f"A bot named '{bot_name}' already exists. Overwrite it?",
                                             QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if reply == QMessageBox.StandardButton.No:
                    return
    
            file_path = os.path.join(self.bot_steps_directory, f"{bot_name}.csv")
            try:
                # --- NEW: Create a sanitized copy of global variables for saving ---
                variables_to_save = {}
                for var_name, var_value in self.global_variables.items():
                    try:
                        # Attempt to serialize the value to see if it's valid
                        json.dumps(var_value)
                        variables_to_save[var_name] = var_value
                    except (TypeError, OverflowError):
                        # If it fails, the value is a non-serializable object. Set to None.
                        variables_to_save[var_name] = None
                # --- End of new section ---
    
                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(["__GLOBAL_VARIABLES__"])
                    # Use the sanitized dictionary for writing
                    for var_name, var_value in variables_to_save.items():
                        writer.writerow([var_name, json.dumps(var_value)])
                    writer.writerow(["__BOT_STEPS__"])
                    writer.writerow(["StepType", "DataJSON"])
                    for step_data_dict in self.added_steps_data:
                        step_data_to_save = json.loads(json.dumps(step_data_dict))
                        step_data_to_save.pop("original_listbox_row_index", None)
                        if step_data_to_save.get("type") == "step" and step_data_to_save.get("parameters_config"):
                            step_data_to_save["parameters_config"].pop("original_listbox_row_index", None)
                        writer.writerow([step_data_to_save["type"], json.dumps(step_data_to_save)])
                QMessageBox.information(self, "Save Successful", f"Bot saved to:\n{file_path}\n(Object variables were reset to 'None' in the file.)")
                self.load_saved_steps_to_tree()
            except Exception as e:
                QMessageBox.critical(self, "Save Error", f"Failed to save bot steps:\n{e}")
    

    def load_steps_from_file(self, file_path: str) -> None:
        self._internal_clear_all_steps()
        try:
            section, loaded_variables, loaded_steps = None, {}, []
            with open(file_path, 'r', newline='', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                for row_num, row in enumerate(reader):
                    if not row:
                        continue
                    if row[0] == "__GLOBAL_VARIABLES__":
                        section = "VARIABLES"
                        continue
                    elif row[0] == "__BOT_STEPS__":
                        section = "STEPS"
                        next(reader, None)
                        continue
                    if section == "VARIABLES":
                        if len(row) == 2:
                            loaded_variables[row[0]] = json.loads(row[1])
                        else:
                            self._log_to_console(f"Warning: Malformed variable row {row_num+1} in {file_path}")
                    elif section == "STEPS":
                        if len(row) == 2:
                            step_data_json = json.loads(row[1])
                            if step_data_json.get("type") == "group_start":
                                try:
                                    self.group_id_counter = max(self.group_id_counter, int(step_data_json.get("group_id", "group_0").split('_')[-1]))
                                except (ValueError, IndexError):
                                    pass
                            elif step_data_json.get("type") == "loop_start":
                                try:
                                    self.loop_id_counter = max(self.loop_id_counter, int(step_data_json.get("loop_id", "loop_0").split('_')[-1]))
                                except (ValueError, IndexError):
                                    pass
                            elif step_data_json.get("type") == "IF_START":
                                try:
                                    self.if_id_counter = max(self.if_id_counter, int(step_data_json.get("if_id", "if_0").split('_')[-1]))
                                except (ValueError, IndexError):
                                    pass
                            loaded_steps.append(step_data_json)
                        else:
                            self._log_to_console(f"Warning: Malformed step row {row_num+1} in {file_path}")
            self.global_variables = loaded_variables
            self._update_variables_list_display()
            if not loaded_steps and not loaded_variables:
                QMessageBox.warning(self, "Load Warning", "No valid steps or variables were found in the file.")
                return
            self.added_steps_data = loaded_steps
            self._rebuild_execution_tree()
            QMessageBox.information(self, "Load Successful", f"{len(loaded_steps)} steps and {len(loaded_variables)} variables loaded.")
        except FileNotFoundError:
            QMessageBox.critical(self, "Load Error", f"The file was not found:\n{file_path}")
        except json.JSONDecodeError as e:
            QMessageBox.critical(self, "Load Error", f"The file is corrupted or not a valid step file.\nError parsing data: {e}")
        except Exception as e:
            QMessageBox.critical(self, "Load Error", f"An unexpected error occurred while loading the file:\n{e}")
            self._log_to_console(f"Load Error: {e}")
    # REPLACE your old show_method_context_menu with this one
    # In the MainWindow class, REPLACE your old show_context_menu with this one

    def show_context_menu(self, position: QPoint):
        item = self.module_tree.itemAt(position)
        if not item:
            return
    
        item_data = self._get_item_data(item)
    
        # --- Handle Methods ---
        if isinstance(item_data, tuple) and len(item_data) == 5:
            _, class_name, method_name, module_name, _ = item_data
            context_menu = QMenu(self)
            read_doc_action = context_menu.addAction("Read Documentation")
            modify_action = context_menu.addAction("Modify Method")
            delete_action = context_menu.addAction("Delete Method")
            action = context_menu.exec(self.module_tree.mapToGlobal(position))
            if action == read_doc_action:
                self.read_method_documentation(module_name, class_name, method_name)
            elif action == modify_action:
                self.modify_method(module_name, class_name, method_name)
            elif action == delete_action:
                self.delete_method(module_name, class_name, method_name)
    
        # --- Handle Templates (UPDATED) ---
        elif isinstance(item_data, dict) and item_data.get('type') == 'template':
            template_name = item_data.get('name')
            if not template_name:
                return
    
            context_menu = QMenu(self)
            add_action = context_menu.addAction("Add to Execution Flow")
            context_menu.addSeparator()
            doc_action = context_menu.addAction("View Documentation")
            delete_action = context_menu.addAction("Delete Template")
            
            doc_path = os.path.join(self.template_document_directory, f"{template_name}.html")
            doc_action.setEnabled(os.path.exists(doc_path))
    
            action = context_menu.exec(self.module_tree.mapToGlobal(position))
    
            if action == add_action:
                self._load_template_by_name(template_name)
            elif action == doc_action:
                self.view_template_documentation(template_name)
            elif action == delete_action:
                self.delete_template(template_name)

    # Add this new method to the MainWindow class
    def view_template_documentation(self, template_name: str):
        """Finds and displays the HTML documentation for a given template."""
        os.makedirs(self.template_document_directory, exist_ok=True) # Ensure folder exists
    
        doc_path = os.path.join(self.template_document_directory, f"{template_name}.html")
    
        if os.path.exists(doc_path):
            dialog = HtmlViewerDialog(doc_path, self)
            dialog.exec()
        else:
            QMessageBox.information(self, "Documentation Not Found",
                                      f"No documentation file found at:\n{doc_path}")

    # Add this new method to the MainWindow class
    
    def delete_template(self, template_name: str):
        """Deletes a template file after user confirmation."""
        reply = QMessageBox.question(self, "Confirm Delete",
                                     f"Are you sure you want to permanently delete the template '{template_name}'?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
    
        if reply == QMessageBox.StandardButton.Yes:
            try:
                template_path = os.path.join(self.steps_template_directory, f"{template_name}.json")
                if os.path.exists(template_path):
                    os.remove(template_path)
                    QMessageBox.information(self, "Success", f"Template '{template_name}' has been deleted.")
                    # Refresh the module tree to reflect the deletion
                    self.load_all_modules_to_tree()
                else:
                    QMessageBox.warning(self, "File Not Found", f"The template file for '{template_name}' could not be found.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"An error occurred while deleting the template:\n{e}")
    # Add this method to the MainWindow class
    def read_method_documentation(self, module_name: str, class_name: str, method_name: str):
        try:
            if self.module_directory not in sys.path:
                sys.path.insert(0, self.module_directory)
            module = importlib.import_module(module_name)
            importlib.reload(module)
            class_obj = getattr(module, class_name)
            method_obj = getattr(class_obj, method_name)
            docstring = inspect.getdoc(method_obj)
            if not docstring:
                docstring = "No documentation found for this method."
            QMessageBox.information(self, f"Documentation for {class_name}.{method_name}", docstring)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not read documentation: {e}")
        finally:
            if self.module_directory in sys.path:
                sys.path.remove(self.module_directory)

    # Add this method to the MainWindow class
    # In MainWindow class, REPLACE the old modify_method with this:
    def modify_method(self, module_name: str, class_name: str, method_name: str):
        try:
            # We need the full path to the module file
            module_path = os.path.join(self.module_directory, f"{module_name}.py")
            if not os.path.exists(module_path):
                QMessageBox.critical(self, "File Not Found", f"The source file for module '{module_name}' could not be found.")
                return
    
            # Create and show the editor dialog
            editor_dialog = CodeEditorDialog(module_path, class_name, method_name, self.all_parsed_method_data, self)
            editor_dialog.exec()
            
            # After saving, reload the modules to reflect any changes
            self.load_all_modules_to_tree()
    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not open the code editor: {e}")

    def load_saved_steps_to_tree(self) -> None:
        """Loads saved bot step files into the QTreeWidget."""
        self.saved_steps_tree.clear()
        try:
            os.makedirs(self.bot_steps_directory, exist_ok=True)
            step_files = sorted([f for f in os.listdir(self.bot_steps_directory) if f.endswith(".csv")], reverse=True)
            for file_name in step_files:
                bot_name = os.path.splitext(file_name)[0]
                # Create a tree item with placeholder text for the new columns
                tree_item = QTreeWidgetItem(self.saved_steps_tree, [bot_name, "Not Set", "Idle"])
            if not step_files:
                self.saved_steps_tree.addTopLevelItem(QTreeWidgetItem(["No saved bots found."]))
        except Exception as e:
            QMessageBox.critical(self, "Error Loading Saved Bots", f"Could not load bot files: {e}")

# --- MODIFY check_schedules ---
    def check_schedules(self):
        """Timer-triggered function to check for and run scheduled bots."""
        # --- ADD THIS CHECK AT THE BEGINNING ---
        if self.is_bot_running:
            self._log_to_console("Scheduler check skipped: A bot is currently running.")
            return
        # --- END ---

        self._log_to_console("Scheduler checking for due bots...")
        now = QDateTime.currentDateTime()
        
        bot_files = [f for f in os.listdir(self.bot_steps_directory) if f.endswith(".csv")]

        for bot_file in bot_files:
            # --- ADD ANOTHER CHECK INSIDE THE LOOP ---
            if self.is_bot_running:
                self._log_to_console("Scheduler check halted: A bot started execution during the check.")
                break # Exit the loop if a bot has started
            # --- END ---

            bot_name = os.path.splitext(bot_file)[0]
            file_path = os.path.join(self.bot_steps_directory, bot_file)
            schedule_data = self._read_schedule_from_csv(file_path)

            if schedule_data and schedule_data.get("enabled"):
                start_datetime = QDateTime.fromString(schedule_data.get("start_datetime"), Qt.DateFormat.ISODate)

                if now >= start_datetime:
                    self._log_to_console(f"Executing scheduled bot: '{bot_name}'")
                    
                    self.load_steps_from_file(file_path)
                    self.execute_all_steps() # This will now set the is_bot_running flag

                    # The rest of the logic for rescheduling remains the same...
                    repeat_mode = schedule_data.get("repeat")
                    if repeat_mode != "Do not repeat":
                        new_start_time = start_datetime
                        if repeat_mode == "Hourly": new_start_time = new_start_time.addSecs(3600)
                        elif repeat_mode == "Daily": new_start_time = new_start_time.addDays(1)
                        elif repeat_mode == "Monthly": new_start_time = new_start_time.addMonths(1)
                        
                        while new_start_time < now:
                            if repeat_mode == "Hourly": new_start_time = new_start_time.addSecs(3600)
                            elif repeat_mode == "Daily": new_start_time = new_start_time.addDays(1)
                            elif repeat_mode == "Monthly": new_start_time = new_start_time.addMonths(1)

                        schedule_data["start_datetime"] = new_start_time.toString(Qt.DateFormat.ISODate)
                        self._write_schedule_to_csv(file_path, schedule_data)
                        self._log_to_console(f"Rescheduled '{bot_name}' to run next at {schedule_data['start_datetime']}")
                    else:
                        schedule_data["enabled"] = False
                        self._write_schedule_to_csv(file_path, schedule_data)
                        self._log_to_console(f"Disabled non-repeating schedule for '{bot_name}'.")


    def schedule_bot(self, bot_name: str):
        """Opens the scheduling dialog and saves the result to the bot's CSV file."""
        file_path = os.path.join(self.bot_steps_directory, f"{bot_name}.csv")
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "File Not Found", f"The bot file '{bot_name}.csv' could not be found to save the schedule.")
            return

        schedule_data = self._read_schedule_from_csv(file_path)
        dialog = ScheduleTaskDialog(bot_name, schedule_data, self)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_schedule_data = dialog.get_schedule_data()
            if self._write_schedule_to_csv(file_path, new_schedule_data):
                self._log_to_console(f"Schedule for '{bot_name}' has been updated.")
                self.load_saved_steps_to_tree() # Refresh the view
            else:
                QMessageBox.critical(self, "Error", f"Failed to save schedule for '{bot_name}'.")

    # --- REPLACE load_saved_steps_to_tree ---
    def load_saved_steps_to_tree(self) -> None:
        """Loads saved bot step files and their schedules into the QTreeWidget."""
        self.saved_steps_tree.clear()
        try:
            os.makedirs(self.bot_steps_directory, exist_ok=True)
            step_files = sorted([f for f in os.listdir(self.bot_steps_directory) if f.endswith(".csv")], reverse=True)
            for file_name in step_files:
                bot_name = os.path.splitext(file_name)[0]
                file_path = os.path.join(self.bot_steps_directory, file_name)
                schedule_info = self._read_schedule_from_csv(file_path)

                schedule_str = "Not Set"
                status_str = "Idle"
                if schedule_info:
                    start_datetime_obj = QDateTime.fromString(schedule_info.get('start_datetime'), Qt.DateFormat.ISODate)
                    schedule_str = f"{schedule_info.get('repeat', 'Once')} at {start_datetime_obj.toString('yyyy-MM-dd hh:mm')}"
                    status_str = "Scheduled" if schedule_info.get("enabled") else "Disabled"

                tree_item = QTreeWidgetItem(self.saved_steps_tree, [bot_name, schedule_str, status_str])
            if not step_files:
                self.saved_steps_tree.addTopLevelItem(QTreeWidgetItem(["No saved bots found."]))
        except Exception as e:
            QMessageBox.critical(self, "Error Loading Saved Bots", f"Could not load bot files: {e}")

    # --- REPLACE load_steps_from_file ---
    def load_steps_from_file(self, file_path: str) -> None:
        self._internal_clear_all_steps()
        try:
            section = None
            loaded_variables, loaded_steps = {}, []
            with open(file_path, 'r', newline='', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                for row_num, row in enumerate(reader):
                    if not row: continue
                    
                    header = row[0]
                    if header == "__SCHEDULE_INFO__":
                        section = "SCHEDULE"
                        continue
                    elif header == "__GLOBAL_VARIABLES__":
                        section = "VARIABLES"
                        continue
                    elif header == "__BOT_STEPS__":
                        section = "STEPS"
                        next(reader, None)  # Skip the "StepType,DataJSON" header row
                        continue
                    
                    if section == "SCHEDULE":
                        # Schedule is handled separately, so we just skip these rows here
                        pass
                    elif section == "VARIABLES":
                        if len(row) == 2:
                            loaded_variables[row[0]] = json.loads(row[1])
                    elif section == "STEPS":
                        if len(row) == 2:
                            loaded_steps.append(json.loads(row[1]))

            self.global_variables = loaded_variables
            self._update_variables_list_display()
            self.added_steps_data = loaded_steps
            self._rebuild_execution_tree()
            self._log_to_console(f"Loaded bot from {os.path.basename(file_path)}")

        except Exception as e:
            QMessageBox.critical(self, "Load Error", f"An unexpected error occurred while loading the file:\n{e}")
            self._log_to_console(f"Load Error: {e}")


    # --- REPLACE save_bot_steps_dialog ---
    def save_bot_steps_dialog(self) -> None:
        if not self.added_steps_data and not self.global_variables:
            QMessageBox.information(self, "Nothing to Save", "The execution queue and global variables are empty.")
            return
        
        os.makedirs(self.bot_steps_directory, exist_ok=True)
        existing_bots = [os.path.splitext(f)[0] for f in os.listdir(self.bot_steps_directory) if f.endswith(".csv")]
        
        dialog = SaveBotDialog(existing_bots, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            bot_name = dialog.get_bot_name()
            if not bot_name: return
    
            file_path = os.path.join(self.bot_steps_directory, f"{bot_name}.csv")
            
            # Preserve existing schedule if overwriting
            existing_schedule = self._read_schedule_from_csv(file_path) if os.path.exists(file_path) else None

            try:
                variables_to_save = {k: v for k, v in self.global_variables.items() if isinstance(v, (str, int, float, bool, list, dict, type(None)))}

                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    
                    # Write schedule first
                    if existing_schedule:
                        writer.writerow(["__SCHEDULE_INFO__"])
                        writer.writerow([json.dumps(existing_schedule)])

                    # Write variables
                    writer.writerow(["__GLOBAL_VARIABLES__"])
                    for var_name, var_value in variables_to_save.items():
                        writer.writerow([var_name, json.dumps(var_value)])
                    
                    # Write steps
                    writer.writerow(["__BOT_STEPS__"])
                    writer.writerow(["StepType", "DataJSON"])
                    for step_data_dict in self.added_steps_data:
                        writer.writerow([step_data_dict["type"], json.dumps(step_data_dict)])
                
                QMessageBox.information(self, "Save Successful", f"Bot saved to:\n{file_path}")
                self.load_saved_steps_to_tree()
            except Exception as e:
                QMessageBox.critical(self, "Save Error", f"Failed to save bot steps:\n{e}")


    # --- HELPER METHODS for CSV schedule handling ---
    def _read_schedule_from_csv(self, file_path: str) -> Optional[Dict[str, Any]]:
        """Reads only the schedule info from a bot's CSV file."""
        if not os.path.exists(file_path):
            return None
        try:
            with open(file_path, 'r', newline='', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                for row in reader:
                    if row and row[0] == "__SCHEDULE_INFO__":
                        # The next row should contain the JSON data
                        schedule_row = next(reader, None)
                        if schedule_row:
                            return json.loads(schedule_row[0])
            return None
        except (Exception, json.JSONDecodeError):
            return None

    def _write_schedule_to_csv(self, file_path: str, schedule_data: Dict[str, Any]) -> bool:
        """Writes or updates the schedule info in a bot's CSV file."""
        lines = []
        schedule_written = False
        try:
            # Read all existing content
            if os.path.exists(file_path):
                with open(file_path, 'r', newline='', encoding='utf-8') as f:
                    lines = list(csv.reader(f))

            # Find and replace schedule, or add it if not present
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                # Check if schedule section already exists
                try:
                    schedule_header_index = [i for i, row in enumerate(lines) if row and row[0] == "__SCHEDULE_INFO__"][0]
                    lines[schedule_header_index + 1] = [json.dumps(schedule_data)]
                    schedule_written = True
                except IndexError:
                    # Header not found, so we'll add it at the top
                    pass
                
                if schedule_written:
                    writer.writerows(lines)
                else:
                    # Write new schedule at the top
                    writer.writerow(["__SCHEDULE_INFO__"])
                    writer.writerow([json.dumps(schedule_data)])
                    # Write the rest of the original content
                    writer.writerows(lines)
            return True
        except Exception as e:
            self._log_to_console(f"Error writing schedule to {file_path}: {e}")
            return False
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
