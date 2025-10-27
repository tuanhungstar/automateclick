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
import re
import ast
import urllib.request
import zipfile
import shutil
from PIL import ImageGrab
import PIL.Image
from PIL.ImageQt import ImageQt
from PyQt6.QtGui import QPixmap, QColor, QFont, QPainter, QPen, QIcon, QPolygonF, QCursor, QAction
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QVariant, QObject, QSize, QPoint, QRegularExpression,QRect,QDateTime, QTimer, QPointF
from PyQt6 import QtWidgets, QtGui, QtCore
from typing import Optional, List, Dict, Any, Tuple, Union
# Ensure my_lib is in the Python path
script_dir = os.path.dirname(os.path.abspath(__file__))
my_lib_dir = os.path.join(script_dir, "my_lib")
if my_lib_dir not in sys.path:
    sys.path.insert(0, my_lib_dir)

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QMainWindow,
    QComboBox, QListWidget, QLabel, QPushButton, QListWidgetItem,
    QMessageBox, QProgressBar, QFileDialog, QDialog,
    QLineEdit, QVBoxLayout as QVBoxLayoutDialog, QFormLayout,
    QDialogButtonBox,
    QRadioButton, QGroupBox, QCheckBox, QTextEdit,
    QTreeWidget, QTreeWidgetItem, QGridLayout, QHeaderView, QSplitter, QInputDialog,
    QStackedLayout, QBoxLayout,QMenu,QPlainTextEdit,QSizePolicy, QTextBrowser,QDateTimeEdit,QTreeWidgetItemIterator,
    QScrollArea, QTabWidget, QFrame, QMenu, 
)

from PyQt6.QtGui import QIntValidator
from PyQt6.QtGui import QSyntaxHighlighter, QTextCharFormat, QColor, QFont, QPainter, QPen, QTextCursor,QTextFormat,QKeyEvent,QTextDocument

# Use the actual libraries from the my_lib folder
from my_lib.shared_context import ExecutionContext, GuiCommunicator
from my_lib.BOT_take_image import MainWindow as BotTakeImageWindow



class SecondWindow(QtWidgets.QDialog): # Or QtWidgets.QMainWindow if you prefer a full window

    screenshot_saved = pyqtSignal(str)

    def __init__(self, image: str, base_dir: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Take and Manage Screenshots")

        # Use the 'base_dir' argument directly
        icon_path = os.path.join(base_dir, "app_icon.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
                
        self.setMinimumSize(700, 300) # Set a minimum size to match the original GUI

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
    # NEW SIGNALS FOR DRAG AND DROP
    step_drag_started = pyqtSignal(dict, int)
    step_reorder_requested = pyqtSignal(int, int)

    def __init__(self, step_data: Dict[str, Any], step_number: int, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.step_data = step_data
        self.step_number = step_number
        
        # Drag and drop attributes
        self.dragging = False
        self.drag_start_position = QPoint()
        self.original_index = step_data.get("original_listbox_row_index", -1)
        self.drag_widget = None  # For visual feedback during drag
        
        self.init_ui()
        
        # Enable drag and drop
        self.setAcceptDrops(True)

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(5)
        self.setObjectName("ExecutionStepCard")
        self.set_status("#dcdcdc")

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

        card_button_style = """
            QPushButton {
                background-color: #f0f0f0;
                color: #333333;
                border: 1px solid #c0c0c0;
                border-radius: 4px;
                padding: 4px 8px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
                border-color: #a0a0a0;
            }
            QPushButton:pressed {
                background-color: #d0d0d0;
            }
        """
        
        for button in [self.up_button, self.down_button, self.edit_button, self.delete_button, self.save_template_button, self.execute_this_button]:
            button.setStyleSheet(card_button_style)

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

        self._original_method_text = self._get_formatted_method_name()
        if self._original_method_text:
            self.method_label = QLabel(self._original_method_text)
            self.method_label.setStyleSheet("font-size: 10pt; padding: 5px; background-color: white; border: 1px solid #E0E0E0; border-radius: 3px;")
            self.method_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            main_layout.addWidget(self.method_label)
        else:
            self.method_label = None

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
                    elif config.get('type') == 'hardcoded_file':
                        full_path_value = config['value']
                        base_name = os.path.basename(full_path_value)
                        value_str = f"File: '{base_name}'"
                    elif config.get('type') == 'variable': value_str = f"Variable: @{config['value']}"
                    param_label = QLabel(f"{param_name}:")
                    value_label = QLineEdit(value_str)
                    value_label.setReadOnly(True)
                    value_label.setStyleSheet("background-color: #FFFFFF; font-size: 9pt; padding: 2px; border: 1px solid #D3D3D3;")
                    params_layout.addRow(param_label, value_label)
                params_group.setLayout(params_layout)
                main_layout.addWidget(params_group)

    # FIXED DRAG AND DROP METHODS
    def mousePressEvent(self, event: QtGui.QMouseEvent):
        if event.button() == Qt.MouseButton.LeftButton:
            self.drag_start_position = event.pos()
            self.dragging = False
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QtGui.QMoveEvent):
        if not (event.buttons() & Qt.MouseButton.LeftButton):
            return
        
        if not self.dragging:
            # Check if we've moved far enough to start dragging
            if ((event.pos() - self.drag_start_position).manhattanLength() >= 
                QApplication.startDragDistance()):
                self.start_drag()
        super().mouseMoveEvent(event)

    def start_drag(self):
        """Start the drag operation with proper Qt drag and drop."""
        self.dragging = True
        
        # Create drag object
        drag = QtGui.QDrag(self)
        mime_data = QtCore.QMimeData()
        
        # Store the source index in mime data
        mime_data.setText(str(self.original_index))
        drag.setMimeData(mime_data)
        
        # Create drag pixmap for visual feedback
        pixmap = self.grab()
        drag.setPixmap(pixmap)
        drag.setHotSpot(self.drag_start_position)
        
        # Visual feedback on source
        self.setStyleSheet("""
            #ExecutionStepCard { 
                background-color: #E3F2FD; 
                border: 3px dashed #2196F3; 
                border-radius: 6px; 
                opacity: 0.7;
            }
        """)
        
        # Emit signal
        self.step_drag_started.emit(self.step_data, self.original_index)
        
        # Execute drag
        drop_action = drag.exec(Qt.DropAction.MoveAction)
        
        # Reset visual feedback
        self.set_status("#dcdcdc")
        self.dragging = False

    def mouseReleaseEvent(self, event: QtGui.QMouseEvent):
        if self.dragging:
            self.dragging = False
            self.set_status("#dcdcdc")
        super().mouseReleaseEvent(event)

    # Drag and drop event handlers
    def dragEnterEvent(self, event: QtGui.QDragEnterEvent):
        if event.mimeData().hasText():
            event.acceptProposedAction()
            # Visual feedback for drop target
            self.setStyleSheet("""
                #ExecutionStepCard { 
                    background-color: #FFF3E0; 
                    border: 2px solid #FF9800; 
                    border-radius: 6px; 
                }
            """)

    def dragMoveEvent(self, event: QtGui.QDragMoveEvent):
        if event.mimeData().hasText():
            event.acceptProposedAction()

    def dragLeaveEvent(self, event: QtGui.QDragLeaveEvent):
        # Reset visual feedback
        self.set_status("#dcdcdc")

    def dropEvent(self, event: QtGui.QDropEvent):
        # Reset visual feedback
        self.set_status("#dcdcdc")
        
        if event.mimeData().hasText():
            source_index = int(event.mimeData().text())
            target_index = self.original_index
            
            # Determine if dropping above or below based on position
            drop_pos = event.position() if hasattr(event, 'position') else event.pos()
            if drop_pos.y() < self.height() / 2:
                # Drop above (before)
                target_index = self.original_index
            else:
                # Drop below (after)
                target_index = self.original_index + 1
            
            if source_index != target_index:
                self.step_reorder_requested.emit(source_index, target_index)
            
            event.acceptProposedAction()

    # Keep all existing methods (unchanged)
    def _get_formatted_title(self) -> str:
        step_type = self.step_data.get("type", "Unknown")
        if step_type == "group_start":
            return f"Group: {self.step_data.get('group_name', 'Unnamed')}"
        
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

    def set_status(self, border_color: str, is_running: bool = False, is_error: bool = False):
        if is_running:
            border_color = "#FFD700"
            background_color = "#FFFBF0"
            border_thickness = 4
        elif is_error:
            border_color = "#DC3545"
            background_color = "#FFF5F5"
            border_thickness = 4
        elif border_color == "darkGreen" or border_color == "#28a745":
            border_color = "#28a745"
            background_color = "#F8FFF8"
            border_thickness = 2
        else:
            background_color = "#F8F8F8"
            border_thickness = 2
        
        self.setStyleSheet(f"""
            #ExecutionStepCard {{ 
                background-color: {background_color}; 
                border: {border_thickness}px solid {border_color}; 
                border-radius: 6px; 
            }}
        """)

    def set_result_text(self, result_message: str):
        #print(f"DEBUG: set_result_text called with message: '{result_message}'")
        #print(f"DEBUG: method_label exists: {self.method_label is not None}")
        if not self.method_label:
            #print("DEBUG: method_label is None!")  # ADD THIS
            return

        if len(result_message) > 300:
             result_message = result_message[:297] + "..."

        assign_to_var = self.step_data.get("assign_to_variable_name")
        
        if self.step_data.get("type") == "step" and assign_to_var and "Result: " in result_message:
            try:
                if " (Assigned to @" in result_message:
                    result_val_str = result_message.split("Result: ")[1].split(" (Assigned to @")[0]
                elif " (Assigned to" in result_message:
                    result_val_str = result_message.split("Result: ")[1].split(" (Assigned to")[0]
                else:
                    result_val_str = result_message.split("Result: ")[1]
            except IndexError:
                result_val_str = "Error parsing result"
            
            display_text = f"@{assign_to_var} = {result_val_str}"
            self.method_label.setText(display_text)
            self.method_label.setStyleSheet("font-size: 10pt; font-style: italic; color: #155724; padding: 5px; background-color: #d4edda; border: 1px solid #c3e6cb; border-radius: 3px;")
        
        else:
            self.method_label.setText(f"✓ {result_message}")
            self.method_label.setStyleSheet("font-size: 10pt; font-style: italic; color: #155724; padding: 5px; background-color: #d4edda; border: 1px solid #c3e6cb; border-radius: 3px;")

    def clear_result(self):
        if self.method_label:
            self.method_label.setText(self._original_method_text)
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
        
        # --- MODIFICATION: Store all paths and derive folder list ---
        self._all_image_relative_paths = image_filenames 
        # Get folder names (e.g., "FolderA") from paths ("FolderA/image1")
        self._folder_list = sorted(list(set([path.split('/')[0] for path in self._all_image_relative_paths if '/' in path])))
        # --- END MODIFICATION ---

        self.param_editors: Dict[str, QLineEdit] = {}
        self.param_var_selectors: Dict[str, QComboBox] = {}
        self.param_value_source_combos: Dict[str, QComboBox] = {}
        self.file_selector_combos: Dict[str, QComboBox] = {}
        # --- NEW: Add dict for new folder combos ---
        self.folder_selector_combos: Dict[str, QComboBox] = {}
        # --- END NEW ---
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

            # --- MODIFICATION: Changed "current_file" to "image_to_click" ---
            if param_name == "image_to_click":
                file_source_combo = QComboBox()
                file_source_combo.addItems(["Select from Files", "Global Variable"])
                self.param_value_source_combos[param_name] = file_source_combo

                # --- NEW: Create and populate the folder dropdown ---
                folder_selector_combo = QComboBox()
                folder_selector_combo.addItem("All Folders")
                folder_selector_combo.addItems(self._folder_list)
                self.folder_selector_combos[param_name] = folder_selector_combo
                # --- END NEW ---

                file_selector_combo = QComboBox()
                file_selector_combo.addItem("-- Select File --")
                # --- MODIFICATION: Populate with all files initially ---
                file_selector_combo.addItems(self._all_image_relative_paths)
                file_selector_combo.addItem("--- Add new image ---")
                self.file_selector_combos[param_name] = file_selector_combo

                variable_select_combo = QComboBox()
                variable_select_combo.addItem("-- Select Variable --")
                variable_select_combo.addItems(current_global_var_names)
                self.param_var_selectors[param_name] = variable_select_combo

                # --- MODIFICATION: Update connect to handle new folder combo ---
                file_source_combo.currentIndexChanged.connect(
                    lambda index, f_fold=folder_selector_combo, f_sel=file_selector_combo, v_sel=variable_select_combo: \
                    self._toggle_file_or_var_input(index, f_fold, f_sel, v_sel)
                )
                # --- END MODIFICATION ---

                # --- NEW: Connect folder combo to filter method ---
                folder_selector_combo.currentIndexChanged.connect(
                    self._update_file_list_based_on_folder
                )
                # --- END NEW ---
                
                file_selector_combo.currentIndexChanged.connect(lambda index, p_name=param_name: self._on_file_selection_changed(p_name, index))
                self.update_image_filenames.connect(self._refresh_file_selector_combo)

                param_h_layout.addWidget(file_source_combo, 0)
                # --- NEW: Add folder combo to layout ---
                param_h_layout.addWidget(folder_selector_combo, 1)
                # --- END NEW ---
                param_h_layout.addWidget(file_selector_combo, 2)
                param_h_layout.addWidget(variable_select_combo, 1)

                if param_name in initial_parameters_config:
                    config = initial_parameters_config[param_name]
                    if config.get('type') == 'hardcoded_file':
                        file_source_combo.setCurrentIndex(0)
                        
                        # --- MODIFICATION: Set folder and file dropdowns from saved path ---
                        saved_path = config['value']
                        saved_folder = "All Folders"
                        if '/' in saved_path:
                            saved_folder = saved_path.split('/')[0]
                        
                        folder_idx = folder_selector_combo.findText(saved_folder)
                        if folder_idx != -1:
                            folder_selector_combo.setCurrentIndex(folder_idx)
                            # Manually trigger filter to populate file list
                            self._update_file_list_based_on_folder(folder_idx) 
                        
                        idx = file_selector_combo.findText(saved_path)
                        if idx != -1:
                            file_selector_combo.setCurrentIndex(idx)
                            self._on_file_selection_changed(param_name, idx)
                        # --- END MODIFICATION ---

                    elif config.get('type') == 'variable':
                        file_source_combo.setCurrentIndex(1)
                        idx = variable_select_combo.findText(config['value'])
                        if idx != -1:
                            variable_select_combo.setCurrentIndex(idx)
                
                # --- MODIFICATION: Pass folder combo to toggle function ---
                self._toggle_file_or_var_input(file_source_combo.currentIndex(), folder_selector_combo, file_selector_combo, variable_select_combo)
                # --- END MODIFICATION ---
            
            else: # Logic for all other parameters (unchanged)
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
        
        # --- Assignment Group Box (Unchanged) ---
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
        # --- End Assignment Group Box ---

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

    # --- MODIFICATION: Added folder_selector_combo ---
    def _toggle_file_or_var_input(self, index: int, folder_selector_combo: QComboBox, file_selector_combo: QComboBox, variable_select_combo: QComboBox) -> None:
        is_file_selection = (index == 0)
        folder_selector_combo.setVisible(is_file_selection) # Show/hide folder combo
        file_selector_combo.setVisible(is_file_selection)
        variable_select_combo.setVisible(not is_file_selection)
        if not is_file_selection:
            folder_selector_combo.setCurrentIndex(0) # Reset folder
            file_selector_combo.setCurrentIndex(0)
    # --- END MODIFICATION ---

    # --- NEW: Method to filter file list based on folder selection ---
    def _update_file_list_based_on_folder(self, index: int = -1):
        """Filters the file selector combo based on the folder selector combo."""
        folder_combo = self.sender()
        if not isinstance(folder_combo, QComboBox):
            # Fallback if called manually (e.g., during init)
            folder_combos = self.folder_selector_combos.values()
            if not folder_combos:
                return
            folder_combo = list(folder_combos)[0]

        param_name = ""
        for name, combo in self.folder_selector_combos.items():
            if combo is folder_combo:
                param_name = name
                break
        
        if not param_name:
            return

        file_combo = self.file_selector_combos.get(param_name)
        if not file_combo:
            return

        selected_folder = folder_combo.currentText()
        current_file = file_combo.currentText() # Save current selection

        file_combo.blockSignals(True)
        file_combo.clear()
        file_combo.addItem("-- Select File --")

        if selected_folder == "All Folders":
            filtered_files = self._all_image_relative_paths
        else:
            # Filter by path prefix
            prefix = selected_folder + '/'
            # Also include root files (those without any '/')
            root_files = [path for path in self._all_image_relative_paths if '/' not in path]
            folder_files = [path for path in self._all_image_relative_paths if path.startswith(prefix)]
            filtered_files = root_files + folder_files

        file_combo.addItems(sorted(filtered_files))
        file_combo.addItem("--- Add new image ---")
        
        idx = file_combo.findText(current_file) # Try to restore selection
        if idx != -1:
            file_combo.setCurrentIndex(idx)
        
        file_combo.blockSignals(False)
    # --- END NEW ---

    def _on_file_selection_changed(self, param_name: str, index: int) -> None:
        # --- MODIFICATION: Changed "current_file" to "image_to_click" ---
        if param_name == "image_to_click" and param_name in self.file_selector_combos:
            file_selector_combo = self.file_selector_combos[param_name]
            selected_text = file_selector_combo.currentText()
            if selected_text == "--- Add new image ---":
                file_selector_combo.blockSignals(True)
                file_selector_combo.setCurrentIndex(0)
                file_selector_combo.blockSignals(False)
                self.request_screenshot.emit()
            elif selected_text != "-- Select File --":
                # selected_text is now the relative path (e.g., "Folder/image")
                self.gui_communicator.update_module_info_signal.emit(selected_text)
            else:
                self.gui_communicator.update_module_info_signal.emit("")

    def _refresh_file_selector_combo(self, new_filenames: List[str], saved_filename: str = "") -> None:
        # --- MODIFICATION: This method is now much more complex ---
        # 1. Update master list
        self._all_image_relative_paths = new_filenames
        # 2. Update folder list
        self._folder_list = sorted(list(set([path.split('/')[0] for path in new_filenames if '/' in path])))
        
        for param_name, combo_box in self.file_selector_combos.items():
            # 3. Check for the correct parameter name
            if param_name == "image_to_click":
                folder_combo = self.folder_selector_combos.get(param_name)
                
                # 4. Determine which folder should be selected
                target_folder_name = "All Folders"
                if saved_filename and '/' in saved_filename:
                    target_folder_name = saved_filename.split('/')[0]
                elif folder_combo: # Keep current folder if no new file is specified
                    target_folder_name = folder_combo.currentText()

                # 5. Re-populate folder combo
                if folder_combo:
                    folder_combo.blockSignals(True)
                    folder_combo.clear()
                    folder_combo.addItem("All Folders")
                    folder_combo.addItems(self._folder_list)
                    
                    folder_idx = folder_combo.findText(target_folder_name)
                    if folder_idx != -1:
                        folder_combo.setCurrentIndex(folder_idx)
                    else:
                        folder_combo.setCurrentIndex(0) # Default to All Folders
                    folder_combo.blockSignals(False)

                # 6. Get the *actual* selected folder (might be "All Folders")
                selected_folder = folder_combo.currentText() if folder_combo else "All Folders"

                # 7. Filter file list based on selected folder
                if selected_folder == "All Folders":
                    filtered_files = self._all_image_relative_paths
                else:
                    # Filter by path prefix
                    prefix = selected_folder + '/'
                    # Also include root files (those without any '/')
                    root_files = [path for path in self._all_image_relative_paths if '/' not in path]
                    folder_files = [path for path in self._all_image_relative_paths if path.startswith(prefix)]
                    filtered_files = root_files + folder_files


                # 8. Re-populate file combo
                combo_box.blockSignals(True)
                combo_box.clear()
                combo_box.addItem("-- Select File --")
                combo_box.addItems(sorted(filtered_files))
                combo_box.addItem("--- Add new image ---")

                # 9. Try to select the newly saved file
                if saved_filename:
                    idx = combo_box.findText(saved_filename)
                    if idx != -1:
                        combo_box.setCurrentIndex(idx)
                        self.gui_communicator.update_module_info_signal.emit(saved_filename)
                
                combo_box.blockSignals(False)
                
                # 10. Emit signal if no file is selected
                if combo_box.currentIndex() <= 0 and saved_filename == "":
                    self.gui_communicator.update_module_info_signal.emit("")
        # --- END MODIFICATION ---

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
            # --- MODIFICATION: Changed "current_file" to "image_to_click" ---
            if param_name == "image_to_click" and param_name in self.file_selector_combos:
                value_source_index = source_combo.currentIndex()
                if value_source_index == 0:
                    # The file name is now a relative path, e.g., "Folder/image"
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
            # --- END MODIFICATION ---

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
    loop_iteration_started = pyqtSignal(str, int)  # loop_id, iteration_number

    def __init__(self, steps_to_execute: List[Dict[str, Any]], module_directory: str, gui_communicator: GuiCommunicator,
                 global_variables_ref: Dict[str, Any], 
                 wait_config: Dict[str, Any],  # <-- ARGUMENT IS NOW A DICT
                 parent: Optional[QWidget] = None,
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
        self.click_image_dir = os.path.normpath(os.path.join(module_directory, "..", "Click_image"))
        self._is_paused = False # ADD THIS LINE
        self.wait_config = wait_config # <-- STORE THE CONFIG DICT
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
        if not self.steps_to_execute: 
            self.execution_finished_all.emit(self.context, False, -1)
            return
            
        total_steps_for_progress = len(self.steps_to_execute) * 2 if not self.single_step_mode else 1
        if total_steps_for_progress == 0: 
            total_steps_for_progress = 1
            
        current_execution_item_count = 0
        original_sys_path = sys.path[:]
        
        if self.module_directory not in sys.path: 
            sys.path.insert(0, self.module_directory)
            
        self.context.set_click_image_base_dir(self.click_image_dir)    
        self.loop_stack = []
        self.conditional_stack = []
        step_index = self.selected_start_index if self.single_step_mode else 0
        original_listbox_row_index = 0
        
        try:
            while step_index < len(self.steps_to_execute):
                
                # --- ADD THIS BLOCK for PAUSE logic ---
                while self._is_paused:
                    if self._is_stopped: # Allow stop to break out of pause
                        break
                    QThread.msleep(100) # Sleep while paused
                # --- END ADD ---                
                
                if self._is_stopped: 
                    break
                if self.single_step_mode and current_execution_item_count >= 1: 
                    break
                
                # --- REPLACE THE OLD WAIT LOGIC WITH THIS NEW, DYNAMIC LOGIC ---
                if current_execution_item_count > 0 and not self.single_step_mode:
                    wait_seconds = 0
                    
                    # 1. Resolve the wait time from the config
                    if self.wait_config['type'] == 'variable':
                        var_name = self.wait_config['value']
                        # Get the current value from the shared global variables dict
                        var_value = self.global_variables.get(var_name, 0)
                        try:
                            # Ensure the value from the variable is a valid number
                            wait_seconds = float(var_value)
                        except (ValueError, TypeError):
                            self.context.add_log(f"Warning: Global variable '@{var_name}' does not contain a valid number for wait time. Using 0.")
                            wait_seconds = 0
                    else: # It's hardcoded
                        wait_seconds = self.wait_config['value']

                    # 2. Perform the wait if necessary
                    if wait_seconds > 0:
                        self.context.add_log(f"Waiting for {wait_seconds:.2f} second(s)...")
                        # QThread.msleep() takes milliseconds, so we multiply by 1000
                        QThread.msleep(int(wait_seconds * 1000))
                # --- END OF REPLACEMENT ---







                
                step_data = self.steps_to_execute[step_index]
                step_type = step_data["type"]
                original_listbox_row_index = step_data.get("original_listbox_row_index", step_index)
                current_execution_item_count += 1
                progress_percentage = int((current_execution_item_count / total_steps_for_progress) * 100)
                self.execution_progress.emit(min(progress_percentage, 100))
                
                # Check for conditional skipping logic
                is_skipping = False
                if self.conditional_stack:
                    current_if = self.conditional_stack[-1]
                    if not current_if.get('condition_result', True) and not current_if.get('else_taken', False):
                        if not (step_type == "ELSE" and step_data.get("if_id") == current_if["if_id"]): 
                            is_skipping = True
                    elif current_if.get('condition_result', False) and current_if.get('else_taken', False):
                        if not (step_type == "IF_END" and step_data.get("if_id") == current_if["if_id"]): 
                            is_skipping = True
                
                if is_skipping:
                    self.execution_item_started.emit(step_data, original_listbox_row_index)
                    if step_type == "IF_START": 
                        self.conditional_stack.append({'if_id': step_data['if_id'], 'skipped_marker': True})
                    elif step_type == "loop_start": 
                        self.loop_stack.append({'loop_id': step_data['loop_id'], 'skipped_marker': True})
                    elif step_type == "IF_END":
                        if self.conditional_stack and self.conditional_stack[-1].get('skipped_marker'): 
                            self.conditional_stack.pop()
                    elif step_type == "loop_end":
                        if self.loop_stack and self.loop_stack[-1].get('skipped_marker'): 
                            self.loop_stack.pop()
                    self.execution_item_finished.emit(step_data, "SKIPPED", original_listbox_row_index)
                    step_index += 1
                    continue
                
                self.execution_item_started.emit(step_data, original_listbox_row_index)
                QThread.msleep(50)

                if step_type in ["group_start", "group_end"]:
                    self.execution_item_finished.emit(step_data, "Organizational Step", original_listbox_row_index)
                    step_index += 1
                    continue

                # ENHANCED LOOP HANDLING WITH RESET LOGIC
                elif step_type == "loop_start":
                    loop_id, loop_config = step_data["loop_id"], step_data["loop_config"]
                    is_new_loop = not (self.loop_stack and self.loop_stack[-1].get('loop_id') == loop_id)
                    
                    if is_new_loop:
                        # First time entering this loop
                        loop_end_index, nesting_level = -1, 0
                        for i in range(step_index + 1, len(self.steps_to_execute)):
                            s, s_type = self.steps_to_execute[i], self.steps_to_execute[i].get("type")
                            if s_type in ["loop_start", "IF_START"]: 
                                nesting_level += 1
                            elif s_type == "loop_end" and s.get("loop_id") == loop_id and nesting_level == 0: 
                                loop_end_index = i
                                break
                            elif s_type in ["loop_end", "IF_END"] and nesting_level > 0: 
                                nesting_level -= 1
                                
                        if loop_end_index == -1: 
                            raise ValueError(f"Loop '{loop_id}' has no matching 'loop_end' marker.")
                            
                        total_iterations = self._resolve_loop_count(loop_config)
                        current_loop_info = {
                            'loop_id': loop_id, 
                            'start_index': step_index, 
                            'end_index': loop_end_index, 
                            'current_iteration': 1, 
                            'total_iterations': total_iterations, 
                            'loop_config': loop_config
                        }
                        self.loop_stack.append(current_loop_info)
                        
                        # Emit signal for first iteration (no reset needed for first iteration)
                        self.loop_iteration_started.emit(loop_id, 1)
                        
                    else:
                        # Subsequent iterations - RESET LOOP STEPS HERE
                        current_loop_info = self.loop_stack[-1]
                        current_loop_info['current_iteration'] += 1
                        current_loop_info['total_iterations'] = self._resolve_loop_count(loop_config)
                        
                        # Emit signal to reset loop steps before continuing
                        iteration_num = current_loop_info['current_iteration']
                        self.loop_iteration_started.emit(loop_id, iteration_num)

                    current_loop_info = self.loop_stack[-1]
                    if current_loop_info['current_iteration'] > current_loop_info['total_iterations']: 
                        self.loop_stack.pop()
                        step_index = current_loop_info['end_index']
                        self.execution_item_finished.emit(step_data, "Loop Finished", original_listbox_row_index)
                    else:
                        assign_var = loop_config.get("assign_iteration_to_variable")
                        if assign_var: 
                            self.global_variables[assign_var] = current_loop_info['current_iteration']
                            self.context.add_log(f"Assigned iteration {current_loop_info['current_iteration']} to @{assign_var}")
                        self.execution_item_finished.emit(step_data, f"Iter {current_loop_info['current_iteration']}/{current_loop_info['total_iterations']}", original_listbox_row_index)

                elif step_type == "loop_end":
                    if not self.loop_stack or self.loop_stack[-1].get('loop_id') != step_data['loop_id']: 
                        raise ValueError(f"Mismatched loop_end for ID: {step_data['loop_id']}")
                    step_index = self.loop_stack[-1]['start_index'] - 1
                    self.execution_item_finished.emit(step_data, "Looping...", original_listbox_row_index)

                # ... (keep all other step type handling logic the same) ...
                
                elif step_type == "step":
                    class_name, method_name, module_name = step_data["class_name"], step_data["method_name"], step_data["module_name"]
                    parameters_config = step_data["parameters_config"]
                    assign_to_variable_name = step_data["assign_to_variable_name"]
                    resolved_parameters, params_str_debug = {}, []
                    
                    for param_name, config in parameters_config.items():
                        if param_name == "original_listbox_row_index": 
                            continue
                        if config['type'] == 'hardcoded': 
                            resolved_parameters[param_name] = config['value']
                            params_str_debug.append(f"{param_name}={repr(config['value'])}")
                        elif config['type'] == 'hardcoded_file': 
                            resolved_parameters[param_name] = config['value']
                            params_str_debug.append(f"{param_name}=FILE('{config['value']}')")
                        elif config['type'] == 'variable':
                            var_name = config['value']
                            if var_name in self.global_variables: 
                                resolved_parameters[param_name] = self.global_variables[var_name]
                                params_str_debug.append(f"{param_name}=@{var_name}({repr(self.global_variables[var_name])})")
                            else: 
                                raise ValueError(f"Global variable '{var_name}' not found for parameter '{param_name}'.")
                    
                    self.context.add_log(f"Executing: {class_name}.{method_name}({', '.join(params_str_debug)})")
                    
                    try:
                        module = importlib.import_module(module_name)
                        importlib.reload(module)
                        class_obj = getattr(module, class_name)
                        instance_key = (class_name, module_name)
                        
                        if instance_key not in self.instantiated_objects:
                            init_kwargs = {}
                            if 'context' in inspect.signature(class_obj.__init__).parameters: 
                                init_kwargs['context'] = self.context
                            self.instantiated_objects[instance_key] = class_obj(**init_kwargs)
                        
                        instance = self.instantiated_objects[instance_key]
                        method_func = getattr(instance, method_name)
                        method_kwargs = {k:v for k,v in resolved_parameters.items()}
                        
                        if 'context' in inspect.signature(method_func).parameters: 
                            method_kwargs['context'] = self.context
                            
                        result = method_func(**method_kwargs)
                        
                        if assign_to_variable_name: 
                            self.global_variables[assign_to_variable_name] = result
                            result_msg = f"Result: {result} (Assigned to @{assign_to_variable_name})"
                        else: 
                            result_msg = f"Result: {result}"
                            
                        self.execution_item_finished.emit(step_data, result_msg, original_listbox_row_index)
                        
                    except Exception as e: 
                        raise e

                elif step_type == "IF_START":
                    condition_config = step_data["condition_config"]
                    if_id = step_data["if_id"]
                    condition_result = self._evaluate_condition(condition_config["condition"])
                    self.context.add_log(f"IF '{if_id}' evaluated: {condition_result}")
                    self.conditional_stack.append({'if_id': if_id, 'condition_result': condition_result, 'else_taken': False})
                    self.execution_item_finished.emit(step_data, f"Condition: {condition_result}", original_listbox_row_index)

                elif step_type == "ELSE":
                    if not self.conditional_stack or self.conditional_stack[-1].get('if_id') != step_data['if_id']: 
                        raise ValueError(f"Mismatched ELSE for ID: {step_data['if_id']}")
                    current_if = self.conditional_stack[-1]
                    current_if['else_taken'] = True
                    self.execution_item_finished.emit(step_data, "Branching", original_listbox_row_index)
                    
                    if current_if.get('condition_result', False):
                        nesting_level = 0
                        for i in range(step_index + 1, len(self.steps_to_execute)):
                            s, s_type = self.steps_to_execute[i], self.steps_to_execute[i].get("type")
                            if s_type in ["loop_start", "IF_START"]: 
                                nesting_level += 1
                            elif s_type == "IF_END" and s.get("if_id") == current_if["if_id"] and nesting_level == 0: 
                                step_index = i-1
                                break
                            elif s_type in ["loop_end", "IF_END"] and nesting_level > 0: 
                                nesting_level -= 1

                elif step_type == "IF_END":
                    if not self.conditional_stack or self.conditional_stack[-1].get('if_id') != step_data['if_id']: 
                        raise ValueError(f"Mismatched IF_END for ID: {step_data['if_id']}")
                    self.conditional_stack.pop()
                    self.execution_item_finished.emit(step_data, "End of Conditional", original_listbox_row_index)

                step_index += 1
                
        except Exception as e:
            error_msg = f"Error at step {original_listbox_row_index+1}: {type(e).__name__}: {e}"
            self.context.add_log(error_msg)
            self.execution_error.emit(self.steps_to_execute[step_index] if step_index < len(self.steps_to_execute) else {}, error_msg, original_listbox_row_index)
            self._is_stopped = True
            
        finally:
            sys.path = original_sys_path
            if self.single_step_mode:
                next_index = -1
                if not self._is_stopped and step_index < len(self.steps_to_execute): 
                    next_index = self.steps_to_execute[step_index].get("original_listbox_row_index", -1)
                self.next_step_index_to_select = next_index
            self.execution_finished_all.emit(self.context, self._is_stopped, self.next_step_index_to_select)

# Add these new methods anywhere inside the ExecutionWorker class
    def pause(self):
        """Sets the pause flag and logs the action."""
        self._is_paused = True
        self.context.add_log("Execution PAUSED.")
        self.execution_started.emit("Execution PAUSED.") # Re-use signal for logging

    def resume(self):
        """Clears the pause flag and logs the action."""
        self._is_paused = False
        self.context.add_log("Execution RESUMED.")
        self.execution_started.emit("Execution RESUMED.") # Re-use signal for logging
# --- NEW WIDGET: RearrangeStepItemWidget ---
# --- REPLACEMENT WIDGET: RearrangeStepItemWidget (WITH GUARANTEED STEP NUMBER FIRST) ---
class RearrangeStepItemWidget(QWidget):
    """A compact widget to display a step's details for the rearrange dialog."""
    def __init__(self, step_data: Dict[str, Any], step_number: int, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.step_data = step_data
        self.step_number = step_number

        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 3, 5, 3)

        # Get the detailed text for the label
        full_text = self._get_formatted_display_text()
        
        self.label = QLabel(full_text)
        self.label.setWordWrap(True)
        
        # Add an icon to indicate draggability
        self.drag_handle_label = QLabel("☰")
        self.drag_handle_label.setToolTip("Drag to reorder")
        self.drag_handle_label.setStyleSheet("font-weight: bold; color: #7f8c8d;")
        self.drag_handle_label.setFixedWidth(20)
        layout.addWidget(self.drag_handle_label)
        layout.addWidget(self.label)
        
        # Set background color based on step type for better readability
        step_type = self.step_data.get("type")
        if step_type in ["loop_start", "IF_START", "group_start"]:
            self.setStyleSheet("background-color: #e8f4fd; border-radius: 3px;")
        elif step_type in ["loop_end", "IF_END", "group_end", "ELSE"]:
            self.setStyleSheet("background-color: #f1f2f3; border-radius: 3px;")
        else:
             self.setStyleSheet("background-color: #ffffff; border-radius: 3px;")

    def _get_formatted_display_text(self) -> str:
        """
        Creates a descriptive string for the step, ensuring the step number is always first.
        Format: Step X: @var = method(params...)
        """
        step_type = self.step_data.get("type")
        step_num = self.step_number

        # For non-step types (loops, ifs, etc.), use the reliable formatting from ExecutionStepCard
        if step_type != "step":
            temp_card = ExecutionStepCard(self.step_data, step_num)
            base_title = temp_card.step_label.text()
            details_text = temp_card.method_label.text() if temp_card.method_label else ""
            if details_text:
                return f"{base_title}: {details_text}"
            return base_title

        # --- THIS IS THE UPDATED LOGIC FOR 'step' TYPE ---
        
        # 1. Start with the Step number, which is always present.
        display_parts = [f"Step {step_num}:"]
        
        # 2. Add the variable assignment, if it exists.
        assign_var = self.step_data.get("assign_to_variable_name")
        if assign_var:
            display_parts.append(f"@{assign_var} =")
        
        # 3. Construct the method call string with parameters.
        method_name = self.step_data.get("method_name", "Unknown")
        params_config = self.step_data.get("parameters_config", {})
        param_details = []

        for name, config in params_config.items():
            if name == "original_listbox_row_index":
                continue
            
            value_display = ""
            param_type = config.get('type')
            
            if param_type == 'hardcoded':
                value_display = repr(config['value'])
            elif param_type == 'hardcoded_file':
                base_name = os.path.basename(config['value'])
                value_display = f"File:'{base_name}'"
            elif param_type == 'variable':
                value_display = f"@{config['value']}"
            else:
                value_display = "???"

            if len(value_display) > 25:
                value_display = value_display[:22] + "..."

            param_details.append(f"{name}={value_display}")
        
        param_str = ", ".join(param_details)
        method_call_str = f"{method_name}({param_str})"
        
        # 4. Add the method call string to our list of parts.
        display_parts.append(method_call_str)
        
        # 5. Join all the parts together with spaces.
        return " ".join(display_parts)
# --- NEW DIALOG: RearrangeStepsDialog ---

class RearrangeStepsDialog(QDialog):
    """A dialog to reorder steps using drag-and-drop."""
    def __init__(self, steps_data: List[Dict[str, Any]], parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Rearrange Bot Steps")
        self.setMinimumSize(800, 700) # Give it a good default size
        
        self.original_steps = steps_data
        
        main_layout = QVBoxLayout(self)

        # Instructions
        info_label = QLabel("Drag and drop the steps below to change their execution order.")
        info_label.setStyleSheet("font-style: italic; color: #555;")
        main_layout.addWidget(info_label)

        # The list widget that will hold the steps
        self.steps_list_widget = QListWidget()
        
        # --- Critical part: Enable Drag and Drop ---
        self.steps_list_widget.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        self.steps_list_widget.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.steps_list_widget.setAlternatingRowColors(True)
        self.steps_list_widget.setStyleSheet("font-size: 12pt;")
        main_layout.addWidget(self.steps_list_widget)

        # Populate the list with our custom widgets
        self.populate_list()
        
        # OK and Cancel buttons
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        main_layout.addWidget(button_box)

    def populate_list(self):
        """Fills the QListWidget with the steps."""
        self.steps_list_widget.clear()
        for i, step_data in enumerate(self.original_steps):
            list_item = QListWidgetItem(self.steps_list_widget)
            
            # Use our custom widget for the display
            item_widget = RearrangeStepItemWidget(step_data, i + 1)
            
            # Store the original data with the item
            list_item.setData(Qt.ItemDataRole.UserRole, step_data)
            
            # Set the size hint and associate the widget with the list item
            list_item.setSizeHint(item_widget.sizeHint())
            self.steps_list_widget.addItem(list_item)
            self.steps_list_widget.setItemWidget(list_item, item_widget)

    def get_rearranged_steps(self) -> List[Dict[str, Any]]:
        """Returns the new, reordered list of step data."""
        new_steps_data = []
        for i in range(self.steps_list_widget.count()):
            list_item = self.steps_list_widget.item(i)
            # Retrieve the data we stored earlier
            step_data = list_item.data(Qt.ItemDataRole.UserRole)
            new_steps_data.append(step_data)
        return new_steps_data
        
class StepInsertionDialog(QDialog):
    def __init__(self, execution_tree: QTreeWidget, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Insert Step At...")
        self.setFixedSize(600, 800)
        self.layout = QVBoxLayout(self)
        
        self.execution_tree_view = QTreeWidget()
        self.execution_tree_view.setHeaderHidden(True)
        self.execution_tree_view.setSelectionMode(QTreeWidget.SelectionMode.SingleSelection)
        
        self.insertion_mode_group = QGroupBox("Insertion Mode")
        self.insertion_mode_layout = QHBoxLayout()
        
        # SIMPLIFIED: Always enable both before and after
        self.insert_before_radio = QRadioButton("Insert Before Selected")
        self.insert_after_radio = QRadioButton("Insert After Selected")
        
        # Default to "Insert After"
        self.insert_after_radio.setChecked(True)
        
        self.insertion_mode_layout.addWidget(self.insert_before_radio)
        self.insertion_mode_layout.addWidget(self.insert_after_radio)
        self.insertion_mode_group.setLayout(self.insertion_mode_layout)
        
        self.layout.addWidget(QLabel("Select position for new step:"))
        self.layout.addWidget(self.execution_tree_view)
        self.layout.addWidget(self.insertion_mode_group)
        
        self._populate_tree_view(execution_tree)
        
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.layout.addWidget(self.button_box)
        
        self.selected_parent_item: Optional[QTreeWidgetItem] = None
        self.insert_mode: str = "after"

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

    def get_insertion_point(self) -> Tuple[Optional[QTreeWidgetItem], str]:
        self.selected_parent_item = self.execution_tree_view.currentItem()
        if self.insert_before_radio.isChecked(): 
            self.insert_mode = "before"
        else: 
            self.insert_mode = "after"
        return self.selected_parent_item, self.insert_mode


# --- REPLACEMENT CLASS: TemplateVariableMappingDialog (with intelligent default) ---
class TemplateVariableMappingDialog(QDialog):
    """A dialog to map variables from a template to global variables."""
# In the TemplateVariableMappingDialog class, replace the __init__ method

    def __init__(self, template_variables: set, global_variables: list, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Map Template Variables")
        self.setMinimumWidth(600)  # Increase width to fit the new field
        self.template_variables = sorted(list(template_variables))
        self.global_variables = global_variables
        self.mapping_widgets = {}

        main_layout = QVBoxLayout(self)
        form_layout = QFormLayout()

        for var_name in self.template_variables:
            row_layout = QHBoxLayout()
            
            action_combo = QComboBox()
            action_combo.addItems(["Map to Existing", "Create New"])
            
            existing_var_combo = QComboBox()
            existing_var_combo.addItem("-- Select Existing --")
            existing_var_combo.addItems(self.global_variables)
            
            new_var_name_editor = QLineEdit(var_name)
            new_var_name_editor.setPlaceholderText("Variable Name")

            # --- NEW: Add a textbox for the initial value ---
            new_var_value_editor = QLineEdit()
            new_var_value_editor.setPlaceholderText("Initial Value (optional)")
            # --- END NEW ---

            if var_name in self.global_variables:
                action_combo.setCurrentText("Map to Existing")
                existing_var_combo.setCurrentText(var_name)
            else:
                action_combo.setCurrentText("Create New")

            row_layout.addWidget(action_combo, 1)
            row_layout.addWidget(existing_var_combo, 2)
            row_layout.addWidget(new_var_name_editor, 2)
            
            # --- NEW: Add the value editor to the layout ---
            row_layout.addWidget(new_var_value_editor, 2)
            # --- END NEW ---
            
            form_layout.addRow(QLabel(f"Template Variable '{var_name}':"), row_layout)

            self.mapping_widgets[var_name] = {
                'action': action_combo,
                'existing': existing_var_combo,
                'new_name': new_var_name_editor,
                'new_value': new_var_value_editor # --- NEW: Store the widget ---
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


# In the TemplateVariableMappingDialog class, replace the _toggle_inputs method

    def _toggle_inputs(self, var_name: str, index: int):
        """Shows or hides the input widgets based on the selected action."""
        widgets = self.mapping_widgets[var_name]
        is_mapping_to_existing = (index == 0)
        
        widgets['existing'].setVisible(is_mapping_to_existing)
        widgets['new_name'].setVisible(not is_mapping_to_existing)
        widgets['new_value'].setVisible(not is_mapping_to_existing) # --- NEW: Toggle the value editor ---

# In the TemplateVariableMappingDialog class, replace the get_mapping method

    def get_mapping(self) -> Optional[Tuple[Dict[str, str], Dict[str, Any]]]:
        mapping = {}
        new_variables = {}
        for var_name, widgets in self.mapping_widgets.items():
            action_index = widgets['action'].currentIndex()
            
            if action_index == 0:  # Map to Existing
                target_var = widgets['existing'].currentText()
                if target_var == "-- Select Existing --":
                    QMessageBox.warning(self, "Input Error", f"Please select an existing variable to map '{var_name}' to.")
                    return None
                mapping[var_name] = target_var
            
            else:  # Create New
                target_var_name = widgets['new_name'].text().strip()
                if not target_var_name:
                    QMessageBox.warning(self, "Input Error", f"The new variable name for '{var_name}' cannot be empty.")
                    return None
                
                mapping[var_name] = target_var_name
                
                if target_var_name not in self.global_variables:
                    # --- NEW: Get the initial value from the textbox ---
                    initial_value_str = widgets['new_value'].text()

                    # If the textbox is empty, the value will be None
                    if not initial_value_str:
                        new_variables[target_var_name] = None
                    else:
                        # Safely evaluate the string to get Python types (int, float, bool, etc.)
                        try:
                            # ast.literal_eval is a safe way to evaluate simple literals
                            parsed_value = ast.literal_eval(initial_value_str)
                        except (ValueError, SyntaxError):
                            # If it can't be parsed (e.g., just text), treat it as a string
                            parsed_value = initial_value_str
                        
                        new_variables[target_var_name] = parsed_value
                    # --- END NEW ---

        return mapping, new_variables

# --- NEW CLASS: SaveTemplateDialog ---
# --- NEW DIALOG: WaitTimeConfigDialog ---
class WaitTimeConfigDialog(QDialog):
    """A dialog to configure the wait time using a hardcoded value or a global variable."""
    def __init__(self, global_variables: List[str], initial_config: Dict[str, Any], parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Configure Wait Time Between Steps")
        self.setMinimumWidth(400)
        self.global_variables = global_variables
        
        main_layout = QVBoxLayout(self)
        form_layout = QFormLayout()

        # --- Input Mode Selection ---
        self.hardcoded_radio = QRadioButton("Use a fixed number of seconds:")
        self.variable_radio = QRadioButton("Use a Global Variable:")
        
        # --- Input Widgets ---
        self.wait_time_editor = QLineEdit()
        # Use a QDoubleValidator to allow for fractional seconds (e.g., 0.5)
        self.wait_time_editor.setValidator(QtGui.QDoubleValidator(0, 300, 2)) 

        self.global_var_combo = QComboBox()
        self.global_var_combo.addItem("-- Select Variable --")
        self.global_var_combo.addItems(sorted(self.global_variables))

        # --- Layout ---
        form_layout.addRow(self.hardcoded_radio, self.wait_time_editor)
        form_layout.addRow(self.variable_radio, self.global_var_combo)
        main_layout.addLayout(form_layout)

        # --- Buttons ---
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        main_layout.addWidget(button_box)
        
        # --- Connections ---
        self.hardcoded_radio.toggled.connect(self._toggle_inputs)
        
        # --- Set Initial State ---
        self.set_config(initial_config)

    def _toggle_inputs(self):
        """Enable/disable input fields based on which radio button is selected."""
        self.wait_time_editor.setEnabled(self.hardcoded_radio.isChecked())
        self.global_var_combo.setEnabled(self.variable_radio.isChecked())

    def set_config(self, config: Dict[str, Any]):
        """Sets the dialog's state from an existing configuration."""
        if config.get('type') == 'variable':
            self.variable_radio.setChecked(True)
            idx = self.global_var_combo.findText(config.get('value', ''))
            if idx != -1:
                self.global_var_combo.setCurrentIndex(idx)
        else: # Default to hardcoded
            self.hardcoded_radio.setChecked(True)
            self.wait_time_editor.setText(str(config.get('value', 0)))
        self._toggle_inputs()

    def get_config(self) -> Optional[Dict[str, Any]]:
        """Returns the new configuration dictionary."""
        if self.hardcoded_radio.isChecked():
            try:
                # Allow floating point numbers for more precision
                value = float(self.wait_time_editor.text())
                if value < 0:
                    raise ValueError
                return {'type': 'hardcoded', 'value': value}
            except (ValueError, TypeError):
                QMessageBox.warning(self, "Input Error", "Please enter a valid non-negative number for the wait time.")
                return None
        elif self.variable_radio.isChecked():
            var_name = self.global_var_combo.currentText()
            if var_name == "-- Select Variable --":
                QMessageBox.warning(self, "Input Error", "Please select a global variable for the wait time.")
                return None
            return {'type': 'variable', 'value': var_name}
        return None
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


# In main_app.py, REPLACE your entire GroupedTreeWidget class with this:

class GroupedTreeWidget(QTreeWidget):
    step_reorder_requested = pyqtSignal(int, int)  # NEW SIGNAL
    
    def __init__(self, main_window: 'MainWindow', parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.main_window = main_window
        
        # Enable drag and drop
        self.setAcceptDrops(True)
        self.setDragEnabled(False)  # We handle dragging in cards
        self.setDefaultDropAction(Qt.DropAction.MoveAction)

    def dragEnterEvent(self, event: QtGui.QDragEnterEvent):
        if event.mimeData().hasText():
            event.acceptProposedAction()

    def dragMoveEvent(self, event: QtGui.QDragMoveEvent):
        if event.mimeData().hasText():
            event.acceptProposedAction()
            
            # Provide visual feedback
            item = self.itemAt(event.position().toPoint() if hasattr(event, 'position') else event.pos())
            if item:
                self.setCurrentItem(item)

    def dropEvent(self, event: QtGui.QDropEvent):
        if event.mimeData().hasText():
            source_index = int(event.mimeData().text())
            
            # Calculate target position
            drop_pos = event.position().toPoint() if hasattr(event, 'position') else event.pos()
            target_item = self.itemAt(drop_pos)
            
            if target_item:
                target_card = self.itemWidget(target_item, 0)
                if isinstance(target_card, ExecutionStepCard):
                    target_index = target_card.step_data.get("original_listbox_row_index", -1)
                    
                    # Determine if dropping above or below based on position
                    item_rect = self.visualItemRect(target_item)
                    if drop_pos.y() < item_rect.center().y():
                        # Drop above (before)
                        final_target = target_index
                    else:
                        # Drop below (after)
                        final_target = target_index + 1
                    
                    if source_index != final_target:
                        self.step_reorder_requested.emit(source_index, final_target)
            else:
                # Drop at end
                if source_index < len(self.main_window.added_steps_data):
                    self.step_reorder_requested.emit(source_index, len(self.main_window.added_steps_data))
            
            event.acceptProposedAction()

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:
        # First, run the default paint event to draw all the items.
        super().paintEvent(event)
        
        # Now, draw our custom lines on top.
        painter = QPainter(self.viewport())
        pen = QPen(QColor("#C0392B"))
        pen.setWidth(2)
        painter.setPen(pen)

        self._draw_group_lines_recursive(self.invisibleRootItem(), painter)

    def _draw_group_lines_recursive(self, parent_item: QTreeWidgetItem, painter: QPainter):
        for i in range(parent_item.childCount()):
            child_item = parent_item.child(i)
            
            item_data = self.main_window._get_item_data(child_item)
            
            if (item_data and item_data.get("type") in ["group_start", "loop_start", "IF_START"] and child_item.isExpanded()):
                
                start_index = item_data.get("original_listbox_row_index")
                _ , end_index = self.main_window._find_block_indices(start_index)
                
                end_item = self.main_window.data_to_item_map.get(end_index)
                
                start_rect = self.visualItemRect(child_item)
                
                if end_item and not start_rect.isEmpty():
                    end_rect = self.visualItemRect(end_item)
                    
                    x_offset = 8
                    tick_width = 5
                    start_y = start_rect.center().y()
                    
                    end_y = end_rect.center().y() if not end_rect.isEmpty() else self.viewport().height()

                    if end_y < start_y:
                        continue
                        
                    painter.drawLine(x_offset, start_y, x_offset, end_y)
                    painter.drawLine(x_offset, start_y, x_offset + tick_width, start_y)
                    if not end_rect.isEmpty():
                        painter.drawLine(x_offset, end_y, x_offset + tick_width, end_y)

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
class UpdateDialog(QDialog):
    def __init__(self, update_folder, target_folder, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Update Application")
        self.resize(800, 600)

        self.update_folder = update_folder
        self.target_folder = target_folder
        
        self.layout = QVBoxLayout(self)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["File", "Status"])
        self.tree.setColumnWidth(0, 500)
        self.layout.addWidget(self.tree)

        # Button Layout
        button_layout = QHBoxLayout()
        self.select_all_btn = QPushButton("Select All")
        self.select_all_btn.clicked.connect(self.select_all)
        button_layout.addWidget(self.select_all_btn)

        self.deselect_all_btn = QPushButton("Deselect All")
        self.deselect_all_btn.clicked.connect(self.deselect_all)
        button_layout.addWidget(self.deselect_all_btn)

        button_layout.addStretch() # Add space between button groups

        self.update_button = QPushButton("Update Selected Files")
        self.update_button.clicked.connect(self.update_files)
        button_layout.addWidget(self.update_button)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject) # Closes the dialog
        button_layout.addWidget(self.cancel_button)
        
        self.layout.addLayout(button_layout)

        self.populate_tree()

    def _set_all_items_checked_state(self, state):
        """Helper function to set the check state of all items."""
        iterator = QTreeWidgetItemIterator(self.tree)
        while iterator.value():
            item = iterator.value()
            item.setCheckState(0, state)
            iterator += 1

    def select_all(self):
        self._set_all_items_checked_state(Qt.CheckState.Checked)

    def deselect_all(self):
        self._set_all_items_checked_state(Qt.CheckState.Unchecked)

    def populate_tree(self):
        for root, _, files in os.walk(self.update_folder):
            for file in files:
                update_path = os.path.join(root, file)
                relative_path = os.path.relpath(update_path, self.update_folder)
                target_path = os.path.join(self.target_folder, relative_path)

                status = "New"
                if os.path.exists(target_path):
                    status = "Overwrite"

                self.add_tree_item(relative_path, status)

    def add_tree_item(self, path, status):
        parent_item = self.tree.invisibleRootItem()
        parts = path.split(os.sep)
        for i, part in enumerate(parts):
            is_file = i == len(parts) - 1
            item = self.find_item(parent_item, part)
            if item is None:
                item = QTreeWidgetItem(parent_item, [part, ""])
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(0, Qt.CheckState.Unchecked)
                if not is_file:
                     item.setText(1, "Folder")
                else:
                    item.setText(1, status)
            parent_item = item

    def find_item(self, parent, text):
        for i in range(parent.childCount()):
            child = parent.child(i)
            if child.text(0) == text:
                return child
        return None

    def update_files(self):
        checked_items = self.get_checked_items()
        if not checked_items:
            QMessageBox.information(self, "No Files Selected", "Please select files to update.")
            return
            
        for item_path in checked_items:
            source_path = os.path.join(self.update_folder, item_path)
            dest_path = os.path.join(self.target_folder, item_path)
            os.makedirs(os.path.dirname(dest_path), exist_ok=True)
            shutil.copy2(source_path, dest_path)
        QMessageBox.information(self, "Update Complete", "Selected files have been updated.")
        self.accept()

    def get_checked_items(self):
        checked = []
        root = self.tree.invisibleRootItem()
        self.get_checked_recursive(root, "", checked)
        return checked

    def get_checked_recursive(self, item, base_path, checked_list):
        for i in range(item.childCount()):
            child = item.child(i)
            current_path = os.path.join(base_path, child.text(0))
            if child.checkState(0) == Qt.CheckState.Checked:
                if child.childCount() == 0:  # It's a file
                    checked_list.append(current_path)
                else: # It's a folder
                    self.get_all_children(child, current_path, checked_list)
            else:
                self.get_checked_recursive(child, current_path, checked_list)

    def get_all_children(self, item, base_path, all_children_list):
        for i in range(item.childCount()):
            child = item.child(i)
            child_path = os.path.join(base_path, child.text(0))
            if child.childCount() == 0:
                all_children_list.append(child_path)
            else:
                self.get_all_children(child, child_path, all_children_list)

# In main_app.py
# REPLACE the entire WorkflowCanvas class with this:

# In main_app.py
# REPLACE the entire WorkflowCanvas class with this enhanced version:

class WorkflowCanvas(QWidget):
    """A custom widget that draws the bot's workflow with smart layout and navigation."""
    execute_step_requested = pyqtSignal(dict)
    def __init__(self, workflow_tree: List[Dict[str, Any]], main_window, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.workflow_tree = workflow_tree
        self.main_window = main_window  # Add this line
        
        # (rect, text, shape, step_data)
        self.nodes: List[Tuple[QRect, str, str, Dict[str, Any]]] = [] 
        
        self.edges: List[Tuple[Optional[str], Tuple[int, str], Tuple[int, str]]] = []
        self.merge_lines: List[Tuple[int, int]] = []
        
        # Smart layout parameters
        self.NODE_WIDTH = 220
        self.NODE_HEIGHT = 50
        self.V_SPACING = 40
        self.H_SPACING = 120
        self.GRID_SIZE = 20
        
        # Canvas navigation
        self.canvas_offset = QPoint(0, 0)
        self.dragging_canvas = False
        self.last_pan_point = QPoint()
        
        self.total_width = 0
        self.total_height = 0

        self.dragging_node_index: Optional[int] = None
        self.drag_offset: QPoint = QPoint()
        self.setMouseTracking(True)
        
        # Auto-layout on initialization
        self._smart_redraw_layout()
        self._adjust_canvas_size()
        self.setSizePolicy(QSizePolicy.Policy.MinimumExpanding, QSizePolicy.Policy.MinimumExpanding)
        
    def _should_show_group_collapsed(self, group_start_index: int) -> bool:
        """
        Determines if a group should be shown as collapsed (single box) or expanded (individual steps).
        
        Returns True if group should be collapsed (show as single box)
        Returns False if group should be expanded (show individual steps)
        """
        # Get the main window reference to access added_steps_data
        main_window = self.main_window
        if not hasattr(main_window, 'added_steps_data'):
            return True  # Default to collapsed if we can't access the data
        
        added_steps_data = main_window.added_steps_data
        
        # Check if there are any steps before this group
        has_steps_before = group_start_index > 0
        
        # Find the end of this group to check what comes after
        if group_start_index >= len(added_steps_data):
            return True
            
        group_step = added_steps_data[group_start_index]
        group_id = group_step.get("group_id")
        group_end_index = group_start_index
        
        # Find the matching group_end
        nesting_level = 0
        for i in range(group_start_index + 1, len(added_steps_data)):
            step = added_steps_data[i]
            step_type = step.get("type")
            
            if step_type == "group_start":
                nesting_level += 1
            elif step_type == "group_end":
                if nesting_level == 0 and step.get("group_id") == group_id:
                    group_end_index = i
                    break
                else:
                    nesting_level -= 1
        
        # Check if there are any steps after this group
        has_steps_after = group_end_index < len(added_steps_data) - 1
        
        # If there are steps before OR after, show as collapsed (single box)
        # If there are NO steps before AND NO steps after, show expanded (individual steps)
        return has_steps_before or has_steps_after       
        
    def _smart_redraw_layout(self):
        """Intelligently redraws the entire workflow with optimal spacing and positioning."""
        # Clear existing user positions to force smart layout
        self._clear_all_user_positions()
        
        # Clear existing layout
        self.nodes.clear()
        self.edges.clear()
        self.merge_lines.clear()
        
        # Calculate optimal starting position (centered at top)
        canvas_center_x = max(800, self.width()) // 2
        start_x = canvas_center_x - self.NODE_WIDTH // 2
        start_y = 30
        
        # Build the new layout
        self._build_smart_layout(start_x, start_y)
        
        # Center the workflow vertically if it fits in view
        self._center_workflow_if_possible()
        
        self.update()
        
    def _clear_all_user_positions(self):
        """Removes all user-defined positions to allow smart auto-layout."""
        def clear_positions_recursive(nodes: List[Dict]):
            for node in nodes:
                step_data = node['step_data']
                if "workflow_pos" in step_data:
                    del step_data["workflow_pos"]
                    
                clear_positions_recursive(node.get('children', []))
                clear_positions_recursive(node.get('false_children', []))
                
                if node.get('end_node'):
                    end_step_data = node['end_node']['step_data']
                    if "workflow_pos" in end_step_data:
                        del end_step_data["workflow_pos"]
        
        clear_positions_recursive(self.workflow_tree)
        
    def _build_smart_layout(self, start_x: int, start_y: int):
        """Builds the layout with smart spacing and positioning."""
        self._recursive_smart_layout(self.workflow_tree, start_x, start_y, -1)
        
    def _center_workflow_if_possible(self):
        """Centers the workflow vertically in the canvas if it fits."""
        if not self.nodes:
            return
            
        # Find the bounds of all nodes
        min_y = min(rect.y() for rect, _, _, _ in self.nodes)
        max_y = max(rect.bottom() for rect, _, _, _ in self.nodes)
        workflow_height = max_y - min_y
        
        available_height = self.height()
        if workflow_height < available_height:
            # Center vertically
            offset_y = (available_height - workflow_height) // 2 - min_y
            if offset_y > 0:
                # Move all nodes down by offset_y
                new_nodes = []
                for rect, text, shape, step_data in self.nodes:
                    new_rect = QRect(rect.x(), rect.y() + offset_y, rect.width(), rect.height())
                    step_data["workflow_pos"] = (new_rect.x(), new_rect.y())
                    new_nodes.append((new_rect, text, shape, step_data))
                self.nodes = new_nodes

    def _recursive_smart_layout(self, nodes: List[Dict], x: int, y: int, 
                               last_node_idx: int) -> Tuple[int, int, int]:
        """Smart recursive layout with optimal spacing."""
        current_y = y
        last_node_idx_in_level = last_node_idx
        max_x = x

        for node in nodes:
            step_data = node['step_data']
            step_type = step_data.get("type")

            current_step_num = step_data.get("original_listbox_row_index", -1) + 1
            text = self._get_node_text(step_data, current_step_num)
            
            # Determine node shape and color
            node_shape = self._get_node_shape(step_type)
            # ADD THIS NEW LOGIC FOR GROUPS
            if step_type == "group_start":
                group_start_index = step_data.get("original_listbox_row_index", -1)
                
                if self._should_show_group_collapsed(group_start_index):
                    # Show as single collapsed box (EXISTING behavior)
                    node_y = current_y
                    node_idx = len(self.nodes)
                    
                    # Add the group node as a single box
                    self.nodes.append((QRect(x, node_y, self.NODE_WIDTH, self.NODE_HEIGHT), text, node_shape, step_data))
                    step_data["workflow_pos"] = (x, node_y)
                    
                    if last_node_idx_in_level != -1:
                        self.edges.append((None, (last_node_idx_in_level, "bottom"), (node_idx, "top")))
                    
                    last_node_idx_in_level = node_idx
                    current_y += self.NODE_HEIGHT + self.V_SPACING
                    max_x = max(max_x, x + self.NODE_WIDTH)
                else:
                    # Show individual steps within the group (NEW behavior)
                    # Skip the group_start node itself and process children directly
                    children = node.get('children', [])
                    if children:
                        child_y, child_last_idx, child_max_x = self._recursive_smart_layout(
                            children, x, current_y, last_node_idx_in_level
                        )
                        current_y = child_y
                        last_node_idx_in_level = child_last_idx
                        max_x = max(max_x, child_max_x)
                    # Skip group_end as well since we're showing individual steps
                
                # Continue to next node
                continue
            if step_type in ["IF_START", "loop_start"]:
                node_y = current_y
                node_idx = len(self.nodes)
                
                # Add the control node
                self.nodes.append((QRect(x, node_y, self.NODE_WIDTH, self.NODE_HEIGHT), text, node_shape, step_data))
                step_data["workflow_pos"] = (x, node_y)
                
                if last_node_idx_in_level != -1:
                    self.edges.append((None, (last_node_idx_in_level, "bottom"), (node_idx, "top")))

                # Layout branches with smart spacing
                true_children = node.get('children', [])
                false_children = node.get('false_children', [])
                end_node_data = node.get('end_node')
                
                branch_start_y = node_y + self.NODE_HEIGHT + self.V_SPACING
                
                # Calculate branch positions with better spacing
                if true_children and false_children:
                    # Both branches exist - spread them out
                    true_x = x - self.H_SPACING
                    false_x = x + self.H_SPACING
                elif true_children:
                    # Only true branch - keep it centered under the control node
                    true_x = x
                    false_x = x
                else:
                    # No branches - continue straight down
                    true_x = x
                    false_x = x
                
                # Layout true branch
                first_true_node_idx = len(self.nodes)
                true_y_end, last_true_node_idx, true_max_x = self._recursive_smart_layout(
                    true_children, true_x, branch_start_y, -1
                )
                
                if true_children:
                    true_label = "True" if step_type == "IF_START" else "Loop Body"
                    true_port = "left" if step_type == "IF_START" and false_children else "bottom"
                    self.edges.append((true_label, (node_idx, true_port), (first_true_node_idx, "top")))

                # Layout false branch (for IF statements)
                false_y_end = branch_start_y
                last_false_node_idx = -1
                false_max_x = false_x
                
                if step_type == "IF_START" and false_children:
                    first_false_node_idx = len(self.nodes)
                    false_y_end, last_false_node_idx, false_max_x = self._recursive_smart_layout(
                        false_children, false_x, branch_start_y, -1
                    )
                    self.edges.append(("False", (node_idx, "right"), (first_false_node_idx, "top")))

                # Position end node smartly
                if end_node_data:
                    end_step_num = end_node_data['step_data'].get("original_listbox_row_index", -1) + 1
                    end_text = self._get_node_text(end_node_data['step_data'], end_step_num)
                    
                    # Smart positioning of end node
                    end_y = max(true_y_end, false_y_end) + self.V_SPACING
                    end_x = x  # Center under the control node
                    
                    end_node_rect = QRect(end_x, end_y, self.NODE_WIDTH, self.NODE_HEIGHT)
                    end_node_data['step_data']["workflow_pos"] = (end_x, end_y)
                    
                    end_node_idx = len(self.nodes)
                    end_node_shape = self._get_node_shape("IF_END" if step_type == "IF_START" else "loop_end")
                    self.nodes.append((end_node_rect, end_text, end_node_shape, end_node_data['step_data']))
                    
                    # Connect branches to end node
                    if last_true_node_idx != -1:
                        self.merge_lines.append((last_true_node_idx, end_node_idx))
                    if last_false_node_idx != -1:
                        self.merge_lines.append((last_false_node_idx, end_node_idx))
                    
                    # Add loop-back edge for loops
                    if step_type == "loop_start":
                        self.edges.append(("Loop Again", (end_node_idx, "right"), (node_idx, "top")))

                    last_node_idx_in_level = end_node_idx
                    current_y = end_y + self.NODE_HEIGHT + self.V_SPACING
                else:
                    current_y = max(true_y_end, false_y_end)
                    if last_true_node_idx != -1:
                        last_node_idx_in_level = last_true_node_idx
                
                max_x = max(max_x, true_max_x, false_max_x)
                
            else:
                # Regular step - simple vertical layout
                current_node_idx = len(self.nodes)
                rect = QRect(x, current_y, self.NODE_WIDTH, self.NODE_HEIGHT)
                step_data["workflow_pos"] = (x, current_y)
                
                self.nodes.append((rect, text, node_shape, step_data))
                
                if last_node_idx_in_level != -1:
                    self.edges.append((None, (last_node_idx_in_level, "bottom"), (current_node_idx, "top")))
                
                last_node_idx_in_level = current_node_idx
                current_y += self.NODE_HEIGHT + self.V_SPACING
                max_x = max(max_x, x + self.NODE_WIDTH)
            
            self.total_height = max(self.total_height, current_y)

        self.total_width = max(self.total_width, max_x)
        return current_y, last_node_idx_in_level, max_x

    def _get_node_shape(self, step_type: str) -> str:
        """Returns the appropriate shape for a given step type."""
        if step_type == "IF_START":
            return "diamond"
        elif step_type in ["loop_start", "loop_end"]:
            return "loop_rect"
        elif step_type in ["ELSE", "IF_END"]:
            return "rect_gray"
        elif step_type == "group_start":
            return "rect_group"
        else:
            return "rect"
        
    def _draw_grid(self, painter: QPainter):
        """Draws a subtle grid."""
        pen = QPen(QColor("#F0F0F0"))
        pen.setStyle(Qt.PenStyle.DotLine)
        painter.setPen(pen)
        
        # Adjust grid based on canvas offset
        offset_x = self.canvas_offset.x() % self.GRID_SIZE
        offset_y = self.canvas_offset.y() % self.GRID_SIZE
        
        width = self.width()
        height = self.height()
        
        for x in range(-offset_x, width, self.GRID_SIZE):
            painter.drawLine(x, 0, x, height)
            
        for y in range(-offset_y, height, self.GRID_SIZE):
            painter.drawLine(0, y, width, y)

    def wheelEvent(self, event: QtGui.QWheelEvent):
        """Enhanced wheel event with better zoom and pan."""
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            # Zoom functionality
            delta = event.angleDelta().y()
            zoom_factor = 1.15 if delta > 0 else (1 / 1.15)
            
            anchor_point = event.position()
            self._scale_layout(zoom_factor, anchor_point)
            self._adjust_canvas_size()
            self.update()
            event.accept()
        else:
            # Pan with mouse wheel (shift for horizontal)
            delta = event.angleDelta().y()
            pan_speed = 30
            
            if event.modifiers() == Qt.KeyboardModifier.ShiftModifier:
                # Horizontal pan
                self.canvas_offset.setX(self.canvas_offset.x() + (pan_speed if delta > 0 else -pan_speed))
            else:
                # Vertical pan
                self.canvas_offset.setY(self.canvas_offset.y() + (pan_speed if delta > 0 else -pan_speed))
            
            self.update()
            event.accept()

    # In the WorkflowCanvas class, modify the mousePressEvent method
    def mousePressEvent(self, event: QtGui.QMouseEvent):
        if event.button() == Qt.MouseButton.LeftButton:
            # ... (existing left-click logic for dragging remains the same)
            # Check if clicking on a node first
            clicked_pos = event.pos() - self.canvas_offset
            node_clicked = False
            
            for i in range(len(self.nodes) - 1, -1, -1):
                rect, _, _, _ = self.nodes[i]
                if rect.contains(clicked_pos):
                    self.dragging_node_index = i
                    self.drag_offset = clicked_pos - rect.topLeft()
                    self.setCursor(Qt.CursorShape.ClosedHandCursor)
                    
                    # Move clicked node to front
                    node = self.nodes.pop(i)
                    self.nodes.append(node)
                    self._remap_edges_for_drag(i, len(self.nodes) - 1)
                    self.dragging_node_index = len(self.nodes) - 1
                    self.update()
                    node_clicked = True
                    break
            
            if not node_clicked:
                self.dragging_canvas = True
                self.last_pan_point = event.pos()
                self.setCursor(Qt.CursorShape.ClosedHandCursor)
        
        elif event.button() == Qt.MouseButton.RightButton:
            # --- NEW LOGIC FOR RIGHT-CLICK ---
            clicked_pos_local = event.pos() - self.canvas_offset
            clicked_node_data = None
            for rect, _, _, step_data in self.nodes:
                if rect.contains(clicked_pos_local):
                    clicked_node_data = step_data
                    break
            
            # Show the context menu
            self._show_context_menu(event.pos(), clicked_node_data)
            # --- END NEW LOGIC ---

        # Keep this call for other event handling
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QtGui.QMoveEvent):
        if self.dragging_node_index is not None:
            # Node dragging
            new_top_left = event.pos() - self.canvas_offset - self.drag_offset
            
            _, text, shape, step_data = self.nodes[self.dragging_node_index]
            old_rect = self.nodes[self.dragging_node_index][0]
            
            self.nodes[self.dragging_node_index] = (QRect(new_top_left, old_rect.size()), text, shape, step_data)
            self.update()
            
        elif self.dragging_canvas:
            # Canvas panning
            delta = event.pos() - self.last_pan_point
            self.canvas_offset += delta
            self.last_pan_point = event.pos()
            self.update()
            
        else:
            # Update cursor based on what's under the mouse
            cursor = Qt.CursorShape.ArrowCursor
            clicked_pos = event.pos() - self.canvas_offset
            
            for rect, _, _, _ in self.nodes:
                if rect.contains(clicked_pos):
                    cursor = Qt.CursorShape.OpenHandCursor
                    break
            
            self.setCursor(cursor)
        
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event: QtGui.QMouseEvent):
        if event.button() == Qt.MouseButton.LeftButton:
            if self.dragging_node_index is not None:
                # Snap node to grid
                rect, text, shape, step_data = self.nodes[self.dragging_node_index]
                top_left = rect.topLeft()

                snapped_x = round(top_left.x() / self.GRID_SIZE) * self.GRID_SIZE
                snapped_y = round(top_left.y() / self.GRID_SIZE) * self.GRID_SIZE

                new_top_left = QPoint(snapped_x, snapped_y)
                snapped_rect = QRect(new_top_left, rect.size())

                step_data["workflow_pos"] = (snapped_rect.x(), snapped_rect.y())
                self.nodes[self.dragging_node_index] = (snapped_rect, text, shape, step_data)
                self.update()

            self.dragging_node_index = None
            self.dragging_canvas = False
            self.setCursor(Qt.CursorShape.ArrowCursor)

        super().mouseReleaseEvent(event)

# In the WorkflowCanvas class, update the _show_context_menu method
    def _show_context_menu(self, pos: QPoint, clicked_node_data: Optional[Dict] = None):
        """Shows a context menu with workflow options."""
        context_menu = QMenu(self)
        
        # --- NEW LOGIC: Store actions to check which was clicked ---
        configure_action = None
        execute_action = None
        
        if clicked_node_data:
            step_type = clicked_node_data.get("type")
            
            # Action 1: Execute This Step (only for executable steps)
            if step_type not in ["group_end", "loop_end", "IF_END", "ELSE"]:
                execute_action = context_menu.addAction("▶️ Execute This Step")

            # Action 2: Configure Parameters
            if step_type in ["step", "loop_start", "IF_START", "group_start"]:
                configure_action = context_menu.addAction("⚙️ Configure Parameters")

            if execute_action or configure_action:
                 context_menu.addSeparator()
        
        # General Canvas Actions
        redraw_action = context_menu.addAction("🔄 Smart Redraw Layout")
        reset_zoom_action = context_menu.addAction("🔍 Reset Zoom")
        center_action = context_menu.addAction("🎯 Center Workflow")
        
        # Show the menu and get the selected action
        action = context_menu.exec(self.mapToGlobal(pos))
        
        # --- Handle the clicked action ---
        if action == execute_action and clicked_node_data:
            # Emit our new signal with the step's data
            self.execute_step_requested.emit(clicked_node_data)

        elif action == configure_action and clicked_node_data:
            # This existing logic remains the same
            self.main_window.edit_step_from_data(clicked_node_data)
            
        elif action == redraw_action:
            self._smart_redraw_layout()
        elif action == reset_zoom_action:
            self._reset_zoom()
        elif action == center_action:
            self._center_workflow()

    def _reset_zoom(self):
        """Resets zoom to default size."""
        self.NODE_WIDTH = 220
        self.NODE_HEIGHT = 50
        self.V_SPACING = 40
        self.H_SPACING = 120
        self.GRID_SIZE = 20
        self._smart_redraw_layout()

    def _center_workflow(self):
        """Centers the workflow in the current view."""
        if not self.nodes:
            return
        
        # Calculate workflow bounds
        min_x = min(rect.x() for rect, _, _, _ in self.nodes)
        max_x = max(rect.right() for rect, _, _, _ in self.nodes)
        min_y = min(rect.y() for rect, _, _, _ in self.nodes)
        max_y = max(rect.bottom() for rect, _, _, _ in self.nodes)
        
        workflow_center_x = (min_x + max_x) // 2
        workflow_center_y = (min_y + max_y) // 2
        
        canvas_center_x = self.width() // 2
        canvas_center_y = self.height() // 2
        
        # Calculate offset to center the workflow
        offset_x = canvas_center_x - workflow_center_x
        offset_y = canvas_center_y - workflow_center_y
        
        self.canvas_offset = QPoint(offset_x, offset_y)
        self.update()

    # ... (keep all the existing methods like _scale_layout, _get_port_pos, _draw_right_angle_line, etc.)
    # I'll include the essential ones that might need updates:
    
    def _scale_layout(self, factor: float, anchor: QPointF):
        """Enhanced scaling with better anchor point handling."""
        # Adjust anchor point for canvas offset
        adjusted_anchor = anchor - QPointF(self.canvas_offset)
        
        # Scale the base layout parameters
        self.NODE_WIDTH = max(50, int(self.NODE_WIDTH * factor))
        self.NODE_HEIGHT = max(30, int(self.NODE_HEIGHT * factor))
        self.V_SPACING = max(10, int(self.V_SPACING * factor))
        self.H_SPACING = max(20, int(self.H_SPACING * factor))
        self.GRID_SIZE = max(5, int(self.GRID_SIZE * factor))

        # Scale USER-DEFINED positions
        def scale_user_pos_recursive(nodes: List[Dict]):
            for node in nodes:
                step_data = node['step_data']
                if "workflow_pos" in step_data and step_data["workflow_pos"]:
                    try:
                        current_pos_tuple = step_data["workflow_pos"]
                        if isinstance(current_pos_tuple, (list, tuple)) and len(current_pos_tuple) == 2:
                            current_pos = QPointF(float(current_pos_tuple[0]), float(current_pos_tuple[1]))
                            new_pos_f = (current_pos - adjusted_anchor) * factor + adjusted_anchor
                            step_data["workflow_pos"] = (int(new_pos_f.x()), int(new_pos_f.y()))
                        else:
                            del step_data["workflow_pos"]
                    except (TypeError, ValueError, IndexError):
                        if "workflow_pos" in step_data:
                            del step_data["workflow_pos"]

                scale_user_pos_recursive(node.get('children', []))
                scale_user_pos_recursive(node.get('false_children', []))
                
                if node.get('end_node'):
                    end_node_step_data = node['end_node']['step_data']
                    if "workflow_pos" in end_node_step_data and end_node_step_data["workflow_pos"]:
                        try:
                            current_pos_tuple = end_node_step_data["workflow_pos"]
                            if isinstance(current_pos_tuple, (list, tuple)) and len(current_pos_tuple) == 2:
                                current_pos = QPointF(float(current_pos_tuple[0]), float(current_pos_tuple[1]))
                                new_pos_f = (current_pos - adjusted_anchor) * factor + adjusted_anchor
                                end_node_step_data["workflow_pos"] = (int(new_pos_f.x()), int(new_pos_f.y()))
                            else:
                                del end_node_step_data["workflow_pos"]
                        except (TypeError, ValueError, IndexError):
                            if "workflow_pos" in end_node_step_data:
                                del end_node_step_data["workflow_pos"]

        scale_user_pos_recursive(self.workflow_tree)

        # Clear and rebuild layout
        self.nodes.clear()
        self.edges.clear()
        self.merge_lines.clear()
        self.total_width = 0
        self.total_height = 0

        # Rebuild with new parameters
        canvas_center_x = max(800, self.width()) // 2
        start_x = canvas_center_x - self.NODE_WIDTH // 2
        start_y = 30
        self._recursive_smart_layout(self.workflow_tree, start_x, start_y, -1)

    def paintEvent(self, event: QtGui.QPaintEvent):
        """Enhanced paint event with execution status visualization."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Apply canvas offset
        painter.translate(self.canvas_offset)
        
        self._draw_grid(painter)
        
        font = QFont("Arial", max(8, int(9 * (self.NODE_WIDTH / 220))))
        painter.setFont(font)

        labels_to_draw = []

        # Draw edges with enhanced styling
        for label, (from_idx, from_port), (to_idx, to_port) in self.edges:
            try:
                p1 = self._get_port_pos(from_idx, from_port)
                p2 = self._get_port_pos(to_idx, to_port)
            except IndexError:
                continue

            # Use different colors for different edge types
            if label == "True":
                painter.setPen(QPen(QColor("#28a745"), 2))  # Green for True
            elif label == "False":
                painter.setPen(QPen(QColor("#dc3545"), 2))  # Red for False  
            elif label and "Loop" in label:
                painter.setPen(QPen(QColor("#007bff"), 2))  # Blue for Loop
            else:
                painter.setPen(QPen(Qt.GlobalColor.black, 2))  # Black for normal flow

            line_points = self._draw_right_angle_line(painter, p1, p2, from_port, to_port)

            if label:
                # Enhanced label positioning
                if from_port == "left":
                    mid_point = line_points[0] + QPoint(-40, -15)
                    label_rect = QRect(-30, -10, 60, 20)
                elif from_port == "right":
                    mid_point = line_points[0] + QPoint(40, -15)
                    label_rect = QRect(-30, -10, 60, 20)
                elif label == "Loop Again":
                    mid_point = (line_points[1] + line_points[2]) / 2 + QPoint(0, -15)
                    label_rect = QRect(-40, -10, 80, 20)
                else:
                    mid_point = line_points[0] + QPoint(40, 15)
                    label_rect = QRect(-40, -10, 80, 20)
                
                labels_to_draw.append((label, mid_point, label_rect))

        # Draw merge lines
        painter.setPen(QPen(Qt.GlobalColor.gray, 2, Qt.PenStyle.DashLine))
        for from_idx, to_idx in self.merge_lines:
            try:
                from_rect, _, _, _ = self.nodes[from_idx]
                to_rect, _, _, _ = self.nodes[to_idx]
            except IndexError:
                continue

            from_port = "bottom"
            if self.nodes[from_idx][2] == "diamond":
                from_port = "left" if from_rect.center().x() >= to_rect.center().x() else "right"
            
            p1 = self._get_port_pos(from_idx, from_port)
            p2 = self._get_port_pos(to_idx, "top")
            
            merge_y = p2.y() - self.V_SPACING // 2
            points = [p1, QPoint(p1.x(), merge_y), QPoint(p2.x(), merge_y), p2]
            
            for i in range(len(points) - 1):
                painter.drawLine(points[i], points[i+1])

        # Draw nodes with execution status visualization
        for rect, text, shape, step_data in self.nodes:
            
            # Determine execution status from step_data
            execution_status = step_data.get("execution_status", "normal")  # normal, running, completed, error
            
            # Set border properties based on execution status
            if execution_status == "running":
                border_color = QColor("#FFD700")  # Gold/Yellow
                border_width = 4  # 100% thicker
                fill_color = QColor("#FFFBF0")  # Light yellow background
            elif execution_status == "error":
                border_color = QColor("#DC3545")  # Red
                border_width = 4  # 100% thicker  
                fill_color = QColor("#FFF5F5")  # Light red background
            elif execution_status == "completed":
                border_color = QColor("#28a745")  # Green
                border_width = 2  # Normal thickness
                fill_color = QColor("#F8FFF8")  # Light green background
            else:
                # Default styling based on node type
                border_color = QColor("#333333")
                border_width = 2
                if shape == "rect":
                    fill_color = QColor("#f8f9fa")
                elif shape == "diamond":
                    fill_color = QColor("#e3f2fd")
                elif shape == "loop_rect":
                    fill_color = QColor("#e8f5e8")
                elif shape == "rect_gray":
                    fill_color = QColor("#f5f5f5")
                elif shape == "rect_group":
                    fill_color = QColor("#fce4ec")
                else:
                    fill_color = QColor("#f8f9fa")
            
            # Draw node with dynamic styling
            painter.setPen(QPen(border_color, border_width))
            painter.setBrush(fill_color)
            
            # Draw node shape
            if shape == "diamond":
                poly = QPolygonF([
                    QPointF(rect.center() - QPoint(rect.width() // 2, 0)),
                    QPointF(rect.center() - QPoint(0, rect.height() // 2)),
                    QPointF(rect.center() + QPoint(rect.width() // 2, 0)),
                    QPointF(rect.center() + QPoint(0, rect.height() // 2))
                ])
                painter.drawPolygon(poly)
            else:
                painter.drawRoundedRect(rect, 8, 8)

            # Add execution status indicator (optional - small icon in corner)
            if execution_status == "running":
                # Draw a small spinning indicator or pulse effect
                painter.save()
                painter.setPen(QPen(QColor("#FF8C00"), 2))
                painter.setBrush(Qt.BrushStyle.NoBrush)
                indicator_rect = QRect(rect.right() - 15, rect.top() + 5, 10, 10)
                painter.drawEllipse(indicator_rect)
                painter.restore()
            elif execution_status == "completed":
                # Draw a small checkmark
                painter.save()
                painter.setPen(QPen(QColor("#28a745"), 3))
                check_x = rect.right() - 15
                check_y = rect.top() + 10
                painter.drawLine(check_x, check_y, check_x + 3, check_y + 3)
                painter.drawLine(check_x + 3, check_y + 3, check_x + 8, check_y - 2)
                painter.restore()
            elif execution_status == "error":
                # Draw a small X mark
                painter.save()
                painter.setPen(QPen(QColor("#DC3545"), 3))
                x_mark_x = rect.right() - 15
                x_mark_y = rect.top() + 5
                painter.drawLine(x_mark_x, x_mark_y, x_mark_x + 8, x_mark_y + 8)
                painter.drawLine(x_mark_x + 8, x_mark_y, x_mark_x, x_mark_y + 8)
                painter.restore()

            # --- MODIFICATION TO SHOW RESULT STARTS HERE ---
            # Draw text with better formatting, including the result if available
            painter.setPen(Qt.GlobalColor.black)
            text_rect = rect.adjusted(5, 5, -5, -5)
            
            # 1. Check for a completed result
            result_message = step_data.get("execution_result")
            execution_status = step_data.get("execution_status")
            
            if result_message and execution_status == "completed":
                # 2. Try to parse the result (similar to ExecutionStepCard)
                display_result_text = ""
                assign_to_var = step_data.get("assign_to_variable_name")
                
                if step_data.get("type") == "step" and assign_to_var and "Result: " in result_message:
                    try:
                        # Extract the actual result value
                        if " (Assigned to @" in result_message:
                            result_val_str = result_message.split("Result: ")[1].split(" (Assigned to @")[0]
                        elif " (Assigned to" in result_message:
                             result_val_str = result_message.split("Result: ")[1].split(" (Assigned to")[0]
                        else:
                            result_val_str = result_message.split("Result: ")[1]
                        
                        # Truncate if very long
                        if len(result_val_str) > 30:
                            result_val_str = result_val_str[:27] + "..."
                            
                        display_result_text = f"@{assign_to_var} = {result_val_str}"
                    except IndexError:
                        display_result_text = "✓ Completed" # Fallback
                
                else:
                    # For other types (loops, ifs) or steps without assignment
                    display_result_text = result_message
                    if len(display_result_text) > 35:
                         display_result_text = display_result_text[:32] + "..."

                # 3. Draw the main text (title) in the top half
                title_rect = QRect(text_rect.x(), text_rect.y(), text_rect.width(), text_rect.height() // 2)
                painter.drawText(title_rect, Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap, text)
                
                # 4. Draw the result text in the bottom half
                result_rect = QRect(text_rect.x(), text_rect.y() + text_rect.height() // 2, text_rect.width(), text_rect.height() // 2)
                
                result_font = QFont("Arial", max(7, int(8 * (self.NODE_WIDTH / 220))))
                result_font.setItalic(True)
                painter.setFont(result_font)
                painter.setPen(QColor("#155724")) # Dark green, like on the card
                
                painter.drawText(result_rect, Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap, display_result_text)
                
                # 5. Restore the original font for the next node
                painter.setFont(font)
                
            else:
                # 6. If no result, just draw the main text centered
                painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap, text)
            # --- MODIFICATION ENDS HERE ---

        # Draw edge labels with enhanced styling
        for label, mid_point, label_rect in labels_to_draw:
            painter.save()
            painter.translate(mid_point)
            painter.setPen(QPen(Qt.GlobalColor.darkGray, 1))
            painter.setBrush(QColor("#ffffff"))
            painter.drawRoundedRect(label_rect, 3, 3)
            
            # Color-coded text
            if label == "True":
                painter.setPen(QColor("#28a745"))
            elif label == "False":
                painter.setPen(QColor("#dc3545"))
            elif "Loop" in label:
                painter.setPen(QColor("#007bff"))
            else:
                painter.setPen(Qt.GlobalColor.black)
                
            painter.drawText(label_rect, Qt.AlignmentFlag.AlignCenter, label)
            painter.restore()
            
    # Keep all other existing methods like _get_port_pos, _draw_right_angle_line, 
    # _get_node_text, _adjust_canvas_size, _remap_edges_for_drag, etc.
    # [Previous methods remain the same]
    
    def _get_port_pos(self, node_idx: int, port: str) -> QPoint:
        """Gets the QPoint for a given port on a node."""
        rect = self.nodes[node_idx][0]
        if port == "top":
            return rect.center() - QPoint(0, rect.height() // 2)
        if port == "bottom":
            return rect.center() + QPoint(0, rect.height() // 2)
        if port == "left":
            return rect.center() - QPoint(rect.width() // 2, 0)
        if port == "right":
            return rect.center() + QPoint(rect.width() // 2, 0)
        return rect.center()

    def _draw_right_angle_line(self, painter: QPainter, p1: QPoint, p2: QPoint, from_port: str, to_port: str):
        """Draws a rectilinear line between two points."""
        
        if from_port == "bottom" and to_port == "top":
            # Standard vertical connection
            mid_y = p1.y() + self.V_SPACING // 2
            points = [p1, QPoint(p1.x(), mid_y), QPoint(p2.x(), mid_y), p2]
        
        elif from_port in ("left", "right") and to_port == "top":
            # IF branch connection (True or False)
            points = [p1, QPoint(p2.x(), p1.y()), p2]
        
        elif from_port == "right" and to_port == "top":
            # Loop Recycle (loop_end -> loop_start)
            mid_x = p1.x() + self.H_SPACING // 2
            points = [p1, QPoint(mid_x, p1.y()), QPoint(mid_x, p2.y() - self.V_SPACING // 2),
                      QPoint(p2.x(), p2.y() - self.V_SPACING // 2), p2]
        
        else:
            # Default fallback
            points = [p1, QPoint(p1.x(), p2.y()), p2]
            
        for i in range(len(points) - 1):
            painter.drawLine(points[i], points[i+1])
        return points

    def _get_node_text(self, step_data: Dict[str, Any], step_num: int) -> str:
        """Generates a concise text for the workflow node."""
        step_type = step_data.get("type")
        
        title_prefix = f"{step_num}. "
        
        if step_type == "step":
            method_name = step_data.get("method_name", "Unknown")
            assign_var = step_data.get("assign_to_variable_name")
            if assign_var:
                return f"{title_prefix}@{assign_var} = {method_name}()"
            return f"{title_prefix}{method_name}()"
            
        elif step_type == "loop_start":
            config = step_data.get("loop_config", {})
            count_config = config.get("iteration_count_config", {})
            val = count_config.get("value", "N")
            if count_config.get("type") == "variable":
                val = f"@{val}"
            return f"{title_prefix}Loop {val} times"
            
        elif step_type == "IF_START":
            cond = step_data.get("condition_config", {}).get("condition", {})
            left = cond.get("left_operand", {}).get("value", "L")
            if cond.get("left_operand", {}).get("type") == "variable":
                left = f"@{left}"
            right = cond.get("right_operand", {}).get("value", "R")
            if cond.get("right_operand", {}).get("type") == "variable":
                right = f"@{right}"
            op = cond.get("operator", "??")
            return f"{title_prefix}IF ({left} {op} {right})"
            
        elif step_type == "ELSE":
            return f"{title_prefix}ELSE"
        elif step_type == "group_start":
            return f"{title_prefix}Group: {step_data.get('group_name', 'Unnamed')}"
        elif step_type in ["loop_end", "IF_END", "group_end"]:
            return f"End {step_type.split('_')[0].capitalize()}"
            
        return f"{title_prefix}{step_type.replace('_', ' ').title()}"

    def _adjust_canvas_size(self):
        """Recalculates the minimum size of the canvas to fit all nodes."""
        if not self.nodes:
            # If no nodes, a small default is fine.
            self.setMinimumSize(100, 100)
            return

        # Calculate the actual bounds of all nodes
        max_x = 0
        max_y = 0
        for rect, _, _, _ in self.nodes:
            max_x = max(max_x, rect.right())
            max_y = max(max_y, rect.bottom())
        
        # Add a 200px padding around the content
        required_width = max_x + 200
        required_height = max_y + 200
        
        # THIS IS THE KEY: Set the minimum size to the actual required size.
        # This tells the parent QScrollArea how big the canvas needs to be.
        self.setMinimumSize(required_width, required_height)
        
    def _remap_edges_for_drag(self, old_idx: int, new_idx: int):
        """Updates all edge indices when a node is moved in the list."""
        new_edges = []
        for label, (from_idx, from_port), (to_idx, to_port) in self.edges:
            new_from_idx = from_idx
            new_to_idx = to_idx
            
            if from_idx == old_idx: 
                new_from_idx = new_idx
            elif from_idx > old_idx: 
                new_from_idx -= 1
            
            if to_idx == old_idx: 
                new_to_idx = new_idx
            elif to_idx > old_idx: 
                new_to_idx -= 1
                
            new_edges.append((label, (new_from_idx, from_port), (new_to_idx, to_port)))
        self.edges = new_edges
        
        new_merges = []
        for from_idx, to_idx in self.merge_lines:
            new_from_idx = from_idx
            new_to_idx = to_idx

            if from_idx == old_idx: 
                new_from_idx = new_idx
            elif from_idx > old_idx: 
                new_from_idx -= 1
            
            if to_idx == old_idx: 
                new_to_idx = new_idx
            elif to_idx > old_idx: 
                new_to_idx -= 1
            
            new_merges.append((new_from_idx, new_to_idx))
        self.merge_lines = new_merges
        
# --- MAIN APPLICATION WINDOW ---
class MainWindow(QMainWindow):
    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.wait_time_between_steps = {'type': 'hardcoded', 'value': 0}
        self.setWindowTitle("Automate Your Task By Simple Bot - Developed by Phung Tuan Hung")
        
        # Get the geometry of the primary screen
        screen = QApplication.primaryScreen().geometry()
        width = int(screen.width() * 0.9)
        height = int(screen.height() * 0.9)
        x = int((screen.width() - width) / 2)
        y = int((screen.height() - height) / 2)
        self.setGeometry(x, y, width, height)

        # --- Initialize attributes ---
        self.gui_communicator = GuiCommunicator()
        self.base_directory = os.path.dirname(os.path.abspath(__file__))
        self.module_subfolder = "Bot_module"
        self.module_directory = os.path.join(self.base_directory, self.module_subfolder)
        self.click_image_dir = os.path.join(self.base_directory, "Click_image")
        os.makedirs(self.click_image_dir, exist_ok=True)
        self.schedules_directory = os.path.join(self.base_directory, "Schedules")
        self.schedules = {}
        
        icon_path = os.path.join(self.base_directory, "app_icon.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
            
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
        self.is_bot_running = False
        self.is_paused = False # ADD THIS
        self.worker: Optional[ExecutionWorker] = None # ADD THIS
        # --- Create UI ---
        self.init_ui()

        # --- Connect signals and start logic ---
        self.gui_communicator.log_message_signal.connect(self._log_to_console)
        self.gui_communicator.update_module_info_signal.connect(self.update_label_info_from_module)
        self.gui_communicator.update_click_signal.connect(self.update_click_signal_from_module)


        # --- Scheduler Setup ---
        self.schedule_timer = QTimer(self)
        self.schedule_timer.timeout.connect(self.check_schedules)
        self.schedule_timer.start(60000)
        self._log_to_console("Scheduler started. Will check for due tasks every minute.")
        
        # --- Load initial data ---
        self.load_all_modules_to_tree()
        self.load_saved_steps_to_tree()
        self._update_variables_list_display()
        self.minimized_for_execution = False
        self.original_geometry = None
        #--------------------------------------------------
        self.minimized_for_execution = False
        self.original_geometry = None
        self.widget_homes = {} # <<< ADD THIS LINE
        self.is_bot_running = False
        self.is_paused = False
        self.worker: Optional[ExecutionWorker] = None
        
    def init_ui(self) -> None:
        """Initialize the new UI with left menu, center content, and right panels."""
        os.makedirs(self.steps_template_directory, exist_ok=True)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        self.create_left_menu()
        self.create_center_content()
        self.create_right_panels()
        
        self.toggle_log_checkbox.toggled.connect(self.log_widget.setVisible)
        self.log_widget.setVisible(False)
        
        main_layout.addWidget(self.left_menu)
        main_layout.addWidget(self.center_content)
        main_layout.addWidget(self.right_panels)
        
        main_layout.setStretch(0, 0)
        main_layout.setStretch(1, 1)
        main_layout.setStretch(2, 0)

    def create_left_menu(self):
        """Creates the collapsible left-side menu."""
        self.left_menu = QWidget()
        self.left_menu.setFixedWidth(250)
        self.left_menu.setStyleSheet("""
            QWidget {
                background-color: #ecf0f1;
                border-right: 1px solid #bdc3c7;
            }
            QPushButton {
                text-align: left;
                padding: 8px 12px;
                border: none;
                border-radius: 4px;
                margin: 2px 8px;
                font-size: 13px;
                color: #2c3e50;
                background-color: #ffffff;
            }
            QPushButton:hover {
                background-color: #e8f4fd;
                border-left: 3px solid #3498db;
            }
            QPushButton:disabled {
                background-color: #f8f9f9;
                color: #95a5a6;
            }
            QPushButton.section-header {
                font-weight: bold;
                color: #34495e;
                background-color: transparent;
                border-bottom: 1px solid #dadedf;
                border-radius: 0;
                margin: 8px 0 4px 0;
            }
            QCheckBox {
                padding: 8px 12px;
                font-size: 13px;
            }
        """)
        menu_layout = QVBoxLayout(self.left_menu)
        
        # --- Brand ---
        self.website_label = QLabel('<a href="http://www.AutomateTask.Click" style="color: #3498db; text-decoration: none; font-size: 18pt; font-weight: bold;">AutomateTask</a>')
        self.website_label.setOpenExternalLinks(True)
        self.website_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        menu_layout.addWidget(self.website_label)
        
        # Helper function for sections
        def create_section(title, buttons):
            menu_layout.addWidget(QLabel(title, objectName="section-header"))
            for btn in buttons:
                menu_layout.addWidget(btn)

        # --- Execution ---
        self.execute_all_button = QPushButton("🚀 Execute All")
        self.execute_one_step_button = QPushButton("▶️ Execute Selected")
        create_section("Execution", [self.execute_all_button, self.execute_one_step_button])

        # --- Bot Management ---
        self.save_steps_button = QPushButton("💾 Save Bot")
        self.change_bot_folder_button = QPushButton("📂 Change Folder")
        create_section("Bot Management", [self.save_steps_button, self.change_bot_folder_button])

        # --- Step Editing ---
        self.add_loop_button = QPushButton("🔄 Add Loop")
        self.add_conditional_button = QPushButton("❓ Add Conditional")
        self.group_steps_button = QPushButton("📦 Group Selected")
        
        self.rearrange_steps_button = QPushButton("↕️ Rearrange Steps")
        
        self.clear_selected_button = QPushButton("🗑️ Clear Selected")
        self.remove_all_steps_button = QPushButton("❌ Remove All")

        create_section("Step Editing", [self.add_loop_button, self.add_conditional_button, self.group_steps_button, self.rearrange_steps_button, self.clear_selected_button, self.remove_all_steps_button])
        
        # --- Tools & Settings ---
        self.open_screenshot_tool_button = QPushButton("📷 Screenshot Tool")
        self.view_workflow_button = QPushButton("📈 View Workflow")
        self.set_wait_time_button = QPushButton("⏱️ Wait between Steps")
        self.update_app_btn = QPushButton("🔄 Update App")
        self.always_on_top_button = QPushButton("📌 Always On Top: Off")
        self.toggle_log_checkbox = QCheckBox("📋 Show Log")
        create_section("Tools & Settings", [self.open_screenshot_tool_button, self.view_workflow_button, self.update_app_btn, self.always_on_top_button, self.toggle_log_checkbox])
        create_section("Tools & Settings", [self.open_screenshot_tool_button, self.view_workflow_button, self.set_wait_time_button, self.update_app_btn, self.always_on_top_button, self.toggle_log_checkbox])
        
        # Connections
        #self.execute_all_button.clicked.connect(self.execute_all_steps)
        self.execute_all_button.clicked.connect(self._handle_execute_pause_resume)
        self.execute_one_step_button.clicked.connect(self.execute_one_step)
        self.save_steps_button.clicked.connect(self.save_bot_steps_dialog)
        self.change_bot_folder_button.clicked.connect(self.select_bot_steps_folder)
        self.add_loop_button.clicked.connect(self.add_loop_block)
        self.add_conditional_button.clicked.connect(self.add_conditional_block)
        self.group_steps_button.clicked.connect(self.group_selected_steps)
        self.rearrange_steps_button.clicked.connect(self.open_rearrange_steps_dialog)
        self.clear_selected_button.clicked.connect(self.clear_selected_steps)
        self.remove_all_steps_button.clicked.connect(self.clear_all_steps)
        self.open_screenshot_tool_button.clicked.connect(self.open_screenshot_tool)
        self.view_workflow_button.clicked.connect(lambda: self._update_workflow_tab(switch_to_tab=True))
        self.update_app_btn.clicked.connect(self.update_application)
        self.always_on_top_button.setCheckable(True)
        self.set_wait_time_button.clicked.connect(self.set_wait_time)
        self.always_on_top_button.clicked.connect(self.toggle_always_on_top)

        menu_layout.addStretch()

        # Progress Bar & Exit
        self.progress_bar = QProgressBar(); self.progress_bar.hide(); menu_layout.addWidget(self.progress_bar)
        self.exit_button = QPushButton("🚪 Exit"); self.exit_button.clicked.connect(QApplication.instance().quit); menu_layout.addWidget(self.exit_button)
        
# In MainWindow class
# REPLACE your existing create_center_content method with this one:
# In MainWindow class
# REPLACE your existing create_center_content method with this one:

    def create_center_content(self):
        """Creates the main tabbed interface, now with a dedicated focus mode layout."""
        self.center_content = QWidget()
        # Main layout that holds everything in the center panel
        main_center_layout = QVBoxLayout(self.center_content)
        main_center_layout.setContentsMargins(5, 5, 5, 5)

        # 1. Normal Mode UI (Tabs and Info)
        # Create a widget to hold the normal UI so we can hide/show it easily
        self.normal_mode_widget = QWidget()
        normal_mode_layout = QVBoxLayout(self.normal_mode_widget)
        normal_mode_layout.setContentsMargins(0, 0, 0, 0) # No extra margins needed here

        self.main_tab_widget = QTabWidget()
        self.create_saved_bots_tab()
        self.create_execution_flow_tab()
        self.create_workflow_tab()
        self.create_log_tab()
        normal_mode_layout.addWidget(self.main_tab_widget) # Add tabs to normal widget

        # Create the info widget (as before)
        self.info_widget = QFrame()
        self.info_widget.setFrameShape(QFrame.Shape.StyledPanel)
        info_layout = QHBoxLayout(self.info_widget)
        self.label_info1 = QLabel("Info:")
        self.label_info2 = QLabel("Image Preview")
        self.label_info2.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label_info2.setFixedSize(200, 30)
        self.label_info2.setStyleSheet("QLabel { background-color: #f0f0f0; border: 1px solid #dcdcdc; border-radius: 4px; color: #888; }")
        self.label_info3 = QLabel("Image Name")
        self.label_info3.setAlignment(Qt.AlignmentFlag.AlignRight)
        info_layout.addWidget(self.label_info1)
        info_layout.addWidget(self.label_info2)
        info_layout.addWidget(self.label_info3)
        normal_mode_layout.addWidget(self.info_widget) # Add info widget to normal widget

        # Add the entire normal mode widget to the main layout
        main_center_layout.addWidget(self.normal_mode_widget)


        # 2. Focus Mode UI (Initially hidden)
        # This is the new container for our focus mode UI
        self.focus_mode_widget = QWidget()
        self.focus_mode_layout = QVBoxLayout(self.focus_mode_widget)
        self.focus_mode_layout.setContentsMargins(0, 0, 0, 0)
        self.focus_mode_widget.setVisible(False) # Start hidden

        # Add the focus mode widget to the main layout
        main_center_layout.addWidget(self.focus_mode_widget)

    def create_right_panels(self):
        """Creates the right-side panels for modules and variables."""
        self.right_panels = QWidget()
        self.right_panels.setFixedWidth(400)
        self.right_panels.setStyleSheet("background-color: #f4f6f7;")
        right_layout = QVBoxLayout(self.right_panels)
        
        right_splitter = QSplitter(Qt.Orientation.Vertical)
        
        # Module Browser
        module_panel = QWidget()
        module_layout = QVBoxLayout(module_panel)
        module_layout.addWidget(QLabel("🔧 Module Browser", objectName="section-header"))
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Filter:"))
        self.module_filter_dropdown = QComboBox(); self.module_filter_dropdown.addItem("-- Show All --"); filter_layout.addWidget(self.module_filter_dropdown)
        module_layout.addLayout(filter_layout)
        self.search_box = QLineEdit(); self.search_box.setPlaceholderText("Search modules..."); module_layout.addWidget(self.search_box)
        self.module_tree = QTreeWidget(); self.module_tree.setHeaderHidden(True); module_layout.addWidget(self.module_tree)
        
        # Variables Panel
        vars_panel = QWidget()
        vars_layout = QVBoxLayout(vars_panel)
        vars_layout.addWidget(QLabel("📊 Global Variables", objectName="section-header"))
        self.variables_list = QListWidget(); vars_layout.addWidget(self.variables_list)
        btn_layout = QHBoxLayout()
        
        # --- Example Custom Styling ---
        # You can change these values as you like.
        button_style = """
            QPushButton {
                color: #000000;                 /* Text color (white) */
                background-color: #3498db;     /* A blue background */
                font-size: 13px;               /* Text size */
                padding: 5px 8px;              /* Vertical (5px) and horizontal (8px) padding */
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #2980b9;     /* Darker blue on hover */
            }
            QPushButton:pressed {
                background-color: #2471a3;     /* Even darker blue when pressed */
            }
        """

        # Create the buttons
        self.add_var_button = QPushButton("➕ Add")
        self.edit_var_button = QPushButton("✏️ Edit")
        self.delete_var_button = QPushButton("❌ Delete")
        self.clear_vars_button = QPushButton("🔄 Reset")

        # Apply the style to all four buttons
        self.add_var_button.setStyleSheet(button_style)
        self.edit_var_button.setStyleSheet(button_style)
        
        # --- Example of a different style for Delete/Reset ---
        # Just copy the 'button_style' string, change the colors, and apply it.
        # For example, to make Delete and Reset red:
        danger_style = button_style.replace("#3498db", "#e74c3c")  # Replace blue with red
                                  #.replace("#2980b9", "#c0392b")
                                  #.replace("#2471a3", "#a93226")
                                  
        self.delete_var_button.setStyleSheet(button_style)
        self.clear_vars_button.setStyleSheet(button_style)


        # Add buttons to the layout
        btn_layout.addWidget(self.add_var_button)
        btn_layout.addWidget(self.edit_var_button)
        btn_layout.addWidget(self.delete_var_button)
        btn_layout.addWidget(self.clear_vars_button)
        
        vars_layout.addLayout(btn_layout)

        right_splitter.addWidget(module_panel)
        right_splitter.addWidget(vars_panel)
        right_splitter.setSizes([500, 250])
        right_layout.addWidget(right_splitter)

        # Connections
        self.module_filter_dropdown.currentIndexChanged.connect(self.filter_module_tree)
        self.search_box.textChanged.connect(self.search_module_tree)
        self.module_tree.itemDoubleClicked.connect(self.add_item_to_execution_tree)
        self.module_tree.itemClicked.connect(self.update_selected_method_info)
        self.module_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.module_tree.customContextMenuRequested.connect(self.show_context_menu)
        self.add_var_button.clicked.connect(self.add_variable)
        self.edit_var_button.clicked.connect(self.edit_variable)
        self.delete_var_button.clicked.connect(self.delete_variable)
        self.clear_vars_button.clicked.connect(self.reset_all_variable_values)
        
    def create_execution_flow_tab(self):
        """
        Creates the 'Execution Flow' tab, containing only the title label
        and the main execution tree widget. The info labels have been moved out.
        """
        # Create the main widget for this tab
        exec_widget = QWidget()
        layout = QVBoxLayout(exec_widget)
        layout.setContentsMargins(5, 5, 5, 5) # Add some padding

        # Create and add the title label for this tab
        self.execution_flow_label = QLabel("Execution Flow")
        self.execution_flow_label.setStyleSheet("font-weight: bold; font-size: 14px; color: #2c3e50;")
        layout.addWidget(self.execution_flow_label)

        # Create and add the tree widget for displaying steps
        self.execution_tree = GroupedTreeWidget(self)
        self.execution_tree.setHeaderHidden(True)
        self.execution_tree.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.ExtendedSelection)
        self.execution_tree.itemSelectionChanged.connect(self._toggle_execute_one_step_button)
        layout.addWidget(self.execution_tree)

        # The info labels (label_info1, label_info2, label_info3) are
        # intentionally NOT created or added here anymore. They are now
        # handled in the `create_center_content` method to be globally visible.

        # Add the completed widget as a new tab
        self.main_tab_widget.addTab(exec_widget, "📋 Execution Flow")

    def create_saved_bots_tab(self):
        bots_widget = QWidget()
        layout = QVBoxLayout(bots_widget)
        layout.addWidget(QLabel("Saved Bots"))
        self.saved_steps_tree = QTreeWidget(); self.saved_steps_tree.setHeaderLabels(["Bot", "Schedule", "Status"]); self.saved_steps_tree.itemDoubleClicked.connect(self.saved_step_tree_item_selected); self.saved_steps_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu); self.saved_steps_tree.customContextMenuRequested.connect(self.show_saved_bot_context_menu); layout.addWidget(self.saved_steps_tree)
        self.main_tab_widget.addTab(bots_widget, "🤖 Saved Bots")

    def create_workflow_tab(self):
        flow_widget = QWidget()
        layout = QVBoxLayout(flow_widget)
        self.bot_workflow_label = QLabel("Bot Workflow")
        layout.addWidget(self.bot_workflow_label)
        self.workflow_scroll_area = QScrollArea(); 
        self.workflow_scroll_area.setWidgetResizable(True); 
        self.workflow_scroll_area.setMinimumSize(50, 50)
        layout.addWidget(self.workflow_scroll_area)

        self.workflow_tab_index = self.main_tab_widget.indexOf(flow_widget)
        self.main_tab_widget.addTab(flow_widget, "📈 Workflow")
        self.workflow_tab_index = self.main_tab_widget.indexOf(flow_widget)

    def create_log_tab(self):
        self.log_widget = QWidget()
        layout = QVBoxLayout(self.log_widget)
        layout.addWidget(QLabel("Execution Log"))
        self.log_console = QTextEdit(); self.log_console.setReadOnly(True); layout.addWidget(self.log_console)
        self.clear_log_button = QPushButton("Clear Log"); self.clear_log_button.clicked.connect(self.log_console.clear); layout.addWidget(self.clear_log_button)
        self.main_tab_widget.addTab(self.log_widget, "📜 Log")

    # All other methods from the original MainWindow go here...
    def _get_item_data(self, item: QTreeWidgetItem) -> Optional[Dict[str, Any]]:
        if not item: return None
        data = item.data(0, Qt.ItemDataRole.UserRole)
        return data.value() if isinstance(data, QVariant) else data

    def select_bot_steps_folder(self):
        """Opens a dialog to allow the user to select a different folder for saved bots."""
        current_dir = self.bot_steps_directory
        new_dir = QFileDialog.getExistingDirectory(self, "Select Bot Steps Folder", current_dir)
        
        if new_dir and new_dir != current_dir:
            self.bot_steps_directory = new_dir
            self.load_saved_steps_to_tree() # Refresh the tree from the new location
            self._log_to_console(f"Changed bot steps folder to: {new_dir}")

    def _handle_screenshot_request_from_param_dialog(self) -> None:
        if self.active_param_input_dialog:
            self.active_param_input_dialog.hide()
        self._log_to_console("ParameterInputDialog hidden, opening screenshot tool.")
        self.open_screenshot_tool()

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
        # This correctly reads from the central 'schedules.json' for the dialog
        schedule_data = self.schedules.get(bot_name)
        dialog = ScheduleTaskDialog(bot_name, schedule_data, self)
        
        if dialog.exec() == QDialog.DialogCode.Accepted:
            # Get the new schedule data from the dialog
            new_schedule_data = dialog.get_schedule_data()

            # --- FIX STARTS HERE ---

            # 1. Write to System 1 (schedules.json) - This part is the same as before
            self.schedules[bot_name] = new_schedule_data
            self.save_schedules()

            # 2. Write to System 2 (the bot's .csv file) - This is the new, required logic
            bot_file_path = os.path.join(self.bot_steps_directory, f"{bot_name}.csv")
            
            if not os.path.exists(bot_file_path):
                QMessageBox.warning(self, "Error", f"Could not find bot file '{bot_name}.csv' to save schedule.")
                return

            # Use the existing helper function to write the schedule into the .csv
            if not self._write_schedule_to_csv(bot_file_path, new_schedule_data):
                QMessageBox.critical(self, "Schedule Save Error", f"Failed to write schedule data to {bot_name}.csv.")
            else:
                self._log_to_console(f"Schedule for '{bot_name}' saved to .json and .csv")
                
            # --- FIX ENDS HERE ---

            # Refresh the "Saved Bots" tab display
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

    def _get_expansion_state(self) -> set:
        """Recursively finds all expanded block items and returns their unique IDs."""
        expanded_ids = set()
        
        def traverse(parent_item):
            for i in range(parent_item.childCount()):
                child_item = parent_item.child(i)
                if child_item.isExpanded():
                    item_data = self._get_item_data(child_item)
                    if item_data:
                        item_id = item_data.get("group_id") or item_data.get("loop_id") or item_data.get("if_id")
                        if item_id:
                            expanded_ids.add(item_id)
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
    
    def _rebuild_execution_tree(self, item_to_focus_data: Optional[Dict[str, Any]] = None) -> None:
        expanded_state = self._get_expansion_state()

        self.execution_tree.clear()
        self.data_to_item_map.clear()
        
        # CONNECT THE NEW TREE SIGNAL
        if hasattr(self.execution_tree, 'step_reorder_requested'):
            try:
                self.execution_tree.step_reorder_requested.disconnect()
            except TypeError:
                pass
        self.execution_tree.step_reorder_requested.connect(self._handle_step_reorder)
        
        current_parent_stack: List[QTreeWidgetItem] = [self.execution_tree.invisibleRootItem()]
        item_to_focus: Optional[QTreeWidgetItem] = None

        for i, step_data_dict in enumerate(self.added_steps_data):
            step_data_dict["original_listbox_row_index"] = i
            step_type = step_data_dict.get("type")

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

            parent_for_current_item = current_parent_stack[-1]
            tree_item = QTreeWidgetItem(parent_for_current_item)
            tree_item.setData(0, Qt.ItemDataRole.UserRole, QVariant(step_data_dict))
            
            self.data_to_item_map[i] = tree_item

            card = ExecutionStepCard(step_data_dict, i + 1)
            
            # Connect all signals
            card.edit_requested.connect(self._handle_edit_request)
            card.delete_requested.connect(self._handle_delete_request)
            card.move_up_requested.connect(self._handle_move_up_request)
            card.move_down_requested.connect(self._handle_move_down_request)
            card.save_as_template_requested.connect(self._handle_save_as_template_request)
            card.execute_this_requested.connect(self._handle_execute_this_request)
            # NEW DRAG AND DROP CONNECTIONS
            card.step_drag_started.connect(self._handle_step_drag_started)
            card.step_reorder_requested.connect(self._handle_step_reorder)
            
            tree_item.setSizeHint(0, card.sizeHint())
            self.execution_tree.setItemWidget(tree_item, 0, card)

            if item_to_focus_data and step_data_dict == item_to_focus_data:
                item_to_focus = tree_item

            if step_type in ["loop_start", "IF_START", "ELSE", "group_start"]:
                current_parent_stack.append(tree_item)

        if item_to_focus:
            self.execution_tree.setCurrentItem(item_to_focus)
            self.execution_tree.scrollToItem(item_to_focus, QtWidgets.QAbstractItemView.ScrollHint.PositionAtCenter)

        self.update_status_column_for_all_items()
        self._restore_expansion_state(expanded_state)
        self._update_workflow_tab(switch_to_tab=False)


    def _handle_edit_request(self, step_data_to_edit: Dict[str, Any]):
        item_to_edit = self._find_qtreewidget_item(step_data_to_edit)
        if item_to_edit:
            self.edit_step_in_execution_tree(item_to_edit, 0)
        else:
            idx = step_data_to_edit.get("original_listbox_row_index")
            if idx is not None and 0 <= idx < len(self.added_steps_data):
                item_from_map = self.data_to_item_map.get(idx)
                if item_from_map:
                    self.edit_step_in_execution_tree(item_from_map, 0)
                    return
            QMessageBox.critical(self, "Error", "Could not find the step to edit.")

    def _handle_delete_request(self, step_to_delete: Dict[str, Any]):
        step_type = step_to_delete.get("type")
        start_index = -1
        
        start_index = step_to_delete.get("original_listbox_row_index")
        if start_index is None or not (0 <= start_index < len(self.added_steps_data)) or self.added_steps_data[start_index].get("original_listbox_row_index") != start_index:
            try:
                start_index = self.added_steps_data.index(step_to_delete)
            except ValueError:
                 QMessageBox.critical(self, "Error", "Could not find step to delete in data list. Data may be out of sync.")
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
        current_pos = step_data.get("original_listbox_row_index")
        if current_pos is None or not (0 <= current_pos < len(self.added_steps_data)) or self.added_steps_data[current_pos].get("original_listbox_row_index") != current_pos:
            try:
                current_pos = self.added_steps_data.index(step_data)
            except ValueError:
                 return
        
        if current_pos == 0:
            return
        start_idx, end_idx = self._find_block_indices(current_pos)
        
        prev_start_idx, prev_end_idx = self._find_block_indices(start_idx - 1) if (start_idx -1) >= 0 else (start_idx-1, start_idx-1)
        if prev_start_idx == -1:
            return
        
        part_before = self.added_steps_data[0:prev_start_idx]
        part_middle = self.added_steps_data[prev_start_idx : start_idx]
        part_after = self.added_steps_data[end_idx+1:]
        
        block_to_move = self.added_steps_data[start_idx : end_idx + 1]
        
        self.added_steps_data = part_before + block_to_move + part_middle + part_after
        self._rebuild_execution_tree(item_to_focus_data=block_to_move[0])

    def _handle_move_down_request(self, step_data: Dict[str, Any]):
        current_pos = step_data.get("original_listbox_row_index")
        if current_pos is None or not (0 <= current_pos < len(self.added_steps_data)) or self.added_steps_data[current_pos].get("original_listbox_row_index") != current_pos:
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

    def _get_template_names(self) -> List[str]:
        """Returns a sorted list of template names without extensions."""
        os.makedirs(self.steps_template_directory, exist_ok=True)
        return sorted([os.path.splitext(f)[0] for f in os.listdir(self.steps_template_directory) if f.endswith(".json")])

    def _handle_save_as_template_request(self, step_data_to_save: Dict[str, Any]):
        start_index = step_data_to_save.get("original_listbox_row_index")
        if start_index is None or not (0 <= start_index < len(self.added_steps_data)) or self.added_steps_data[start_index].get("original_listbox_row_index") != start_index:
            try:
                start_index = self.added_steps_data.index(step_data_to_save)
            except ValueError:
                 QMessageBox.critical(self, "Error", "Could not find the step to save as a template. Data may be out of sync.")
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
                return

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
                self.load_all_modules_to_tree()
            except Exception as e:
                QMessageBox.critical(self, "Save Error", f"Failed to save the template:\n{e}")

    def _handle_execute_this_request(self, step_data: Dict[str, Any]):
        """Executes a single step requested from an ExecutionStepCard button."""
        if not (step_data and isinstance(step_data, dict)):
            return

        current_row = step_data.get("original_listbox_row_index")
        if current_row is None or not (0 <= current_row < len(self.added_steps_data)) or self.added_steps_data[current_row].get("original_listbox_row_index") != current_row:
            try:
                current_row = self.added_steps_data.index(step_data)
            except ValueError:
                 QMessageBox.critical(self, "Error", "Could not find the selected step in the internal data model. Data may be out of sync.")
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
            {'type': 'hardcoded', 'value': 0}, # Pass a "no wait" config for single steps
            single_step_mode=True, 
            selected_start_index=current_row
        )
        self._connect_worker_signals()
        self.worker.start()

    def update_status_column_for_all_items(self):
        """Enhanced method to clear all execution status."""
        self._clear_all_execution_status()
        self._clear_status_recursive(self.execution_tree.invisibleRootItem())

    def _clear_status_recursive(self, parent_item: QTreeWidgetItem):
        """Enhanced method to clear visual status in tree."""
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            card = self.execution_tree.itemWidget(child, 0)
            if card:
                card.set_status("#D3D3D3")  # Normal gray border, normal thickness
                card.clear_result()
            self._clear_status_recursive(child)

    def toggle_always_on_top(self) -> None:
        if self.always_on_top_button.isChecked():
            self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
            self.always_on_top_button.setText("📌 Always On Top: On")
        else:
            self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, False)
            self.always_on_top_button.setText("📌 Always On Top: Off")
        self.show()

    def open_screenshot_tool(self, initial_image_name: str = "") -> None:
        self.hide()
        current_image_name_for_tool = self.label_info3.text() if self.label_info3.text() != "Image Name" else ""
        self.screenshot_window = SecondWindow(current_image_name_for_tool, base_dir=self.base_directory, parent=self)
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
            is_template_item = module_item.text(0) == "Bot Templates"
            if is_template_item:
                module_item.setHidden(not (selected_module_name == "-- Show All Modules --"))
            else:
                module_item.setHidden(not (selected_module_name == "-- Show All Modules --" or module_item.text(0) == selected_module_name))
        
        self.module_tree.collapseAll()

    def search_module_tree(self, text: str) -> None:
        """Filters the module tree based on the search text."""
        search_text = text.lower()
        root = self.module_tree.invisibleRootItem()

        def filter_recursive(item: QTreeWidgetItem) -> bool:
            child_matches = False
            for i in range(item.childCount()):
                if filter_recursive(item.child(i)):
                    child_matches = True
            
            item_text = item.text(0).lower()
            item_matches = search_text in item_text
            
            should_be_visible = item_matches or child_matches
            item.setHidden(not should_be_visible)

            if should_be_visible and not item_matches and child_matches:
                 item.setExpanded(True)
            elif not search_text:
                item.setExpanded(False)

            return should_be_visible

        for i in range(root.childCount()):
            filter_recursive(root.child(i))

    def saved_step_tree_item_selected(self, item: QTreeWidgetItem, column: int):
            """Loads a bot's steps when its item is double-clicked in the tree."""
            bot_name = item.text(0)
            if bot_name == "No saved bots found.":
                return
        
            file_path = os.path.join(self.bot_steps_directory, f"{bot_name}.csv")
            if os.path.exists(file_path):
                self.load_steps_from_file(file_path, bot_name)
                workflow_tab_index = self.main_tab_widget.indexOf(self.workflow_scroll_area.parentWidget())
                if self.added_steps_data and self.main_tab_widget.currentIndex() != workflow_tab_index: # Only switch if loading was successful
                    self.main_tab_widget.setCurrentWidget(self.execution_tree.parentWidget())                
                
            else:
                QMessageBox.warning(self, "File Not Found", f"The selected bot file was not found:\n{file_path}")
                self.load_saved_steps_to_tree()

    def _extract_variables_from_steps(self, steps: List[Dict[str, Any]]) -> set:
        """Recursively finds all variable names used in a list of steps."""
        found_variables = set()

        def search_dict(d: Dict[str, Any]):
            if "assign_to_variable_name" in d and d["assign_to_variable_name"]:
                found_variables.add(d["assign_to_variable_name"])
            if "assign_iteration_to_variable" in d and d.get("assign_iteration_to_variable"):
                 found_variables.add(d["assign_iteration_to_variable"])
            
            if d.get("type") == "variable" and "value" in d:
                found_variables.add(d["value"])

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
    
    def _apply_variable_mapping(self, steps: List[Dict[str, Any]], mapping: Dict[str, str]) -> List[Dict[str, Any]]:
        """Recursively replaces variable names in steps based on the provided mapping."""
        steps_copy = json.loads(json.dumps(steps))

        def replace_in_dict(d: Dict[str, Any]):
            if "assign_to_variable_name" in d and d["assign_to_variable_name"] in mapping:
                d["assign_to_variable_name"] = mapping[d["assign_to_variable_name"]]
            if "assign_iteration_to_variable" in d and d.get("assign_iteration_to_variable") in mapping:
                d["assign_iteration_to_variable"] = mapping[d["assign_iteration_to_variable"]]

            if d.get("type") == "variable" and d.get("value") in mapping:
                d["value"] = mapping[d["value"]]

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

    def _load_template_by_name(self, template_name: str) -> None:
        """Loads a template file, handles variable mapping, and inserts it into the execution tree."""
        file_path = os.path.join(self.steps_template_directory, f"{template_name}.json")

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                template_steps = json.load(f)

            if not template_steps:
                QMessageBox.warning(self, "Empty Template", "The selected template is empty.")
                return

            template_variables = self._extract_variables_from_steps(template_steps)
            mapped_steps = template_steps

            if template_variables:
                self._log_to_console(f"Template '{template_name}' contains variables: {template_variables}")
                var_dialog = TemplateVariableMappingDialog(template_variables, list(self.global_variables.keys()), self)
                if var_dialog.exec() == QDialog.DialogCode.Accepted:
                    mapping_result = var_dialog.get_mapping()
                    if mapping_result is None:
                        return 
                
                    variable_map, new_vars = mapping_result
                    
                    mapped_steps = self._apply_variable_mapping(template_steps, variable_map)
                    
                    if new_vars:
                        self.global_variables.update(new_vars)
                        self._update_variables_list_display()
                        self._log_to_console(f"Added new global variables: {list(new_vars.keys())}")
                else:
                    self._log_to_console("Template loading cancelled by user at variable mapping stage.")
                    return

            re_id_d_steps = self._re_id_template_blocks(mapped_steps)

            insertion_dialog = StepInsertionDialog(self.execution_tree, parent=self)
            if insertion_dialog.exec() == QDialog.DialogCode.Accepted:
                selected_tree_item, insert_mode = insertion_dialog.get_insertion_point()
                insert_data_index = self._calculate_flat_insertion_index(selected_tree_item, insert_mode)
                
                self.added_steps_data[insert_data_index:insert_data_index] = re_id_d_steps
                self._sync_counters_with_loaded_data()  # ADD THIS LINE
                self._rebuild_execution_tree(item_to_focus_data=re_id_d_steps[0])
                self._log_to_console(f"Loaded template '{template_name}' with {len(re_id_d_steps)} steps.")
                workflow_tab_index = self.main_tab_widget.indexOf(self.workflow_scroll_area.parentWidget())
                if self.main_tab_widget.currentIndex() != workflow_tab_index:
                    self.main_tab_widget.setCurrentWidget(self.execution_tree.parentWidget())

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
            self.gui_communicator.update_module_info_signal.emit("")
        else:
            pass

    def add_item_to_execution_tree(self, item: QTreeWidgetItem, column: int) -> None:
        item_data = self._get_item_data(item)

        if isinstance(item_data, dict) and item_data.get('type') == 'template':
            template_name = item_data.get('name')
            if template_name:
                self._load_template_by_name(template_name)
            return

        if not (item_data and isinstance(item_data, tuple) and len(item_data) == 5):
            return
            
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
        
        if assign_to_variable_name and assign_to_variable_name not in self.global_variables:
            self.global_variables[assign_to_variable_name] = None
            self._update_variables_list_display()

        new_step_data_dict: Dict[str, Any] = {"type": "step", "class_name": class_name, "method_name": method_name, "module_name": module_name, "parameters_config": parameters_config, "assign_to_variable_name": assign_to_variable_name}
        insertion_dialog = StepInsertionDialog(self.execution_tree, parent=self)
        if insertion_dialog.exec() == QDialog.DialogCode.Accepted:
            selected_tree_item, insert_mode = insertion_dialog.get_insertion_point()
            insert_data_index = self._calculate_flat_insertion_index(selected_tree_item, insert_mode)
            self.added_steps_data.insert(insert_data_index, new_step_data_dict)
            self._rebuild_execution_tree(item_to_focus_data=new_step_data_dict)
            workflow_tab_index = self.main_tab_widget.indexOf(self.workflow_scroll_area.parentWidget())
            if self.main_tab_widget.currentIndex() != workflow_tab_index:
                self.main_tab_widget.setCurrentWidget(self.execution_tree.parentWidget())
        
        self.gui_communicator.update_module_info_signal.emit("")
        self.active_param_input_dialog = None

# In the MainWindow class

    def _calculate_flat_insertion_index(self, selected_tree_item: Optional[QTreeWidgetItem], insert_mode: str) -> int:
        """Enhanced method that always allows before/after insertion."""
        return self._calculate_smart_insertion_index(selected_tree_item, insert_mode)


    def add_loop_block(self) -> None:
        dialog = LoopConfigDialog(self.global_variables, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            loop_config = dialog.get_config()
            if loop_config is None:
                return

            new_iter_var = loop_config.get("assign_iteration_to_variable")
            if new_iter_var and new_iter_var not in self.global_variables:
                self.global_variables[new_iter_var] = None
                self._update_variables_list_display()

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

# In the MainWindow class, REPLACE the existing group_selected_steps method with this one:

    def group_selected_steps(self) -> None:
        selected_items = self.execution_tree.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "No Selection", "Please select one or more steps to group.")
            return

        group_name, ok = QInputDialog.getText(self, "Create Group", "Enter a name for the group:")
        if ok and group_name:
            
            # --- FIX STARTS HERE: Correct logic for multi-selection ---
            selected_data = [self._get_item_data(item) for item in selected_items]
            
            # Find the min and max original_listbox_row_index from the selection
            try:
                indices = [d.get("original_listbox_row_index") for d in selected_data if d and d.get("original_listbox_row_index") is not None]
                if not indices:
                    QMessageBox.critical(self, "Error", "Could not identify the selected items. Data may be out of sync.")
                    return
                
                start_index = min(indices)
                last_item_index_in_selection = max(indices)

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not process selected items for grouping: {e}")
                return

            # Find the full block of the last selected item to ensure we wrap the whole thing
            _, end_index = self._find_block_indices(last_item_index_in_selection)
            # --- FIX ENDS HERE ---

            self.group_id_counter += 1
            group_id = f"group_{self.group_id_counter}"
            
            group_start_data = {"type": "group_start", "group_id": group_id, "group_name": group_name}
            group_end_data = {"type": "group_end", "group_id": group_id}

            # Insert the 'group_start' data before the first selected item
            self.added_steps_data.insert(start_index, group_start_data)
            
            # Insert the 'group_end' data after the last selected item.
            # We must add 1 because we already inserted the start item, shifting all subsequent indices.
            self.added_steps_data.insert(end_index + 2, group_end_data)

            self._rebuild_execution_tree(item_to_focus_data=group_start_data)

    def edit_step_in_execution_tree(self, item: QTreeWidgetItem, column: int) -> None:
        step_data_dict = self._get_item_data(item)
        if not step_data_dict or not isinstance(step_data_dict, dict):
            QMessageBox.warning(self, "Invalid Item", "Cannot edit this item type or no data found.")
            return
        step_type = step_data_dict["type"]
        current_row = -1

        current_row = step_data_dict.get("original_listbox_row_index")
        if current_row is None or not (0 <= current_row < len(self.added_steps_data)) or self.added_steps_data[current_row].get("original_listbox_row_index") != current_row:
            try:
                current_row = self.added_steps_data.index(step_data_dict)
            except ValueError:
                 QMessageBox.critical(self, "Error", "Could not find selected item in internal data model for editing. Data may be out of sync.")
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
                
                if new_assign_to_variable_name and new_assign_to_variable_name not in self.global_variables:
                    self.global_variables[new_assign_to_variable_name] = None
                    self._update_variables_list_display()
                
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

                new_iter_var = new_loop_config.get("assign_iteration_to_variable")
                if new_iter_var and new_iter_var not in self.global_variables:
                    self.global_variables[new_iter_var] = None
                    self._update_variables_list_display()

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
            
            self.execution_flow_label.setText("Execution Flow")
            
            self._log_to_console("Internal clear all steps executed.")
            self._update_workflow_tab(switch_to_tab=False)
            
    def clear_all_steps(self) -> None:
        if not self.added_steps_data and not self.global_variables:
            QMessageBox.information(self, "Info", "The execution queue and variables are already empty.")
            return
        if QMessageBox.question(self, "Confirm Remove All", "Are you sure you want to remove ALL steps and variables?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            self._internal_clear_all_steps()
            self._log_to_console("All steps cleared by user.")

# In the MainWindow class
# In MainWindow class
# REPLACE your existing _handle_execute_pause_resume method with this one:

    def _handle_execute_pause_resume(self):
        """
        Handles Start/Pause/Resume. Now correctly builds the focus mode UI
        including the image preview label.
        """
        if not self.is_bot_running:
            # --- START Case ---
            if not self.added_steps_data:
                QMessageBox.information(self, "No Steps", "No steps have been added.")
                return
            if not self._validate_block_structure_on_execution():
                return

            self.is_bot_running = True
            self.is_paused = False
            self.view_workflow_button.click()
            
            if self.always_on_top_button.isChecked():
                # --- ENTERING FOCUS MODE ---
                self.minimized_for_execution = True
                self.original_geometry = self.geometry()
                
                self.workflow_scroll_area.setWidgetResizable(False)


                # --- Store the original homes for ALL widgets we are moving ---
                def store_home(widget):
                    parent_layout = widget.parent().layout()
                    index = parent_layout.indexOf(widget)
                    self.widget_homes[widget] = (parent_layout, index)

                store_home(self.execute_all_button)
                store_home(self.exit_button)
                store_home(self.workflow_scroll_area)
                store_home(self.label_info2) # Store home for image preview label
                
                # Hide the main side panels and the normal center content
                self.left_menu.setVisible(False)
                self.right_panels.setVisible(False)
                self.normal_mode_widget.setVisible(False)

                # --- Build the Focus Mode UI dynamically ---
                # 1. Add the workflow canvas
                self.focus_mode_layout.addWidget(self.workflow_scroll_area)
                self.workflow_scroll_area.setVisible(True)

                # 2. Add the image preview label
                self.focus_mode_layout.addWidget(self.label_info2)
                self.label_info2.setVisible(True)

                # 3. Add the temporary control widget for buttons
                if not hasattr(self, 'execution_control_widget'):
                    self.execution_control_widget = QWidget()
                    self.execution_control_widget.setLayout(QHBoxLayout())
                    self.execution_control_widget.layout().setContentsMargins(5, 5, 5, 5)
                self.focus_mode_layout.addWidget(self.execution_control_widget)
                self.execution_control_widget.layout().addWidget(self.execute_all_button)
                self.execution_control_widget.layout().addWidget(self.exit_button)
                
                # Make the entire focus mode container visible
                self.focus_mode_widget.setVisible(True)
                
                # Resize window
                screen = QApplication.primaryScreen().geometry()
                new_width = int(self.original_geometry.width() * 0.25)
                new_height = int(self.original_geometry.height() * 0.30)
                self.resize(new_width, new_height)
                self.move(screen.width() - new_width - 10, 10)

            else: # Normal execution
                self.main_tab_widget.setCurrentIndex(2)

            # --- Universal Start Logic ---
            self.execute_all_button.setText("⏸️ Pause")
            self.execute_all_button.setToolTip("Pause the running execution")
            self.set_ui_enabled_state(False)
            self.right_panels.setEnabled(False)
            self.main_tab_widget.setEnabled(False)
            self.progress_bar.setValue(0)
            self.progress_bar.show()
            self.update_status_column_for_all_items()
            
            self.worker = ExecutionWorker(self.added_steps_data, self.module_directory, self.gui_communicator, self.global_variables,self.wait_time_between_steps )
            self._connect_worker_signals()
            self.worker.start()

        elif self.is_bot_running and not self.is_paused:
            # --- PAUSE Case ---
            if self.worker:
                self.worker.pause()
                self.is_paused = True
                self.execute_all_button.setText("▶️ Resume")
                self.execute_all_button.setToolTip("Resume the paused execution")
                
        elif self.is_bot_running and self.is_paused:
            # --- RESUME Case ---
            if self.worker:
                self.worker.resume()
                self.is_paused = False
                self.execute_all_button.setText("⏸️ Pause")
                self.execute_all_button.setToolTip("Pause the running execution")


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
        """Executes a single step selected from the main execution tree."""
        selected_items = self.execution_tree.selectedItems()
        if len(selected_items) != 1:
            QMessageBox.information(self, "Selection Error", "Please select exactly ONE step to execute.")
            return

        # Get the data from the selected tree item
        selected_step_data = self._get_item_data(selected_items[0])
        if not (selected_step_data and isinstance(selected_step_data, dict)):
            QMessageBox.critical(self, "Error", "Selected item has no valid data associated with it.")
            return

        # --- THE FIX ---
        # Instead of searching the list for a copied object, get the reliable index.
        current_row = selected_step_data.get("original_listbox_row_index")

        # Validate the index
        if current_row is None or not (0 <= current_row < len(self.added_steps_data)):
            QMessageBox.critical(self, "Error", "Could not find selected step in internal data model. The data may be out of sync. Please try rebuilding the steps.")
            return
        # --- END FIX ---

        # The rest of the logic is the same
        self.set_ui_enabled_state(False)
        self.progress_bar.setValue(0)
        self.progress_bar.show()
        self.update_status_column_for_all_items()
        
        self.worker = ExecutionWorker(
            self.added_steps_data, 
            self.module_directory, 
            self.gui_communicator, 
            self.global_variables, 
            {'type': 'hardcoded', 'value': 0}, # <-- Pass a "no wait" config dict
            single_step_mode=True, 
            selected_start_index=current_row  # Use the reliable index
        )
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
            # Disconnect the new signal if it exists
            if hasattr(self.worker, 'loop_iteration_started'):
                self.worker.loop_iteration_started.disconnect()
        except TypeError:
            pass
            
        self.worker.execution_started.connect(lambda msg: self._log_to_console(f"GUI: {msg}"))
        self.worker.execution_progress.connect(self.progress_bar.setValue)
        self.worker.execution_item_started.connect(self.update_execution_tree_item_status_started)
        self.worker.execution_item_finished.connect(self.update_execution_tree_item_status_finished)
        self.worker.execution_error.connect(self.update_execution_tree_item_status_error)
        self.worker.execution_finished_all.connect(self.on_execution_finished)
        self.worker.execution_finished_all.connect(lambda: self._update_variables_list_display())
        
        # Connect the new loop iteration signal
        if hasattr(self.worker, 'loop_iteration_started'):
            self.worker.loop_iteration_started.connect(self._on_loop_iteration_started)
            
    def _on_loop_iteration_started(self, loop_id: str, iteration_number: int):
        """Handle loop iteration start - reset status of steps within the loop."""
        if iteration_number > 1:  # Only reset for iterations after the first one
            self._reset_nested_loop_steps_status(loop_id)
    def _find_qtreewidget_item(self, target_step_data_dict: Dict[str, Any], parent_item: Optional[QTreeWidgetItem] = None) -> Optional[QTreeWidgetItem]:
        if parent_item is None:
            parent_item = self.execution_tree.invisibleRootItem()
        
        for i in range(parent_item.childCount()):
            child_item = parent_item.child(i)
            item_data = self._get_item_data(child_item)
            
            if item_data:
                # Match by original_listbox_row_index instead of object equality
                target_index = target_step_data_dict.get("original_listbox_row_index")
                item_index = item_data.get("original_listbox_row_index")
                
                if target_index is not None and target_index == item_index:
                    return child_item
            
            # Recursively search children
            found_in_children = self._find_qtreewidget_item(target_step_data_dict, child_item)
            if found_in_children:
                return found_in_children
        
        return None
    def _find_parent_group_data(self, step_index: int) -> Optional[Dict[str, Any]]:
        """
        Finds the 'group_start' data for the group that the step at step_index is inside.
        Returns None if the step is not inside a group.
        """
        if step_index < 0 or step_index >= len(self.added_steps_data):
            return None
            
        group_stack: List[Dict[str, Any]] = []
        
        # Iterate from the beginning up to (but not including) the step itself
        for i in range(step_index):
            step_data = self.added_steps_data[i]
            step_type = step_data.get("type")
            
            if step_type == "group_start":
                group_stack.append(step_data)
            elif step_type == "group_end":
                if group_stack:
                    # Pop if the ID matches. In a valid structure, it always should.
                    if group_stack[-1].get("group_id") == step_data.get("group_id"):
                        group_stack.pop()
                        
        # If the stack is not empty, the step is inside the group at the top
        if group_stack:
            return group_stack[-1]
            
        return None
        
    def update_execution_tree_item_status_started(self, step_data_dict: Dict[str, Any], original_listbox_row_index: int) -> None:
        """Enhanced method with deferred centering for robust workflow visualization."""
        # --- Part 1: Immediate UI Updates (Borders, Logs, etc.) ---
        
        # Set execution status in the data model
        self._set_step_execution_status(step_data_dict, "running")
        
        # Update parent group status
        parent_group_data = self._find_parent_group_data(original_listbox_row_index)
        if parent_group_data and parent_group_data.get("execution_status") != "error":
            self._set_step_execution_status(parent_group_data, "running")
        
        # Update the execution tree card using index map
        if original_listbox_row_index is not None:
            item_widget = self.data_to_item_map.get(original_listbox_row_index)
            if item_widget:
                card = self.execution_tree.itemWidget(item_widget, 0)
                if card:
                    card.set_status("", is_running=True)
                    self._log_to_console(f"Executing: {card._get_formatted_title()}")
                self.execution_tree.setCurrentItem(item_widget)
                self.execution_tree.scrollToItem(item_widget, QtWidgets.QAbstractItemView.ScrollHint.PositionAtCenter)

        # Tell the canvas it needs a repaint (this will happen after layouts are stable)
        if hasattr(self, 'workflow_canvas') and self.workflow_canvas:
            self.workflow_canvas.update()

        # --- Part 2: Deferred Centering Logic (FIXED) ---
        
        def center_the_view():
            """This function will run AFTER Qt has stabilized the layout."""
            if not (hasattr(self, 'workflow_canvas') and self.workflow_canvas and self.workflow_scroll_area.isVisible()):
                return

            # Find the corresponding node by matching the original_listbox_row_index
            target_node_rect: Optional[QRect] = None
            for rect, _, _, node_step_data in self.workflow_canvas.nodes:
                # FIX: Match by index instead of object identity
                node_index = node_step_data.get("original_listbox_row_index")
                if node_index == original_listbox_row_index:
                    target_node_rect = rect
                    break
            
            if target_node_rect:
                try:
                    # Get the center of the node
                    node_center = target_node_rect.center()
                    
                    # Get the STABLE and CORRECT visible size of the scroll area
                    viewport_size = self.workflow_scroll_area.viewport().size()
                    
                    # Calculate the new top-left scroll value to center the node
                    target_x = node_center.x() - (viewport_size.width() // 2)
                    target_y = node_center.y() - (viewport_size.height() // 2)
                    
                    # Set the scroll bars
                    self.workflow_scroll_area.horizontalScrollBar().setValue(target_x)
                    self.workflow_scroll_area.verticalScrollBar().setValue(target_y)

                except Exception as e:
                    self._log_to_console(f"Workflow scroll error: {e}")

        # THE FIX: Schedule the centering function to run as soon as Qt is idle.
        QTimer.singleShot(0, center_the_view)


    def update_execution_tree_item_status_finished(self, step_data_dict: Dict[str, Any], message: str, original_listbox_row_index: int) -> None:
        """Enhanced method with workflow visualization support."""
        
        #print(f"DEBUG: update_execution_tree_item_status_finished called for step {original_listbox_row_index}")
        
        # --- NEW: Find parent group *before* processing ---
        parent_group_data = self._find_parent_group_data(original_listbox_row_index)
        
        # --- NEW: Check if the *current step* is a group_end ---
        is_group_end = step_data_dict.get("type") == "group_end"
        
        if is_group_end:
            # This is a group_end step. Find its matching group_start.
            group_id = step_data_dict.get("group_id")
            group_start_data = None
            for step in self.added_steps_data: # Search the whole list
                if step.get("type") == "group_start" and step.get("group_id") == group_id:
                    group_start_data = step
                    break
            
            # Set the status of the group_start node
            if group_start_data and group_start_data.get("execution_status") != "error":
                self._set_step_execution_status(group_start_data, "completed")
                # Also set the group_end step itself to completed
                self._set_step_execution_status(step_data_dict, "completed")
        
        elif message == "SKIPPED":
            #print("DEBUG: Step was SKIPPED")
            # --- Step was SKIPPED ---
            self._set_step_execution_status(step_data_dict, "normal")
            if "execution_result" in step_data_dict:
                 del step_data_dict["execution_result"]
            
            # If inside a group, keep the group "running"
            if parent_group_data and parent_group_data.get("execution_status") != "error":
                self._set_step_execution_status(parent_group_data, "running")
            
        else:
            #print("DEBUG: Step COMPLETED normally")
            # --- Step COMPLETED normally (and is not a group_end) ---
            self._set_step_execution_status(step_data_dict, "completed")
            step_data_dict["execution_result"] = message
            #self.label_info1.setText(f"Last Result: {message[0:25]}")
            # If inside a group, keep the group "running"
            if parent_group_data and parent_group_data.get("execution_status") != "error":
                self._set_step_execution_status(parent_group_data, "running")

        # --- Update the Card for the current step (regardless of type) ---
        if message == "SKIPPED":
            #print("DEBUG: Processing SKIPPED message")
            item_widget = self._find_qtreewidget_item(step_data_dict)
            #print(f"DEBUG: item_widget = {item_widget}")
            if item_widget:
                card = self.execution_tree.itemWidget(item_widget, 0)
                #print(f"DEBUG: card = {card}")
                if card:
                    card.set_status("#D3D3D3", is_running=False)
                    card.clear_result()
                self._log_to_console(f"Skipped: {card._get_formatted_title()}")
        else:
            #print("DEBUG: Processing normal completion message")
            item_widget = self._find_qtreewidget_item(step_data_dict)
            #print(f"DEBUG: item_widget = {item_widget}")
            if item_widget:
                card = self.execution_tree.itemWidget(item_widget, 0)
                #print(f"DEBUG: card = {card}")
                if card:
                    #print(f"DEBUG: About to call set_result_text with: {message}")
                    card.set_status("darkGreen", is_running=False)
                    card.set_result_text(message)  # <-- This should call it now
                    #print("DEBUG: set_result_text called successfully")
                self._log_to_console(f"Finished: {card._get_formatted_title()} | {message}")
                self._update_variables_list_display()

    def update_execution_tree_item_status_error(self, step_data_dict: Dict[str, Any], error_message: str, original_listbox_row_index: int) -> None:
        """Enhanced method with workflow visualization support."""
        # Set execution status
        self._set_step_execution_status(step_data_dict, "error")
        
        # --- NEW: Propagate error to parent group ---
        parent_group_data = self._find_parent_group_data(original_listbox_row_index)
        if parent_group_data:
            self._set_step_execution_status(parent_group_data, "error")
        # --- END NEW ---

        # Update the execution tree card
        item_widget = self._find_qtreewidget_item(step_data_dict)
        if item_widget:
            card = self.execution_tree.itemWidget(item_widget, 0)
            if card:
                card.set_status("", is_running=False, is_error=True)  # Red border, thick
            self._log_to_console(f"ERROR on {card._get_formatted_title()}: {error_message}")
            self._update_variables_list_display()

# In the MainWindow class
# In the MainWindow class
# In MainWindow class
# REPLACE your existing on_execution_finished method with this one:

    def on_execution_finished(self, context: ExecutionContext, stopped_by_error: bool, next_step_index_to_select: int) -> None:
        self.is_bot_running = False
        self.is_paused = False
        self.worker = None

        self.execute_all_button.setText("🚀 Execute All")
        self.execute_all_button.setToolTip("Execute all steps from the beginning")
        
        self.progress_bar.setValue(100)
        self.last_executed_context = context

        if self.minimized_for_execution:
            # --- EXITING FOCUS MODE ---
            # Hide the focus mode container
            self.focus_mode_widget.setVisible(False)
            
            self.workflow_scroll_area.setWidgetResizable(True)

            # --- Restore ALL widgets back to their original homes ---
            def restore_home(widget):
                if widget in self.widget_homes:
                    layout, index = self.widget_homes[widget]
                    # To be safe, set the widget's parent to None before re-inserting
                    widget.setParent(None)
                    layout.insertWidget(index, widget)
                    widget.setVisible(True) # Ensure it's visible after being moved

            restore_home(self.execute_all_button)
            restore_home(self.exit_button)
            restore_home(self.workflow_scroll_area)
            restore_home(self.label_info2) # Restore the image preview label
                
            self.widget_homes.clear()

            # Restore visibility of main UI panels
            self.left_menu.setVisible(True)
            self.right_panels.setVisible(True)
            self.normal_mode_widget.setVisible(True) # Show the normal UI container
            
            # Restore window geometry
            if self.original_geometry:
                self.setGeometry(self.original_geometry)

            self.minimized_for_execution = False
        
        # --- Universal UI Re-enablement ---
        self.set_ui_enabled_state(True)
        self.right_panels.setEnabled(True)
        self.main_tab_widget.setEnabled(True)

        if next_step_index_to_select != -1:
            item_to_select = self.data_to_item_map.get(next_step_index_to_select)
            if item_to_select:
                self.execution_tree.setCurrentItem(item_to_select)
                self.execution_tree.scrollToItem(item_to_select, QtWidgets.QAbstractItemView.ScrollHint.PositionAtCenter)

        # Switch to the Execution Flow tab to see the final results
        #self.main_tab_widget.setCurrentIndex(1)
# In the MainWindow class, REPLACE the existing update_label_info_from_module method with this one:
    def update_click_signal_from_module(self, message: str) -> None:
        
        self.label_info1.setText(message)
        
    def update_label_info_from_module(self, message: str) -> None:
        """
        Updates the info labels, including the image preview.
        This is the corrected and robust version.
        """
        # 1. Handle cases where the preview should be cleared.
        if not message or message.startswith("Module Info:") or message.startswith("Last Module Log:"):
            #self.label_info2.clear()
            #self.label_info2.setText("Image Preview")
            #self.label_info3.setText("Image Name")
            if message:
                self.label_info1.setText(message)
            return

        # 2. If we have a message, it's an image name. Construct the full path.
        #    THIS IS THE CRITICAL FIX: We need to build the full path to the .txt file.
        image_name = message.replace("Image not found: ","")
        image_data_filepath = os.path.join(self.click_image_dir, f"{image_name}.txt")

        # 3. Check if the data file actually exists.
        if os.path.exists(image_data_filepath):
            self.label_info3.setText(image_name)
            try:
                # 4. Open the .txt file and load the JSON data.
                with open(image_data_filepath, 'r', encoding='utf-8') as json_file:
                    img_data = json.load(json_file)
                    # The image data is the first value in the dictionary
                    base64_string = next(iter(img_data.values()), None)
                
                if base64_string:
                    # 5. Decode, convert, resize, and display the image.
                    pic_png = self.base64_pgn(base64_string)
                    qimage = ImageQt(pic_png)
                    pixmap = self.resize_qimage_and_create_qpixmap(qimage)
                    self.label_info2.setPixmap(pixmap)
                    #self.label_info1.setText(f"Previewing: {image_name}")
                else:
                    # Handle case where the JSON is valid but has no image data
                    #self.label_info1.setText(f"No image data found in '{image_name}.txt'.")
                    self.label_info2.setText("No Data")
                    self.label_info2.clear()

            except (json.JSONDecodeError, Exception) as e:
                self.label_info1.setText(f"Error loading image '{image_name}': {e}")
                self.label_info2.setText("Load Error")
                self.label_info2.clear()
        else:
            # Handle case where the file is missing entirely.
            #self.label_info1.setText(f"Image file for '{image_name}' not found.")
            timestamp = datetime.now().strftime("%H:%M:%S")
            self.log_console.append(f"[{timestamp}] {message}")
            self.label_info2.setText("Not Found")
            self.label_info2.clear()
            self.label_info3.setText("Image Name")

    def base64_pgn(self,text):
        return PIL.Image.open(io.BytesIO(base64.b64decode(text)))

    def resize_qimage_and_create_qpixmap(self,qimage_input, percentage=98):
        if qimage_input.isNull():
            return QPixmap()
        new_width = int(qimage_input.width() * (percentage / 100))
        new_height = int(qimage_input.height() * (percentage / 100))
        return QPixmap.fromImage(qimage_input.scaled(QSize(new_width, new_height), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))

    def set_ui_enabled_state(self, enabled: bool) -> None:
        widgets_to_toggle = [
            # self.execute_all_button, # <-- REMOVED
            self.add_loop_button, self.add_conditional_button, 
            self.save_steps_button, self.clear_selected_button, self.remove_all_steps_button,
            self.module_filter_dropdown, self.module_tree, self.saved_steps_tree,
            self.add_var_button, self.edit_var_button, self.delete_var_button, 
            self.clear_vars_button, self.open_screenshot_tool_button, 
            self.group_steps_button
        ]  # <--- It's right here
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

    def save_bot_steps_dialog(self) -> None:
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
                variables_to_save = {}
                for var_name, var_value in self.global_variables.items():
                    try:
                        json.dumps(var_value)
                        variables_to_save[var_name] = var_value
                    except (TypeError, OverflowError):
                        variables_to_save[var_name] = None
    
                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(["__GLOBAL_VARIABLES__"])
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
    

    def load_steps_from_file(self, file_path: str, bot_name: str = "") -> None:
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
                            next(reader, None)
                            continue
                        
                        if section == "SCHEDULE":
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
                self._sync_counters_with_loaded_data()  # ADD THIS LINE
                self._rebuild_execution_tree()
                
                if bot_name:
                    self.execution_flow_label.setText(f"Execution Flow: {bot_name}")
                    self.bot_workflow_label.setText(f"Bot Workflow: {bot_name}")
    
                self._log_to_console(f"Loaded bot from {os.path.basename(file_path)}")
    
            except Exception as e:
                QMessageBox.critical(self, "Load Error", f"An unexpected error occurred while loading the file:\n{e}")
                self._log_to_console(f"Load Error: {e}")
                
    def show_context_menu(self, position: QPoint):
        item = self.module_tree.itemAt(position)
        if not item:
            return

        item_data = self._get_item_data(item)
        context_menu = QMenu(self)

        is_top_level = (
            item.parent() == self.module_tree.invisibleRootItem() or
            item.parent() is None
        )

        if is_top_level and item.text(0) != "Bot Templates":
            pass

        elif isinstance(item_data, tuple) and len(item_data) == 5:
            _, class_name, method_name, module_name, _ = item_data
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

        elif isinstance(item_data, dict) and item_data.get('type') == 'template':
            template_name = item_data.get('name')
            if not template_name:
                return

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

    def view_template_documentation(self, template_name: str):
        """Finds and displays the HTML documentation for a given template."""
        os.makedirs(self.template_document_directory, exist_ok=True)
    
        doc_path = os.path.join(self.template_document_directory, f"{template_name}.html")
    
        if os.path.exists(doc_path):
            dialog = HtmlViewerDialog(doc_path, self)
            dialog.exec()
        else:
            QMessageBox.information(self, "Documentation Not Found",
                                      f"No documentation file found at:\n{doc_path}")
    
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
                    self.load_all_modules_to_tree()
                else:
                    QMessageBox.warning(self, "File Not Found", f"The template file for '{template_name}' could not be found.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"An error occurred while deleting the template:\n{e}")

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

    def modify_method(self, module_name: str, class_name: str, method_name: str):
        try:
            module_path = os.path.join(self.module_directory, f"{module_name}.py")
            if not os.path.exists(module_path):
                QMessageBox.critical(self, "File Not Found", f"The source file for module '{module_name}' could not be found.")
                return
    
            editor_dialog = CodeEditorDialog(module_path, class_name, method_name, self.all_parsed_method_data, self)
            editor_dialog.exec()
            
            self.load_all_modules_to_tree()
    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not open the code editor: {e}")
    
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
                    self._handle_execute_pause_resume() # This will now set the is_bot_running flag

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

    def _read_schedule_from_csv(self, file_path: str) -> Optional[Dict[str, Any]]:
        """Reads only the schedule info from a bot's CSV file."""
        if not os.path.exists(file_path):
            return None
        try:
            with open(file_path, 'r', newline='', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                for row in reader:
                    if row and row[0] == "__SCHEDULE_INFO__":
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
            if os.path.exists(file_path):
                with open(file_path, 'r', newline='', encoding='utf-8') as f:
                    lines = list(csv.reader(f))

            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                try:
                    schedule_header_index = [i for i, row in enumerate(lines) if row and row[0] == "__SCHEDULE_INFO__"][0]
                    lines[schedule_header_index + 1] = [json.dumps(schedule_data)]
                    schedule_written = True
                except IndexError:
                    pass
                
                if schedule_written:
                    writer.writerows(lines)
                else:
                    writer.writerow(["__SCHEDULE_INFO__"])
                    writer.writerow([json.dumps(schedule_data)])
                    writer.writerows(lines)
            return True
        except Exception as e:
            self._log_to_console(f"Error writing schedule to {file_path}: {e}")
            return False
            
    def _log_to_console(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_console.append(f"[{timestamp}] {message}")

    def update_application(self):
        github_zip_url = "https://github.com/tuanhungstar/automateclick/archive/refs/heads/main.zip"
        self.update_dir = os.path.join(self.base_directory, "update")
        zip_path = os.path.join(self.update_dir, "update.zip")

        os.makedirs(self.update_dir, exist_ok=True)

        try:
            with urllib.request.urlopen(github_zip_url) as response, open(zip_path, 'wb') as out_file:
                shutil.copyfileobj(response, out_file)
            self._process_zip_file(zip_path)

        except urllib.error.URLError:
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Icon.Critical)
            msg_box.setWindowTitle("Download Failed")
            msg_box.setTextFormat(Qt.TextFormat.RichText)
            msg_box.setText(
                "Could not download the update automatically.<br><br>"
                "<b>Step 1:</b> Please download the file manually from this link:<br>"
                f"<a href='{github_zip_url}'>{github_zip_url}</a><br><br>"
                "<b>Step 2:</b> After the download is complete, click the <b>'Downloaded'</b> button below and select the file you just saved."
            )
            downloaded_button = msg_box.addButton("Downloaded", QMessageBox.ButtonRole.ActionRole)
            msg_box.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
            msg_box.exec()

            if msg_box.clickedButton() == downloaded_button:
                file_path, _ = QFileDialog.getOpenFileName(self, "Select Downloaded File", "", "Zip Files (*.zip)")
                if file_path:
                    self._process_zip_file(file_path)

        except Exception as e:
            QMessageBox.critical(self, "Update Error", f"An error occurred: {e}")

    def _process_zip_file(self, zip_path):
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(self.update_dir)

            extracted_folder_name = ""
            for item in os.listdir(self.update_dir):
                if os.path.isdir(os.path.join(self.update_dir, item)) and "automateclick" in item:
                    extracted_folder_name = item
                    break

            if not extracted_folder_name:
                raise Exception("Could not find the main folder in the downloaded zip.")

            extracted_path = os.path.join(self.update_dir, extracted_folder_name)
            dialog = UpdateDialog(extracted_path, self.base_directory, self)
            dialog.exec()
        
        except Exception as e:
            QMessageBox.critical(self, "Update Error", f"An error occurred while processing the file: {e}")

    def _update_workflow_tab(self, switch_to_tab: bool = False) -> None:
        """Builds/Refreshes the workflow canvas. Optionally switches to its tab."""
        
        old_canvas = self.workflow_scroll_area.takeWidget()
        if old_canvas:
            # Disconnect old signals before deleting to prevent memory leaks
            try:
                old_canvas.execute_step_requested.disconnect()
            except (TypeError, AttributeError):
                pass # Ignore if it was not connected or doesn't exist
            old_canvas.deleteLater()

        if not self.added_steps_data:
            self.workflow_canvas = WorkflowCanvas([], self)
            self.workflow_scroll_area.setWidget(self.workflow_canvas)
            if switch_to_tab:
                self.main_tab_widget.setCurrentWidget(self.workflow_scroll_area.parentWidget())
            return

        try:
            workflow_tree = self._build_workflow_tree_data()
            if not workflow_tree:
                if switch_to_tab:
                    QMessageBox.warning(self, "Error", "Could not parse the workflow structure.")
                return
            
            self.workflow_canvas = WorkflowCanvas(workflow_tree, self)
            
            # --- CONNECT THE NEW SIGNAL HERE ---
            self.workflow_canvas.execute_step_requested.connect(self._handle_execute_this_request)
            # -----------------------------------
            
            self.workflow_scroll_area.setWidget(self.workflow_canvas)
            
            if switch_to_tab:
                self.main_tab_widget.setCurrentWidget(self.workflow_scroll_area.parentWidget())
            
        except Exception as e:
            if switch_to_tab:
                QMessageBox.critical(self, "Workflow Error", f"An error occurred while building the workflow: {e}")
            self._log_to_console(f"Workflow build error: {e}")

    def _build_workflow_tree_data(self) -> List[Dict[str, Any]]:
        """
        Parses the flat self.added_steps_data into a nested tree structure
        that supports IF/ELSE and LOOP/END branching. [CORRECTED VERSION]
        """
        def parse_block_recursive(flat_steps: List[Dict[str, Any]], index: int) -> Tuple[List[Dict[str, Any]], int]:
            nodes = []
            i = index
            while i < len(flat_steps):
                step_data = flat_steps[i]
                step_type = step_data.get("type")

                # --- Exit Condition: Stop parsing when we hit an end or else tag ---
                if step_type in ["group_end", "loop_end", "IF_END", "ELSE"]:
                    return nodes, i

                new_node = {'step_data': step_data, 'children': []}
                i += 1 # Move to the next step

                # --- Recursive Parsing for Block Types ---
                if step_type == "group_start":
                    # Parse children until a 'group_end' is found
                    children, end_index = parse_block_recursive(flat_steps, i)
                    new_node['children'] = children
                    i = end_index

                    # Consume the 'group_end' tag
                    if i < len(flat_steps) and flat_steps[i].get("type") == "group_end":
                        i += 1 # Move past the 'group_end' tag
                    nodes.append(new_node)
                    continue

                elif step_type == "IF_START" or step_type == "loop_start":
                    # Parse the main ("true") branch
                    children, next_index = parse_block_recursive(flat_steps, i)
                    new_node['children'] = children
                    i = next_index

                    # --- Correctly Handle IF-ELSE structure ---
                    if step_type == "IF_START" and i < len(flat_steps) and flat_steps[i].get("type") == "ELSE":
                        # The 'ELSE' is NOT a child. We record its steps in 'false_children'.
                        new_node['false_children'] = []
                        i += 1 # Consume the 'ELSE' tag

                        # Parse the 'else' branch until an 'IF_END' is found
                        false_children, end_index = parse_block_recursive(flat_steps, i)
                        # We store the actual steps, not the ELSE tag itself
                        new_node['false_children'] = false_children
                        i = end_index

                    # --- Find and attach the End Node ---
                    end_tag = "IF_END" if step_type == "IF_START" else "loop_end"
                    if i < len(flat_steps) and flat_steps[i].get("type") == end_tag:
                        # Attach the end node data to the start node
                        new_node['end_node'] = {'step_data': flat_steps[i]}
                        i += 1 # Consume the 'end_tag'

                # --- Add the fully constructed node (step, if, loop, or group) ---
                nodes.append(new_node)

            return nodes, i

        # Start the parsing from the beginning of the flat list
        root_nodes, _ = parse_block_recursive(self.added_steps_data, 0)
        return root_nodes

    def _get_image_filenames(self) -> List[str]:
            """Gets a sorted list of image names, including relative paths from subfolders."""
            relative_paths: List[str] = []
            if not os.path.exists(self.click_image_dir):
                return []
                
            for root, _, files in os.walk(self.click_image_dir):
                for filename in files:
                    if filename.lower().endswith(".txt"):
                        full_path = os.path.join(root, filename)
                        # Get the path relative to click_image_dir
                        relative_path = os.path.relpath(full_path, self.click_image_dir)
                        
                        # If the file is in the root, relpath is "filename.txt"
                        # If in subfolder, it's "Subfolder/filename.txt"
                        
                        # Remove the .txt extension
                        relative_path_no_ext = os.path.splitext(relative_path)[0]
                        
                        # Standardize on '/' for path separators
                        relative_paths.append(relative_path_no_ext.replace(os.sep, '/'))
                            
            return sorted(relative_paths)
        
    def _set_step_execution_status(self, step_data: Dict[str, Any], status: str):
            """Sets the execution status for a step in both tree and workflow views."""
            
            if status == "normal":
                if "execution_status" in step_data:
                    del step_data["execution_status"]
            else:
                step_data["execution_status"] = status
            
            # Update the workflow canvas if it exists
            if hasattr(self, 'workflow_canvas') and self.workflow_canvas:
                self.workflow_canvas.update()
    
    def _clear_all_execution_status(self):
        """Clears execution status from all steps."""
        for step_data in self.added_steps_data:
            if "execution_status" in step_data:
                del step_data["execution_status"]
            if "execution_result" in step_data:  # ADD THIS LINE
                del step_data["execution_result"]  # ADD THIS LINE
        
        # Update the workflow canvas
        if hasattr(self, 'workflow_canvas') and self.workflow_canvas:
            self.workflow_canvas.update()
            
    def _reset_loop_steps_status(self, loop_id: str):
        """Resets the execution status of all steps within a specific loop."""
        if not loop_id:
            return
            
        # Find the loop boundaries
        loop_start_index = -1
        loop_end_index = -1
        
        for i, step_data in enumerate(self.added_steps_data):
            if (step_data.get("type") == "loop_start" and 
                step_data.get("loop_id") == loop_id):
                loop_start_index = i
            elif (step_data.get("type") == "loop_end" and 
                  step_data.get("loop_id") == loop_id and 
                  loop_start_index != -1):
                loop_end_index = i
                break
        
        if loop_start_index == -1 or loop_end_index == -1:
            return
            
        # Reset status for all steps within the loop (excluding the loop_start and loop_end)
        for i in range(loop_start_index + 1, loop_end_index):
            step_data = self.added_steps_data[i]
            if "execution_status" in step_data:
                del step_data["execution_status"]
            
            # Also reset the visual status in the execution tree
            item_widget = self._find_qtreewidget_item(step_data)
            if item_widget:
                card = self.execution_tree.itemWidget(item_widget, 0)
                if card:
                    card.set_status("#D3D3D3")  # Normal gray border
                    card.clear_result()
        
        # Update the workflow canvas to reflect the changes
        if hasattr(self, 'workflow_canvas') and self.workflow_canvas:
            self.workflow_canvas.update()
            
        self._log_to_console(f"🔄 Loop '{loop_id}' iteration started - Reset status of {loop_end_index - loop_start_index - 1} steps within the loop")

    def _reset_nested_loop_steps_status(self, loop_id: str):
        """Resets execution status for steps in a loop, handling nested structures properly."""
        if not loop_id:
            return
            
        # Find all steps that belong to this loop (including nested structures)
        loop_steps = []
        current_loop_level = 0
        inside_target_loop = False
        
        for i, step_data in enumerate(self.added_steps_data):
            step_type = step_data.get("type")
            step_loop_id = step_data.get("loop_id")
            step_if_id = step_data.get("if_id")
            step_group_id = step_data.get("group_id")
            
            # Check if we're entering our target loop
            if step_type == "loop_start" and step_loop_id == loop_id:
                inside_target_loop = True
                current_loop_level = 0
                continue
                
            # Check if we're exiting our target loop
            if step_type == "loop_end" and step_loop_id == loop_id and inside_target_loop:
                break
                
            # If we're inside the target loop, track nesting and collect steps
            if inside_target_loop:
                # Track nesting level for other structures
                if step_type in ["loop_start", "IF_START", "group_start"]:
                    current_loop_level += 1
                elif step_type in ["loop_end", "IF_END", "group_end"]:
                    current_loop_level -= 1
                
                # Add this step to our reset list
                loop_steps.append((i, step_data))
        
        # Reset status for all collected steps
        reset_count = 0
        for step_index, step_data in loop_steps:
            if "execution_status" in step_data:
                del step_data["execution_status"]
                reset_count += 1
            if "execution_result" in step_data:  # ADD THIS LINE
                del step_data["execution_result"]  # ADD THIS LINE
            # Also reset the visual status in the execution tree
            item_widget = self._find_qtreewidget_item(step_data)
            if item_widget:
                card = self.execution_tree.itemWidget(item_widget, 0)
                if card:
                    card.set_status("#D3D3D3")  # Normal gray border
                    card.clear_result()
        
        # Update the workflow canvas to reflect the changes
        if hasattr(self, 'workflow_canvas') and self.workflow_canvas:
            self.workflow_canvas.update()
            
        if reset_count > 0:
            self._log_to_console(f"🔄 Loop '{loop_id}' new iteration - Reset status of {reset_count} steps within the loop")
            
    def _handle_step_drag_started(self, step_data: Dict[str, Any], original_index: int):
        """Handle when a step card starts being dragged."""
        self._log_to_console(f"Drag started for step at index {original_index}")

    def _handle_step_reorder(self, source_index: int, target_index: int):
        """Handle reordering of steps via drag and drop."""
        if source_index == target_index or source_index == -1 or target_index == -1:
            return
        
        # Clamp target_index to valid range
        target_index = max(0, min(target_index, len(self.added_steps_data)))
        
        # Perform the reordering
        self._reorder_steps_smart(source_index, target_index)

    def _reorder_steps_smart(self, source_index: int, target_index: int):
        """Smart reordering that respects block boundaries."""
        if not (0 <= source_index < len(self.added_steps_data)):
            return
        
        # Get the block boundaries for the source step
        source_start, source_end = self._find_block_indices(source_index)
        
        # Calculate the actual target position
        if target_index > source_end:
            # Moving down - adjust target to account for removed items
            actual_target = target_index - (source_end - source_start + 1)
        else:
            actual_target = target_index
        
        # Ensure target is within bounds
        actual_target = max(0, min(actual_target, len(self.added_steps_data) - (source_end - source_start + 1)))
        
        # Extract the block to move
        steps_to_move = self.added_steps_data[source_start:source_end + 1]
        
        # Remove from original position
        del self.added_steps_data[source_start:source_end + 1]
        
        # Insert at new position
        for i, step in enumerate(steps_to_move):
            self.added_steps_data.insert(actual_target + i, step)
        
        # Rebuild the tree and focus on the moved block
        self._rebuild_execution_tree(item_to_focus_data=steps_to_move[0])
        
        self._log_to_console(f"Moved step block from index {source_index} to {actual_target}")
    '''
    def _calculate_smart_insertion_index(self, selected_tree_item: Optional[QTreeWidgetItem], insert_mode: str) -> int:
        """Enhanced insertion calculation that properly handles block structures."""
        if selected_tree_item is None:
            # If no item is selected, insert at the end of the entire list.
            return len(self.added_steps_data)

        selected_item_data = self._get_item_data(selected_tree_item)
        if not selected_item_data:
            # If the item has no associated data, treat as inserting at the end.
            return len(self.added_steps_data)

        try:
            selected_flat_index = self.added_steps_data.index(selected_item_data)
        except ValueError as e:
            error_content = str(e)
            self._log_to_console(f"ValueError in insertion index calculation: {error_content}")
            print ("error:", str(ValueError))
            # If the selected item's data is not found in the flat list, insert at the end.
            return len(self.added_steps_data)
        
        selected_step_type = selected_item_data.get("type")

        if insert_mode == "before":
            return selected_flat_index
        elif insert_mode == "after":
            return selected_flat_index + 1
        
        # Fallback, should not be reached with "before" or "after" modes.
        return len(self.added_steps_data)
        '''
    def _calculate_smart_insertion_index(self, selected_tree_item: Optional[QTreeWidgetItem], insert_mode: str) -> int:
        """Enhanced insertion calculation that properly handles block structures."""
        if selected_tree_item is None:
            # If no item is selected, insert at the end of the entire list.
            return len(self.added_steps_data)

        selected_item_data = self._get_item_data(selected_tree_item)
        if not selected_item_data:
            # If the item has no associated data, treat as inserting at the end.
            return len(self.added_steps_data)

        # Use original_listbox_row_index for more reliable lookup
        selected_flat_index = selected_item_data.get("original_listbox_row_index")
        
        # If original_listbox_row_index is not available or invalid, try to find by identity
        if selected_flat_index is None or not (0 <= selected_flat_index < len(self.added_steps_data)):
            try:
                # Try to find by object identity first (fastest)
                for i, step_data in enumerate(self.added_steps_data):
                    if step_data is selected_item_data:
                        selected_flat_index = i
                        break
                else:
                    # If identity search fails, try equality comparison
                    selected_flat_index = self.added_steps_data.index(selected_item_data)
            except ValueError as e:
                error_content = str(e)
                self._log_to_console(f"ValueError in insertion index calculation: {error_content}")
                # If the selected item's data is not found in the flat list, insert at the end.
                return len(self.added_steps_data)
        
        # Verify the index is still valid
        if not (0 <= selected_flat_index < len(self.added_steps_data)):
            return len(self.added_steps_data)
            
        selected_step_type = selected_item_data.get("type")

        if insert_mode == "before":
            return selected_flat_index
        elif insert_mode == "after":
            return selected_flat_index + 1
        
        # Fallback, should not be reached with "before" or "after" modes.
        return len(self.added_steps_data)
    # In the MainWindow class, add this new method
# In the MainWindow class, REPLACE the existing edit_step_from_data method with this one.

    def edit_step_from_data(self, step_data: Dict[str, Any]):
        """
        Finds a step by its unique index and triggers the edit dialog.
        This is called from the WorkflowCanvas context menu.
        This is the robust version that prevents the 'out of sync' error.
        """
        # 1. Get the unique identifier for the step from the workflow node's data.
        step_index = step_data.get("original_listbox_row_index")

        # 2. Basic validation.
        if step_index is None:
            QMessageBox.warning(self, "Edit Error", "The selected workflow shape has no valid identifier.")
            return

        # 3. Check if the index is valid within our main data list.
        if not (0 <= step_index < len(self.added_steps_data)):
            QMessageBox.warning(self, "Edit Error", f"The step index '{step_index}' is out of bounds. The workflow may be out of sync. Please try rebuilding the flow.")
            return

        # 4. Find the corresponding item in the Execution Flow tree using the data_to_item_map.
        #    This map was created specifically for this purpose.
        item_to_edit = self.data_to_item_map.get(step_index)

        # 5. If the item is found, call the same edit function as the card's "Edit" button.
        if item_to_edit:
            # This now behaves exactly like clicking the "Edit" button on the card.
            self.edit_step_in_execution_tree(item_to_edit, 0)
        else:
            # This is a fallback and should rarely happen if the map is up-to-date.
            QMessageBox.warning(self, "Edit Error", "Could not find the corresponding UI element in the Execution Flow. Please try rebuilding the flow.")

# In MainWindow class, add this new method:

    def open_rearrange_steps_dialog(self):
        """Opens the dialog for reordering steps."""
        if not self.added_steps_data:
            QMessageBox.information(self, "No Steps", "There are no steps to rearrange.")
            return
            
        # Create an instance of our new dialog with the current steps
        dialog = RearrangeStepsDialog(self.added_steps_data, self)
        
        # If the user clicks "OK"
        if dialog.exec() == QDialog.DialogCode.Accepted:
            # Get the newly ordered list of steps from the dialog
            new_order = dialog.get_rearranged_steps()
            
            # Check if the order has actually changed
            if new_order != self.added_steps_data:
                self.added_steps_data = new_order
                
                # Rebuild the main execution tree to reflect the new order
                self._rebuild_execution_tree()
                
                self._log_to_console("Steps have been rearranged.")
            else:
                self._log_to_console("Step order remains unchanged.")
                
# In MainWindow class, REPLACE the set_wait_time method with this:

    def set_wait_time(self):
        """Opens an advanced dialog to set the wait time between execution steps."""
        
        # Create an instance of our new, more advanced dialog
        dialog = WaitTimeConfigDialog(
            global_variables=list(self.global_variables.keys()),
            initial_config=self.wait_time_between_steps,
            parent=self
        )
        
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_config = dialog.get_config()
            if new_config:
                self.wait_time_between_steps = new_config
                
                # Update button text to reflect the new setting
                if new_config['type'] == 'variable':
                    display_text = f"@{new_config['value']}"
                    self._log_to_console(f"Wait time will be determined by global variable '{display_text}'.")
                else:
                    display_text = f"{new_config['value']}s"
                    self._log_to_console(f"Wait time between steps set to {display_text}.")
                
                self.set_wait_time_button.setText(f"⏱️ Wait between Steps ({display_text})")
                    
    def _sync_counters_with_loaded_data(self):
        """Synchronizes the ID counters with the highest IDs found in loaded data."""
        max_loop_id = 0
        max_if_id = 0
        max_group_id = 0
        
        for step_data in self.added_steps_data:
            # Check loop IDs
            if "loop_id" in step_data:
                loop_id_str = step_data["loop_id"]
                if isinstance(loop_id_str, str) and loop_id_str.startswith("loop_"):
                    try:
                        loop_num = int(loop_id_str.split("_")[1])
                        max_loop_id = max(max_loop_id, loop_num)
                    except (IndexError, ValueError):
                        pass
            
            # Check IF IDs
            if "if_id" in step_data:
                if_id_str = step_data["if_id"]
                if isinstance(if_id_str, str) and if_id_str.startswith("if_"):
                    try:
                        if_num = int(if_id_str.split("_")[1])
                        max_if_id = max(max_if_id, if_num)
                    except (IndexError, ValueError):
                        pass
            
            # Check group IDs
            if "group_id" in step_data:
                group_id_str = step_data["group_id"]
                if isinstance(group_id_str, str) and group_id_str.startswith("group_"):
                    try:
                        group_num = int(group_id_str.split("_")[1])
                        max_group_id = max(max_group_id, group_num)
                    except (IndexError, ValueError):
                        pass
        
        # Update the counters to be the maximum found (next increment will be max+1)
        self.loop_id_counter = max_loop_id
        self.if_id_counter = max_if_id
        self.group_id_counter = max_group_id
        
        self._log_to_console(f"Synced counters - Loop: {self.loop_id_counter}, IF: {self.if_id_counter}, Group: {self.group_id_counter}")
if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Apply a modern stylesheet
    app.setStyleSheet("""
        QMainWindow, QDialog {
            background-color: #f8f9fa;
        }
        QTabWidget::pane {
            border-top: 1px solid #dee2e6;
        }
        QTabBar::tab {
            background: #e9ecef;
            border: 1px solid #dee2e6;
            border-bottom-color: #dee2e6;
            border-top-left-radius: 4px;
            border-top-right-radius: 4px;
            min-width: 8ex;
            padding: 8px 12px;
            margin-right: 2px;
            font-size: 13px;
        }
        QTabBar::tab:selected, QTabBar::tab:hover {
            background: #ffffff;
        }
        QTabBar::tab:selected {
            border-color: #dee2e6;
            border-bottom-color: #ffffff;
        }
        QSplitter::handle {
            background: #ced4da;
        }
        QSplitter::handle:horizontal {
            width: 3px;
        }
        QSplitter::handle:vertical {
            height: 3px;
        }
        QGroupBox {
            font-weight: bold;
            border: 1px solid #ced4da;
            border-radius: 4px;
            margin-top: 1ex;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            subcontrol-position: top left;
            padding: 0 3px;
        }
        QTreeWidget {
            border: 1px solid #ced4da;
            background-color: #ffffff;
        }
        QListWidget {
            border: 1px solid #ced4da;
            background-color: #ffffff;
        }
        QLineEdit, QTextEdit, QComboBox {
            border: 1px solid #ced4da;
            border-radius: 4px;
            padding: 5px;
            background-color: #ffffff;
        }
        QPushButton {
            background-color: #007bff;
            color: white;
            border: none;
            padding: 8px 12px;
            border-radius: 4px;
            font-size: 13px;
        }
        QPushButton:hover {
            background-color: #0056b3;
        }
        QPushButton:pressed {
            background-color: #004085;
        }
        QPushButton:disabled {
            background-color: #c0c0c0;
            color: #6c757d;
        }
        QLabel#section-header {
            font-weight: bold;
            color: #343a40;
            padding: 4px 0;
            border-bottom: 1px solid #e0e0e0;
            margin-bottom: 4px;
        }
    """)
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
