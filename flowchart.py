import sys
import os
import inspect
import importlib
import uuid
import json
import csv
import time
import ast
# Ensure my_lib is in the Python path
script_dir = os.path.dirname(os.path.abspath(__file__))
my_lib_dir = os.path.join(script_dir, "my_lib")
if my_lib_dir not in sys.path:
    sys.path.insert(0, my_lib_dir)
    
from datetime import datetime # ADD THIS LINE
from typing import Optional, List, Any, Dict, Tuple
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QGraphicsView, QGraphicsScene, QWidget,
    QInputDialog, QGraphicsPolygonItem, QGraphicsTextItem,
    QMessageBox, QPushButton, QVBoxLayout, QHBoxLayout, QFrame, QDialog,
    QTreeWidget, QTreeWidgetItem, QDialogButtonBox, QSplitter, QGroupBox,
    QListWidget, QTextEdit, QProgressBar, QCheckBox, QLineEdit, QFormLayout,
    QFileDialog, QComboBox, QRadioButton, QLabel, QListWidgetItem,
    QGridLayout, QTreeWidgetItemIterator, QGraphicsPathItem, QHeaderView
)
from PyQt6.QtGui import QPolygonF, QBrush, QPen, QFont, QColor, QPainterPath
from PyQt6.QtCore import QPointF, Qt, QRectF, QVariant, pyqtSignal, QLineF, QDateTime, QThread, QObject
from my_lib.shared_context import ExecutionContext, GuiCommunicator


# --- Configuration ---
STEP_WIDTH = 250
STEP_HEIGHT = 120
DECISION_WIDTH = 280
DECISION_HEIGHT = 100
VERTICAL_SPACING = 80
HORIZONTAL_SPACING = 300
GRID_SIZE = 20
LINE_CLEARANCE = 10 

# Define constants for file management (assuming 'Bot_steps' is a folder next to 'flowchart.py')
BOT_STEPS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Bot_steps")
MODULE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Bot_module")

# --- Style Configuration ---
BG_COLOR = "#FFFFFF"
LEFT_PANEL_COLOR = "#E0E0E0"
CANVAS_COLOR = "#FFFFFF"
BUTTON_COLOR = "#DCDCDC"
BUTTON_TEXT_COLOR = "#000000"
SHAPE_FILL_COLOR = "#FFFFFF"
SHAPE_BORDER_COLOR = "#333333"
LINE_COLOR = "#333333"
HIGHLIGHT_COLOR = "#0078D7"
TRUE_BRANCH_COLOR_BOX = QColor("#E6F2FF")
TRUE_BRANCH_COLOR_LINE = QColor("#007BFF")
FALSE_BRANCH_COLOR_BOX = QColor("#F8D7DA")
FALSE_BRANCH_COLOR_LINE = QColor("#DC3545")

# --- EXECUTION DEPENDENCIES (Based on main_app.py and my_lib/shared_context.py) ---

# --- EXECUTION WORKER (Based on main_app.py ExecutionWorker) ---
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
        else:
            value = operand_config["value"]
            # Attempt to parse basic types saved as strings (e.g., in flowchart.py context)
            if isinstance(value, str):
                try:
                    return ast.literal_eval(value)
                except (ValueError, SyntaxError):
                    return value
            return value

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

    # In flowchart.py: ExecutionWorker.run

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
                    # --- Robust Key Access ---
                    class_name = step_data.get("class_name") or step_data.get("class")
                    method_name = step_data.get("method_name") or step_data.get("method")
                    module_name = step_data.get("module_name")
                    
                    parameters_config = step_data.get("parameters_config") or step_data.get("config")
                    assign_to_variable_name = step_data.get("assign_to_variable_name") or step_data.get("assign_to")

                    if not all([class_name, method_name, module_name]):
                        raise KeyError(f"Missing essential execution keys (class={class_name}, method={method_name}, module={module_name}).")
                    # --- End Robust Key Access ---
                    
                    resolved_parameters, params_str_debug = {}, []
                    for param_name, config in parameters_config.items():
                        if config['type'] == 'hardcoded': resolved_parameters[param_name] = self._resolve_operand_value(config); params_str_debug.append(f"{param_name}={repr(config['value'])}")
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
                        # Find end index
                        loop_end_index = -1; nesting_level = 0
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

# --- SAVE BOT DIALOGS (Unchanged) ---
class SaveBotDialog(QDialog):
    # ... (Unchanged from previous turn)
    def __init__(self, existing_bots: list, parent=None):
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
        # Simple sanitization to prevent issues with file paths
        sanitized_name = "".join(c for c in name if c.isalnum() or c in (' ', '_', '-')).rstrip()
        if not sanitized_name:
            QMessageBox.warning(self, "Invalid Name", "Bot name cannot be empty.")
            return None
        return sanitized_name
# --- BOT LOADER DIALOG (Unchanged) ---
class BotLoaderDialog(QDialog):
    # ... (Unchanged from previous turn)
    bot_selected = pyqtSignal(str)

    def __init__(self, bot_steps_dir: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Open Your Bot")
        self.setGeometry(200, 200, 700, 500)
        self.bot_steps_directory = bot_steps_dir
        # Assuming Schedules is a sibling folder to BOT_STEPS_DIR's parent
        self.schedules_dir = os.path.join(os.path.dirname(os.path.dirname(self.bot_steps_directory)), "Schedules")
        self.schedules = {}

        main_layout = QVBoxLayout(self)
        
        # UI Setup for Saved Bots (similar to main_app.py)
        self.saved_steps_tree = QTreeWidget()
        self.saved_steps_tree.setHeaderLabels(["Bot Name", "Schedule", "Status"])
        self.saved_steps_tree.header().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        # Connect double-click event
        self.saved_steps_tree.itemDoubleClicked.connect(self._open_selected_bot)
        main_layout.addWidget(self.saved_steps_tree)

        button_layout = QHBoxLayout()
        self.open_button = QPushButton("Open Bot")
        self.open_button.clicked.connect(self._open_selected_bot)
        self.schedule_button = QPushButton("Schedule")
        self.delete_button = QPushButton("Delete")
        self.close_button = QPushButton("Close")
        
        button_layout.addWidget(self.open_button)
        button_layout.addWidget(self.schedule_button)
        button_layout.addWidget(self.delete_button)
        button_layout.addStretch()
        button_layout.addWidget(self.close_button)

        main_layout.addLayout(button_layout)
        
        self.close_button.clicked.connect(self.reject)
        
        self.loadSavedBotsToWTreeWidget()

    # New method to handle both double-click and button click
    def _open_selected_bot(self):
        selected_item = self.saved_steps_tree.currentItem()
        if not selected_item:
            QMessageBox.information(self, "No Selection", "Please select a bot to open.")
            return
        
        bot_name = selected_item.data(0, Qt.ItemDataRole.UserRole)
        
        if not bot_name: # Handle the "No saved bots found" placeholder item
            QMessageBox.information(self, "Invalid Selection", "Please select a valid bot.")
            return
            
        # Emit signal to FlowchartApp
        self.bot_selected.emit(bot_name)
        self.accept() # Close the dialog

    # Helper function to read schedule from CSV (minimal implementation based on main_app.py)
    def _read_schedule_from_csv(self, file_path: str) -> Optional[Dict[str, Any]]:
        """Reads only the schedule info from a bot's CSV file."""
        if not os.path.exists(file_path): return None
        try:
            with open(file_path, 'r', newline='', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                for row in reader:
                    if row and row[0] == "__SCHEDULE_INFO__":
                        schedule_row = next(reader, None)
                        if schedule_row:
                            return json.loads(schedule_row[0])
            return None
        except:
            return None

    # Helper function to load all schedules from JSON (minimal implementation for fallback)
    def load_schedules(self):
        """Loads schedules from the schedules.json file."""
        schedule_file_path = os.path.join(self.schedules_dir, "schedules.json")
        if os.path.exists(schedule_file_path):
            try:
                with open(schedule_file_path, 'r', encoding='utf-8') as f:
                    self.schedules = json.load(f)
            except:
                self.schedules = {}
        else:
            self.schedules = {}

    # The function requested by the user: load Saved Bots to WTreeWidget
    def loadSavedBotsToWTreeWidget(self) -> None:
        """Loads saved bot step files and their schedules into the QTreeWidget."""
        self.saved_steps_tree.clear()
        self.load_schedules() # Load schedules from schedules.json (if used)
        try:
            os.makedirs(self.bot_steps_directory, exist_ok=True)
            step_files = sorted([f for f in os.listdir(self.bot_steps_directory) if f.endswith(".csv")], reverse=True)
            
            for file_name in step_files:
                bot_name = os.path.splitext(file_name)[0]
                file_path = os.path.join(self.bot_steps_directory, file_name)
                
                # Try to get schedule info from CSV (new main_app format)
                schedule_info = self._read_schedule_from_csv(file_path)

                schedule_str = "Not Set"
                status_str = "Idle"
                
                if schedule_info:
                    # Logic based on main_app.py for displaying date/repeat
                    start_datetime_obj = QDateTime.fromString(schedule_info.get('start_datetime'), Qt.DateFormat.ISODate)
                    if start_datetime_obj.isValid():
                         schedule_str = f"{schedule_info.get('repeat', 'Once')} at {start_datetime_obj.toString('yyyy-MM-dd hh:mm')}"
                    else:
                         schedule_str = f"{schedule_info.get('repeat', 'Once')}"
                    status_str = "Scheduled" if schedule_info.get("enabled") else "Disabled"
                
                # Check for schedules in schedules.json (old main_app format fallback)
                elif bot_name in self.schedules:
                    schedule_info_json = self.schedules[bot_name]
                    start_datetime_obj = QDateTime.fromString(schedule_info_json.get('start_datetime'), Qt.DateFormat.ISODate)
                    if start_datetime_obj.isValid():
                         schedule_str = f"{schedule_info_json.get('repeat', 'Once')} at {start_datetime_obj.toString('yyyy-MM-dd hh:mm')}"
                    else:
                         schedule_str = f"{schedule_info_json.get('repeat', 'Once')}"
                    status_str = "Scheduled" if schedule_info_json.get("enabled") else "Disabled"


                tree_item = QTreeWidgetItem(self.saved_steps_tree, [bot_name, schedule_str, status_str])
                # Store the bot name in the item's UserRole for retrieval on selection/double-click
                tree_item.setData(0, Qt.ItemDataRole.UserRole, bot_name)
                
            if not step_files:
                self.saved_steps_tree.addTopLevelItem(QTreeWidgetItem(["No saved bots found."]))
                
        except Exception as e:
            QMessageBox.critical(self, "Error Loading Saved Bots", f"Could not load bot files: {e}")

# --- OTHER DIALOGS (Unchanged) ---
class GlobalVariableDialog(QDialog):
    # ... (Unchanged from previous turn)
    def __init__(self, variable_name="", variable_value="", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add/Edit Global Variable")
        self.layout = QFormLayout(self)
        self.name_input = QLineEdit(variable_name)
        self.value_input = QLineEdit(str(variable_value))
        self.browse_button = QPushButton("Browse...")
        self.browse_button.clicked.connect(self._open_file_dialog)
        self.layout.addRow("Name:", self.name_input)
        value_layout = QHBoxLayout()
        value_layout.addWidget(self.value_input)
        value_layout.addWidget(self.browse_button)
        self.layout.addRow("Value:", value_layout)
        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        self.layout.addRow(self.buttons)
        self.name_input.textChanged.connect(self._toggle_browse_button)
        self._toggle_browse_button()

    def _toggle_browse_button(self):
        self.browse_button.setVisible("link" in self.name_input.text().lower())

    def _open_file_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "All Files (*)")
        if file_path:
            self.value_input.setText(file_path)

    def get_variable_data(self):
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "Input Error", "Variable name cannot be empty.")
            return None, None
        return name, self.value_input.text()

class ConditionalConfigDialog(QDialog):
    # ... (Unchanged from previous turn)
    def __init__(self, global_variables, parent=None, initial_config=None):
        super().__init__(parent)
        self.setWindowTitle("Configure Conditional Block (IF-ELSE)")
        self.setMinimumWidth(400)
        self.global_variables = global_variables
        main_layout = QVBoxLayout(self)
        form_layout = QFormLayout()
        self.block_name_editor = QLineEdit()
        self.block_name_editor.setPlaceholderText("Optional: Enter a name for this conditional block")
        form_layout.addRow("Block Name:", self.block_name_editor)
        condition_group = QGroupBox("Condition")
        condition_layout = QGridLayout()
        self.left_operand_source_combo = QComboBox()
        self.left_operand_source_combo.addItems(["Hardcoded Value", "Global Variable"])
        self.left_operand_editor = QLineEdit()
        self.left_operand_var_combo = QComboBox()
        self.left_operand_var_combo.addItem("-- Select Variable --")
        self.left_operand_var_combo.addItems(sorted(global_variables.keys()))
        self.left_operand_source_combo.currentIndexChanged.connect(self._toggle_left_operand_input)
        condition_layout.addWidget(QLabel("Left Operand:"), 0, 0)
        condition_layout.addWidget(self.left_operand_source_combo, 0, 1)
        condition_layout.addWidget(self.left_operand_editor, 0, 2)
        condition_layout.addWidget(self.left_operand_var_combo, 0, 2)
        self.operator_combo = QComboBox()
        self.operator_combo.addItems(['==', '!=', '<', '>', '<=', '>=', 'in', 'not in', 'is', 'is not'])
        condition_layout.addWidget(QLabel("Operator:"), 1, 0)
        condition_layout.addWidget(self.operator_combo, 1, 1, 1, 2)
        self.right_operand_source_combo = QComboBox()
        self.right_operand_source_combo.addItems(["Hardcoded Value", "Global Variable"])
        self.right_operand_editor = QLineEdit()
        self.right_operand_var_combo = QComboBox()
        self.right_operand_var_combo.addItem("-- Select Variable --")
        self.right_operand_var_combo.addItems(sorted(global_variables.keys()))
        self.right_operand_source_combo.currentIndexChanged.connect(self._toggle_right_operand_input)
        condition_layout.addWidget(QLabel("Right Operand:"), 2, 0)
        condition_layout.addWidget(self.right_operand_source_combo, 2, 1)
        condition_layout.addWidget(self.right_operand_editor, 2, 2)
        condition_layout.addWidget(self.right_operand_var_combo, 2, 2)
        condition_group.setLayout(condition_layout)
        main_layout.addWidget(condition_group)
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        main_layout.addLayout(form_layout)
        main_layout.addWidget(button_box)
        self._toggle_left_operand_input()
        self._toggle_right_operand_input()

    def _toggle_left_operand_input(self):
        is_using_var = (self.left_operand_source_combo.currentIndex() == 1)
        self.left_operand_editor.setVisible(not is_using_var)
        self.left_operand_var_combo.setVisible(is_using_var)

    def _toggle_right_operand_input(self):
        is_using_var = (self.right_operand_source_combo.currentIndex() == 1)
        self.right_operand_editor.setVisible(not is_using_var)
        self.right_operand_var_combo.setVisible(is_using_var)

    def _parse_value(self, value_str):
        try:
            return json.loads(value_str)
        except (json.JSONDecodeError, TypeError):
            return value_str

    def get_config(self):
        block_name = self.block_name_editor.text().strip()
        left_op_config = {}
        if self.left_operand_source_combo.currentIndex() == 1:
            var_name = self.left_operand_var_combo.currentText()
            if var_name == "-- Select Variable --":
                QMessageBox.warning(self, "Input Error", "Please select a variable for the left operand.")
                return None
            left_op_config = {"type": "variable", "value": var_name}
        else:
            left_op_config = {"type": "hardcoded", "value": self._parse_value(self.left_operand_editor.text())}

        right_op_config = {}
        if self.right_operand_source_combo.currentIndex() == 1:
            var_name = self.right_operand_var_combo.currentText()
            if var_name == "-- Select Variable --":
                QMessageBox.warning(self, "Input Error", "Please select a variable for the right operand.")
                return None
            right_op_config = {"type": "variable", "value": var_name}
        else:
            right_op_config = {"type": "hardcoded", "value": self._parse_value(self.right_operand_editor.text())}

        return {"block_name": block_name or None, "condition": {"left_operand": left_op_config, "operator": self.operator_combo.currentText(), "right_operand": right_op_config}}

class ParameterInputDialog(QDialog):
    # ... (Unchanged from previous turn - Note: This is now just a dummy, the real one is complex)
    def __init__(self, method_name, parameters_to_configure, current_global_var_names, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Configure Parameters for '{method_name}'")
        self.parameters_config = {}
        self.assign_to_variable_name = None
        self.param_editors = {}
        self.param_var_selectors = {}
        self.param_value_source_combos = {}

        main_layout = QVBoxLayout(self)
        form_layout = QFormLayout()

        for param_name, (default_value, _) in parameters_to_configure.items():
            param_h_layout = QHBoxLayout()
            label = QLabel(f"{param_name}:")

            value_source_combo = QComboBox()
            value_source_combo.addItems(["Hardcoded Value", "Global Variable"])
            self.param_value_source_combos[param_name] = value_source_combo

            hardcoded_editor = QLineEdit()
            if default_value is not inspect.Parameter.empty:
                hardcoded_editor.setText(str(default_value))
            self.param_editors[param_name] = hardcoded_editor

            variable_select_combo = QComboBox()
            variable_select_combo.addItem("-- Select Variable --")
            variable_select_combo.addItems(current_global_var_names)
            self.param_var_selectors[param_name] = variable_select_combo

            param_h_layout.addWidget(value_source_combo)
            param_h_layout.addWidget(hardcoded_editor)
            param_h_layout.addWidget(variable_select_combo)
            
            if "link" in param_name.lower():
                browse_button = QPushButton("Browse...")
                browse_button.clicked.connect(lambda _, editor=hardcoded_editor: self._open_file_dialog_for_param(editor))
                param_h_layout.addWidget(browse_button)

            value_source_combo.currentIndexChanged.connect(
                lambda index, editor=hardcoded_editor, selector=variable_select_combo: self._toggle_param_input_type(index, editor, selector)
            )
            self._toggle_param_input_type(0, hardcoded_editor, variable_select_combo)

            form_layout.addRow(label, param_h_layout)

        main_layout.addLayout(form_layout)

        assignment_group_box = QGroupBox("Assign Method Result to Variable")
        assignment_layout = QVBoxLayout()
        self.assign_checkbox = QCheckBox("Assign result")
        self.assign_checkbox.stateChanged.connect(self._toggle_assignment_widgets)
        assignment_layout.addWidget(self.assign_checkbox)
        self.new_var_radio = QRadioButton("New Variable Name:")
        self.new_var_input = QLineEdit()
        self.existing_var_radio = QRadioButton("Existing Variable:")
        self.existing_var_combo = QComboBox()
        self.existing_var_combo.addItem("-- Select Variable --")
        self.existing_var_combo.addItems(current_global_var_names)
        
        self.new_var_radio.toggled.connect(self._toggle_assignment_inputs)

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

    def _open_file_dialog_for_param(self, editor):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "All Files (*)")
        if file_path:
            editor.setText(file_path)

    def _toggle_param_input_type(self, index, hardcoded_editor, variable_select_combo):
        hardcoded_editor.setVisible(index == 0)
        variable_select_combo.setVisible(index == 1)

    def _toggle_assignment_widgets(self):
        is_assign_enabled = self.assign_checkbox.isChecked()
        self.new_var_radio.setVisible(is_assign_enabled)
        self.new_var_input.setVisible(is_assign_enabled)
        self.existing_var_radio.setVisible(is_assign_enabled)
        self.existing_var_combo.setVisible(is_assign_enabled)
        if is_assign_enabled:
            self.new_var_radio.setChecked(True)

    def _toggle_assignment_inputs(self):
        self.new_var_input.setVisible(self.new_var_radio.isChecked())
        self.existing_var_combo.setVisible(not self.new_var_radio.isChecked())

    def get_parameters_config(self):
        for param_name, source_combo in self.param_value_source_combos.items():
            if source_combo.currentIndex() == 0:
                self.parameters_config[param_name] = {'type': 'hardcoded', 'value': self.param_editors[param_name].text()}
            else:
                var_name = self.param_var_selectors[param_name].currentText()
                if var_name == "-- Select Variable --":
                    QMessageBox.warning(self, "Input Error", f"Please select a variable for '{param_name}'.")
                    return None, None
                self.parameters_config[param_name] = {'type': 'variable', 'value': var_name}
        
        assign_to = None
        if self.assign_checkbox.isChecked():
            if self.new_var_radio.isChecked():
                assign_to = self.new_var_input.text().strip()
                if not assign_to:
                    QMessageBox.warning(self, "Input Error", "New variable name cannot be empty.")
                    return None, None
            else:
                assign_to = self.existing_var_combo.currentText()
                if assign_to == "-- Select Variable --":
                    QMessageBox.warning(self, "Input Error", "Please select an existing variable to assign to.")
                    return None, None

        return self.parameters_config, assign_to

class StepInsertionDialog(QDialog):
    # ... (Unchanged from previous turn)
    def __init__(self, flowchart_items, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert Step At...")
        self.layout = QVBoxLayout(self)
        self.list_widget = QListWidget()
        self.insertion_mode_group = QGroupBox("Insertion Mode")
        self.insertion_mode_layout = QHBoxLayout()
        self.insert_before_radio = QRadioButton("Insert Before Selected")
        self.insert_after_radio = QRadioButton("Insert After Selected")
        self.insert_after_radio.setChecked(True)
        self.insertion_mode_layout.addWidget(self.insert_before_radio)
        self.insertion_mode_layout.addWidget(self.insert_after_radio)
        self.insertion_mode_group.setLayout(self.insertion_mode_layout)

        self.layout.addWidget(QLabel("Select an existing step:"))
        self.layout.addWidget(self.list_widget)
        self.layout.addWidget(self.insertion_mode_group)
        
        # NEW: Create a map for quick lookup to check item type
        self.item_map = {item.step_id: item for item in flowchart_items}

        self._populate_list(flowchart_items)

        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.layout.addWidget(self.button_box)
        
        # NEW: Connect signal to update insertion options based on selection
        self.list_widget.currentItemChanged.connect(self._update_insertion_mode_options)
        self._update_insertion_mode_options() # Initial call

    def _populate_list(self, items):
        if not items:
            self.list_widget.addItem("End of Flowchart")
            self.insert_before_radio.setEnabled(False)
            self.insert_after_radio.setChecked(True)
        else:
            for item in items:
                text = item.text_item.toPlainText().split('\n')[0]
                list_item = QListWidgetItem(text)
                list_item.setData(Qt.ItemDataRole.UserRole, item.step_id)
                self.list_widget.addItem(list_item)
            self.list_widget.setCurrentRow(len(items) - 1)

    # NEW METHOD: Update radio buttons based on selected item type
    def _update_insertion_mode_options(self):
        selected_list_item = self.list_widget.currentItem()
        if not selected_list_item:
            self.insert_before_radio.setEnabled(False)
            self.insert_after_radio.setEnabled(False)
            return

        target_item_id = selected_list_item.data(Qt.ItemDataRole.UserRole)

        # Handle 'End of Flowchart' item which has no UserRole data
        if not target_item_id:
            self.insert_before_radio.setEnabled(False)
            self.insert_after_radio.setEnabled(True)
            self.insert_after_radio.setChecked(True)
            return
            
        target_item = self.item_map.get(target_item_id)
        
        # Check if the target item is a DecisionItem (diamond shape/IF_START)
        is_decision_item = isinstance(target_item, DecisionItem)
        
        if is_decision_item:
            # If a DecisionItem is selected, only allow inserting BEFORE it.
            self.insert_after_radio.setEnabled(False)
            self.insert_before_radio.setEnabled(True)
            self.insert_before_radio.setChecked(True) # Force to "Insert Before"
        else:
            # For all other items, allow both before and after.
            self.insert_after_radio.setEnabled(True)
            self.insert_before_radio.setEnabled(True)

    def get_insertion_point(self):
        selected_list_item = self.list_widget.currentItem()
        if not selected_list_item or selected_list_item.text() == "End of Flowchart":
            return None, "after"

        target_item_id = selected_list_item.data(Qt.ItemDataRole.UserRole)
        mode = "before" if self.insert_before_radio.isChecked() else "after"
        return target_item_id, mode

class ModuleSelectionDialog(QDialog):
    # ... (Unchanged from previous turn)
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select a Method")
        self.setMinimumSize(400, 500)
        self.selected_method_info = None

        main_layout = QVBoxLayout(self)

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search for methods...")
        self.search_box.textChanged.connect(self._filter_tree)
        main_layout.addWidget(self.search_box)

        self.module_tree = QTreeWidget()
        self.module_tree.setHeaderLabels(["Module/Class/Method"])
        self.module_tree.itemDoubleClicked.connect(self.accept)
        main_layout.addWidget(self.module_tree)

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        main_layout.addWidget(button_box)

        self.load_modules()

    def _filter_tree(self):
        search_text = self.search_box.text().lower()
        iterator = QTreeWidgetItemIterator(self.module_tree)
        while iterator.value():
            item = iterator.value()
            item_text = item.text(0).lower()
            
            is_match = search_text in item_text

            item.setHidden(not is_match)
            
            if is_match:
                parent = item.parent()
                while parent:
                    parent.setHidden(False)
                    parent.setExpanded(True)
                    parent = parent.parent()
            iterator += 1

    def load_modules(self):
        base_directory = os.path.dirname(os.path.abspath(__file__))
        module_directory = os.path.join(base_directory, "Bot_module")

        if not os.path.exists(module_directory):
            QMessageBox.warning(self, "Module Folder Not Found",
                                f"The 'Bot_module' folder was not found at:\n{module_directory}")
            return

        original_sys_path = sys.path[:]
        if module_directory not in sys.path:
            sys.path.insert(0, module_directory)

        try:
            for module_name in [f[:-3] for f in os.listdir(module_directory) if f.endswith(".py") and f != "__init__.py"]:
                module = importlib.import_module(module_name)
                importlib.reload(module)
                module_item = QTreeWidgetItem(self.module_tree, [module_name])

                for name, obj in inspect.getmembers(module, inspect.isclass):
                    if obj.__module__ == module_name:
                        class_item = QTreeWidgetItem(module_item, [name])
                        for method_name, method_obj in inspect.getmembers(obj, inspect.isfunction):
                             if not method_name.startswith('_'):
                                sig = inspect.signature(method_obj)
                                params = {p.name: (p.default, p.kind) for p in sig.parameters.values() if p.name != 'self'}
                                method_item = QTreeWidgetItem(class_item, [method_name])
                                method_item.setData(0, Qt.ItemDataRole.UserRole, (module_name, name, method_name, params))
        finally:
            sys.path = original_sys_path


    def accept(self):
        selected_item = self.module_tree.currentItem()
        if selected_item and selected_item.childCount() == 0:
            self.selected_method_info = selected_item.data(0, Qt.ItemDataRole.UserRole)
            super().accept()
        elif not selected_item:
             QMessageBox.warning(self, "No Selection", "Please select a method.")
        else:
            QMessageBox.warning(self, "Invalid Selection", "Please select a method, not a class or module.")


# --- Custom Graphics Items (Unchanged) ---
class Connector(QGraphicsPathItem):
    # ... (Unchanged from previous turn)
    def __init__(self, start_item, end_item, label="", color=QColor(LINE_COLOR), parent=None):
        super().__init__(parent)
        self.start_item = start_item
        self.end_item = end_item
        self.setPen(QPen(color, 2))
        self.setZValue(-1)

        self.label = QGraphicsTextItem(label, self)
        self.label.setDefaultTextColor(color)
        self.label.setFont(QFont("Arial", 10, QFont.Weight.Bold))

    def get_connection_points(self):
        start_rect = self.start_item.boundingRect()
        end_rect = self.end_item.boundingRect()

        is_true_branch = self.label.toPlainText() == "True"
        is_false_branch = self.label.toPlainText() == "False"
        
        # Determine START point in Scene coordinates
        if isinstance(self.start_item, DecisionItem):
            if is_true_branch:
                # Left side of diamond
                start_point_local = QPointF(-DECISION_WIDTH/2, 0)
            elif is_false_branch:
                # Right side of diamond
                start_point_local = QPointF(DECISION_WIDTH/2, 0)
            else:
                # Bottom point of diamond
                start_point_local = QPointF(0, DECISION_HEIGHT/2)
        else:
            # Bottom center of rectangle
            start_point_local = QPointF(0, start_rect.height() / 2)
        start_point_scene = self.start_item.mapToScene(start_point_local)

        # Determine END point in Scene coordinates
        if isinstance(self.end_item, DecisionItem) and not (is_true_branch or is_false_branch):
            # Top point of diamond (default entry)
            end_point_local = QPointF(0, -DECISION_HEIGHT/2)
        else:
            # Top center of rectangle/step
            end_point_local = QPointF(0, -end_rect.height() / 2)
        end_point_scene = self.end_item.mapToScene(end_point_local)
        
        return start_point_scene, end_point_scene, is_true_branch, is_false_branch

    def update_position(self):
        start_point, end_point, is_true_branch, is_false_branch = self.get_connection_points()
        
        path = QPainterPath()
        path.moveTo(start_point)
        
        # Determine if this is a connection from the side of a DecisionItem
        is_side_branch_from_decision = isinstance(self.start_item, DecisionItem) and (is_true_branch or is_false_branch)
        
        if is_side_branch_from_decision:
            # --- Smart Horizontal-Vertical-Horizontal (HVH) Route for Decision Branches ---
            
            h_clearance = max(DECISION_WIDTH / 2, self.end_item.boundingRect().width() / 2) + LINE_CLEARANCE
            x_offset = h_clearance * (1 if is_false_branch else -1)
            
            waypoint_x_1 = start_point.x() + x_offset
            
            if (is_false_branch and waypoint_x_1 < end_point.x()) or \
               (is_true_branch and waypoint_x_1 > end_point.x()):
                waypoint_x_1 = end_point.x()
                
            path.lineTo(waypoint_x_1, start_point.y())
            
            waypoint_y_2 = end_point.y()
            path.lineTo(waypoint_x_1, waypoint_y_2)
            
            path.lineTo(end_point)
            
            label_x = start_point.x() + (-40 if is_true_branch else 10)
            label_y = start_point.y() + 5
            
        else:
            # --- Smart Vertical-Horizontal-Vertical (VHV) Route for Standard Connections ---
            
            dy = end_point.y() - start_point.y()
            
            if abs(dy) > (STEP_HEIGHT + LINE_CLEARANCE):
                # If shapes are far apart, use a 3-segment path (V-H-V)
                mid_y = start_point.y() + dy / 2
                path.lineTo(start_point.x(), mid_y)
                path.lineTo(end_point.x(), mid_y)
            else:
                # If shapes are close, use a 5-segment path for better separation
                
                vertical_clearance_start = (self.start_item.boundingRect().height() / 2) + LINE_CLEARANCE
                vertical_clearance_end = (self.end_item.boundingRect().height() / 2) + LINE_CLEARANCE

                waypoint_x_1 = start_point.x()
                waypoint_y_1 = start_point.y() + vertical_clearance_start 
                
                is_mostly_vertical = abs(start_point.x() - end_point.x()) < GRID_SIZE
                
                if is_mostly_vertical:
                    mid_y = (start_point.y() + end_point.y()) / 2
                    path.lineTo(start_point.x(), mid_y)
                    path.lineTo(end_point.x(), mid_y)
                else:
                    path.lineTo(waypoint_x_1, waypoint_y_1)

                    waypoint_x_2 = (start_point.x() + end_point.x()) / 2
                    path.lineTo(waypoint_x_2, waypoint_y_1)

                    waypoint_y_3 = end_point.y() - vertical_clearance_end
                    if waypoint_y_3 < waypoint_y_1:
                        waypoint_y_3 = max(waypoint_y_1, end_point.y() + vertical_clearance_start)
                        
                    path.lineTo(waypoint_x_2, waypoint_y_3)

                    waypoint_x_4 = end_point.x()
                    path.lineTo(waypoint_x_4, waypoint_y_3)
                
            path.lineTo(end_point)

            label_x = (start_point.x() + end_point.x()) / 2 - 20
            label_y = (start_point.y() + end_point.y()) / 2 - 10
            
        self.setPath(path)
        self.label.setPos(label_x, label_y)

class FlowchartItem(QGraphicsPolygonItem):
    # ... (Unchanged from previous turn)
    def __init__(self, polygon, text="Default", step_data=None, parent=None):
        super().__init__(polygon, parent)
        self.step_id = str(uuid.uuid4())
        self.step_data = step_data if step_data is not None else {}
        self.connectors = []
        self.setBrush(QBrush(QColor(SHAPE_FILL_COLOR)))
        self.setPen(QPen(QColor(SHAPE_BORDER_COLOR), 2))
        self.setAcceptHoverEvents(True)
        self.setFlag(QGraphicsPolygonItem.GraphicsItemFlag.ItemIsSelectable)
        self.setFlag(QGraphicsPolygonItem.GraphicsItemFlag.ItemIsMovable)
        self.setFlag(QGraphicsPolygonItem.GraphicsItemFlag.ItemSendsGeometryChanges)

        self.text_item = QGraphicsTextItem(text, self)
        font = QFont("Arial", 10)
        self.text_item.setDefaultTextColor(QColor(BUTTON_TEXT_COLOR))
        self.text_item.setFont(font)
        self.text_item.setTextWidth(self.boundingRect().width() * 0.9)
        
        self.set_step_details(1)

    def set_step_details(self, step_number, status=""):
        prefix = ""
        if self.step_data.get('if_branch') == 'true':
            prefix = "IF_TRUE: "
        elif self.step_data.get('if_branch') == 'false':
            prefix = "IF_FALSE: "

        display_text = f"Step {step_number} {prefix}: "
        
        if self.step_data.get('type') == 'IF_START':
            cond = self.step_data['condition_config']['condition']
            left = cond['left_operand']['value']
            if cond['left_operand']['type'] == 'variable': left = f"@{left}"
            right = cond['right_operand']['value']
            if cond['right_operand']['type'] == 'variable': right = f"@{right}"
            display_text = f"Step {step_number}: IF: {left} {cond['operator']} {right}\n"
        elif self.step_data.get('type') == 'IF_END':
            display_text = f"Step {step_number}: END IF\n"
        else:
            method_name = self.step_data.get('class', '') + "." + self.step_data.get('method', '')
            if method_name != ".":
                display_text += f"{method_name}\n"
        
        params = self.step_data.get('config', {})
        if params:
            for i, (param_name, param_config) in enumerate(params.items()):
                value_display = str(param_config.get('value', ''))
                if param_config.get('type') == 'variable':
                    value_display = f"@{value_display}"
                if len(value_display) > 20:
                    value_display = value_display[:17] + "..."
                display_text += f"Param {i+1}: {param_name}={value_display}\n"
        
        assign_to = self.step_data.get('assign_to')
        if assign_to:
            display_text += f"Assign: {assign_to}\n"
        
        display_text += f"Status: {status}\n"

        self.text_item.setPlainText(display_text.strip())
        self.text_item.setTextWidth(self.boundingRect().width() * 0.9)
        self.center_text()
        
        if self.step_data.get('if_branch') == 'true':
            self.setBrush(QBrush(TRUE_BRANCH_COLOR_BOX))
            self.setPen(QPen(TRUE_BRANCH_COLOR_LINE, 2, Qt.PenStyle.DashLine))
        elif self.step_data.get('if_branch') == 'false':
            self.setBrush(QBrush(FALSE_BRANCH_COLOR_BOX))
            self.setPen(QPen(FALSE_BRANCH_COLOR_LINE, 2, Qt.PenStyle.DashLine))
        else:
            self.setBrush(QBrush(QColor(SHAPE_FILL_COLOR)))
            self.setPen(QPen(QColor(SHAPE_BORDER_COLOR), 2, Qt.PenStyle.SolidLine))


    def center_text(self):
        self.text_item.setPos(-self.boundingRect().width() / 2 * 0.95, -self.boundingRect().height() / 2 * 0.95)

    def add_connector(self, connector):
        self.connectors.append(connector)

    def itemChange(self, change, value):
        if change == QGraphicsPolygonItem.GraphicsItemChange.ItemPositionChange:
            # Snap the item position to the grid
            new_pos = value
            x = round(new_pos.x() / GRID_SIZE) * GRID_SIZE
            y = round(new_pos.y() / GRID_SIZE) * GRID_SIZE
            return QPointF(x, y)
        
        if change == QGraphicsPolygonItem.GraphicsItemChange.ItemPositionHasChanged:
            # This is the crucial line that triggers the connector update when the shape moves
            for connector in self.connectors:
                connector.update_position()
            # Also update connectors of items connected to this one's output
            for item in self.scene().items():
                if isinstance(item, Connector) and item.end_item == self:
                    item.update_position()
        return super().itemChange(change, value)

    def mouseDoubleClickEvent(self, event):
        QMessageBox.information(None, "Edit Step", "Double-click logic for editing step details is not yet implemented.")
        super().mouseDoubleClickEvent(event)

    def hoverEnterEvent(self, event):
        temp_pen = self.pen()
        temp_pen.setColor(QColor(HIGHLIGHT_COLOR))
        temp_pen.setWidth(3)
        self.setPen(temp_pen)
        super().hoverEnterEvent(event)

    def hoverLeaveEvent(self, event): # FIX: Added 'event' argument
        if self.step_data.get('if_branch') == 'true':
            self.setPen(QPen(TRUE_BRANCH_COLOR_LINE, 2, Qt.PenStyle.DashLine))
        elif self.step_data.get('if_branch') == 'false':
            self.setPen(QPen(FALSE_BRANCH_COLOR_LINE, 2, Qt.PenStyle.DashLine))
        else:
            self.setPen(QPen(QColor(SHAPE_BORDER_COLOR), 2, Qt.PenStyle.SolidLine))
        super().hoverLeaveEvent(event)

class StepItem(FlowchartItem):
    # ... (Unchanged from previous turn)
    def __init__(self, text="Step", step_data=None, parent=None):
        rect = QPolygonF([
            QPointF(-STEP_WIDTH / 2, -STEP_HEIGHT / 2), QPointF(STEP_WIDTH / 2, -STEP_HEIGHT / 2),
            QPointF(STEP_WIDTH / 2, STEP_HEIGHT / 2), QPointF(-STEP_WIDTH / 2, STEP_HEIGHT / 2),
        ])
        super().__init__(rect, text, step_data, parent)

class DecisionItem(FlowchartItem):
    # ... (Unchanged from previous turn)
    def __init__(self, text="If...", step_data=None, parent=None):
        diamond = QPolygonF([
            QPointF(0, -DECISION_HEIGHT / 2), QPointF(DECISION_WIDTH / 2, 0),
            QPointF(0, DECISION_HEIGHT / 2), QPointF(-DECISION_WIDTH / 2, 0),
        ])
        super().__init__(diamond, text, step_data, parent)

# --- Custom QGraphicsView (Unchanged) ---
class FlowchartView(QGraphicsView):
    # ... (Unchanged from previous turn)
    def __init__(self, scene, parent=None):
        super().__init__(scene, parent)
        self.main_window = parent
        self.setStyleSheet(f"background-color: {CANVAS_COLOR}; border: 1px solid #C0C0C0;")
        self.temp_line = None
        self.setRenderHint(self.renderHints().Antialiasing)
        self.setDragMode(QGraphicsView.DragMode.NoDrag)

    # Override drawBackground to draw the grid
    def drawBackground(self, painter, rect):
        super().drawBackground(painter, rect)

        grid_size = GRID_SIZE
        light_gray = QColor("#E8E8E8")
        pen = QPen(light_gray)
        pen.setWidth(0) # hairline

        painter.setPen(pen)

        left = int(rect.left())
        right = int(rect.right())
        top = int(rect.top())
        bottom = int(rect.bottom())

        # Draw vertical lines
        x = left - (left % grid_size)
        while x < right:
            painter.drawLine(x, top, x, bottom)
            x += grid_size

        # Draw horizontal lines
        y = top - (top % grid_size)
        while y < bottom:
            painter.drawLine(left, y, right, y)
            y += grid_size

    def mousePressEvent(self, event):
        if self.main_window.connection_mode and event.button() == Qt.MouseButton.LeftButton:
            scene_pos = self.mapToScene(event.pos())
            item = self.scene().itemAt(scene_pos, self.transform())
            
            # Find the FlowchartItem if we clicked on a text or the shape itself
            flowchart_item = None
            if isinstance(item, FlowchartItem):
                flowchart_item = item
            elif isinstance(item, QGraphicsTextItem) and isinstance(item.parentItem(), FlowchartItem):
                flowchart_item = item.parentItem()
            
            if flowchart_item:
                if self.main_window.start_item is None:
                    # First click - select start item
                    self.main_window.start_item = flowchart_item
                    pen = flowchart_item.pen()
                    pen.setColor(QColor(HIGHLIGHT_COLOR))
                    pen.setWidth(3)
                    flowchart_item.setPen(pen)
                    
                    # Create temporary line for visual feedback
                    self.temp_line = QGraphicsPathItem()
                    self.temp_line.setPen(QPen(QColor(HIGHLIGHT_COLOR), 2, Qt.PenStyle.DashLine))
                    self.scene().addItem(self.temp_line)
                else:
                    # Second click - connect to end item
                    end_item = flowchart_item
                    if self.main_window.start_item != end_item:
                        # Determine label based on connection type
                        label = ""
                        color = QColor(LINE_COLOR)
                        
                        # If the starting item is a DecisionItem, ask which branch
                        if isinstance(self.main_window.start_item, DecisionItem):
                            branch_choice, ok = QInputDialog.getItem(self, "Branch Selection", "Select branch type:", ["Default (Bottom)", "True (Left)", "False (Right)"], 0, False)
                            if not ok:
                                # User cancelled, reset start item visual
                                pen = self.main_window.start_item.pen()
                                pen.setColor(QColor(SHAPE_BORDER_COLOR))
                                pen.setWidth(2)
                                self.main_window.start_item.setPen(pen)
                                self.main_window.start_item = None
                                if self.temp_line:
                                    self.scene().removeItem(self.temp_line)
                                    self.temp_line = None
                                return
                                
                            if branch_choice == "True (Left)":
                                label = "True"
                                color = TRUE_BRANCH_COLOR_LINE
                            elif branch_choice == "False (Right)":
                                label = "False"
                                color = FALSE_BRANCH_COLOR_LINE

                        connector = Connector(self.main_window.start_item, end_item, label, color)
                        self.scene().addItem(connector)
                        connector.update_position()
                        
                        # --- Manual Connection Registration (Keep this) ---
                        self.main_window.start_item.add_connector(connector)
                        end_item.add_connector(connector)
                        # ------------------------------------------------
                        
                    # Reset start item
                    pen = self.main_window.start_item.pen()
                    if self.main_window.start_item.step_data.get('if_branch') == 'true':
                        pen.setColor(TRUE_BRANCH_COLOR_LINE)
                        pen.setStyle(Qt.PenStyle.DashLine)
                    elif self.main_window.start_item.step_data.get('if_branch') == 'false':
                        pen.setColor(FALSE_BRANCH_COLOR_LINE)
                        pen.setStyle(Qt.PenStyle.DashLine)
                    else:
                        pen.setColor(QColor(SHAPE_BORDER_COLOR))
                        pen.setStyle(Qt.PenStyle.SolidLine)
                    pen.setWidth(2)
                    self.main_window.start_item.setPen(pen)
                    self.main_window.start_item = None
                    
                    # Remove temporary line
                    if self.temp_line:
                        self.scene().removeItem(self.temp_line)
                        self.temp_line = None
        else:
            # Normal mode - allow dragging blocks
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.main_window.connection_mode and self.temp_line and self.main_window.start_item:
            # Update temporary line to follow mouse
            scene_pos = self.mapToScene(event.pos())
            start_item = self.main_window.start_item
            
            # Use the connector's logic to determine the start point to be consistent
            temp_connector = Connector(start_item, start_item) # Dummy connector
            start_point_scene, _, is_true_branch, is_false_branch = temp_connector.get_connection_points()
            
            path = QPainterPath()
            path.moveTo(start_point_scene)
            
            # Simple line to follow the mouse for temporary visual
            path.lineTo(scene_pos)
            
            self.temp_line.setPath(path)
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        super().mouseReleaseEvent(event)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape and self.main_window.connection_mode:
            # Cancel connection mode
            if self.main_window.start_item:
                pen = self.main_window.start_item.pen()
                if self.main_window.start_item.step_data.get('if_branch') == 'true':
                    pen.setColor(TRUE_BRANCH_COLOR_LINE)
                    pen.setStyle(Qt.PenStyle.DashLine)
                elif self.main_window.start_item.step_data.get('if_branch') == 'false':
                    pen.setColor(FALSE_BRANCH_COLOR_LINE)
                    pen.setStyle(Qt.PenStyle.DashLine)
                else:
                    pen.setColor(QColor(SHAPE_BORDER_COLOR))
                    pen.setStyle(Qt.PenStyle.SolidLine)
                pen.setWidth(2)
                self.main_window.start_item.setPen(pen)
                self.main_window.start_item = None
            
            if self.temp_line:
                self.scene().removeItem(self.temp_line)
                self.temp_line = None
        else:
            super().keyPressEvent(event)
            
    # NEW METHOD: Handle Ctrl + Scroll for Zoom
    def wheelEvent(self, event):
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            # Determine zoom factor
            zoom_in_factor = 1.15
            zoom_out_factor = 1.0 / zoom_in_factor

            # Get the point in the view where the scroll wheel occurred
            point_before_scale = self.mapToScene(event.position().toPoint())

            # Apply transformation
            if event.angleDelta().y() > 0:
                # Zoom in
                self.scale(zoom_in_factor, zoom_in_factor)
            else:
                # Zoom out
                self.scale(zoom_out_factor, zoom_out_factor)

            # Get the point after scaling and translate the view to keep the initial point constant
            point_after_scale = self.mapToScene(event.position().toPoint())
            
            # Calculate the translation required
            translation_delta = point_before_scale - point_after_scale
            
            # Apply the translation
            self.translate(translation_delta.x(), translation_delta.y())

            # Accept the event to prevent default scrolling
            event.accept()
        else:
            # If Ctrl is not pressed, handle as a normal scroll (panning)
            super().wheelEvent(event)

# --- Main Application Window ---
class FlowchartApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Flowchart Creator")
        self.setGeometry(100, 100, 1400, 900)
        self.setStyleSheet(f"background-color: {BG_COLOR};")

        self.connection_mode = False
        self.start_item = None
        self.global_variables: Dict[str, Any] = {}
        self.flow_steps: List[str] = [] # Stores list of step_ids in execution order
        self.is_bot_running: bool = False
        self.gui_communicator = GuiCommunicator()
        self.worker: Optional[ExecutionWorker] = None

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        left_panel = QFrame()
        left_panel.setFixedWidth(250)
        left_panel.setStyleSheet(f"background-color: {LEFT_PANEL_COLOR}; border-radius: 5px;")
        left_panel_layout = QVBoxLayout(left_panel)
        left_panel_layout.setContentsMargins(10, 10, 10, 10)
        left_panel_layout.setSpacing(8)

        button_style = f"""
            QPushButton {{
                background-color: {BUTTON_COLOR}; color: {BUTTON_TEXT_COLOR}; border: 1px solid #B0B0B0;
                padding: 5px; font-size: 13px; border-radius: 5px;
            }}
            QPushButton:hover {{ background-color: #E6E6E6; border-color: #999999; }}
            QPushButton:pressed {{ background-color: #C0C0C0; }}
            QPushButton:checkable:checked {{ background-color: {HIGHLIGHT_COLOR}; color: #FFFFFF; border-color: #005A9E; }}
        """
        btn_add_step = QPushButton("Add Step"); btn_add_step.setStyleSheet(button_style); btn_add_step.clicked.connect(self.add_step)
        left_panel_layout.addWidget(btn_add_step)
        btn_add_decision = QPushButton("Add Decision Branch"); btn_add_decision.setStyleSheet(button_style); btn_add_decision.clicked.connect(self.add_decision_branch)
        left_panel_layout.addWidget(btn_add_decision)
        self.btn_connect_manual = QPushButton("Connect Manually"); self.btn_connect_manual.setStyleSheet(button_style); self.btn_connect_manual.setCheckable(True); self.btn_connect_manual.toggled.connect(self.toggle_connection_mode)
        left_panel_layout.addWidget(self.btn_connect_manual)
        
        # --- NEW BUTTON: "Open Your Bot" ---
        self.btn_open_bot = QPushButton("Open Your Bot")
        self.btn_open_bot.setStyleSheet(button_style)
        self.btn_open_bot.clicked.connect(self.open_bot_loader)
        left_panel_layout.addWidget(self.btn_open_bot)
        # --- END NEW BUTTON ---
        
        self.execute_all_button = QPushButton("Execute All Steps"); self.execute_all_button.setStyleSheet(button_style); self.execute_all_button.clicked.connect(self.execute_all_steps)
        left_panel_layout.addWidget(self.execute_all_button)
        self.execute_one_step_button = QPushButton("Execute 1 Step"); self.execute_one_step_button.setStyleSheet(button_style); self.execute_one_step_button.setEnabled(False); self.execute_one_step_button.clicked.connect(self.execute_one_step)
        left_panel_layout.addWidget(self.execute_one_step_button)
        self.add_loop_button = QPushButton("Add Loop"); self.add_loop_button.setStyleSheet(button_style)
        left_panel_layout.addWidget(self.add_loop_button)
        self.add_conditional_button = QPushButton("Add Conditional"); self.add_conditional_button.setStyleSheet(button_style)
        left_panel_layout.addWidget(self.add_conditional_button)
        self.save_steps_button = QPushButton("Save Bot"); self.save_steps_button.setStyleSheet(button_style); self.save_steps_button.clicked.connect(self.save_bot_steps_dialog)
        left_panel_layout.addWidget(self.save_steps_button)
        self.group_steps_button = QPushButton("Group Selected"); self.group_steps_button.setStyleSheet(button_style)
        left_panel_layout.addWidget(self.group_steps_button)
        self.clear_selected_button = QPushButton("Clear Selected"); self.clear_selected_button.setStyleSheet(button_style)
        left_panel_layout.addWidget(self.clear_selected_button)
        self.remove_all_steps_button = QPushButton("Remove All"); self.remove_all_steps_button.setStyleSheet(button_style); self.remove_all_steps_button.clicked.connect(self.clear_all)
        left_panel_layout.addWidget(self.remove_all_steps_button)
        self.progress_bar = QProgressBar(); self.progress_bar.hide()
        left_panel_layout.addWidget(self.progress_bar)
        self.always_on_top_button = QPushButton("Always On Top: Off"); self.always_on_top_button.setStyleSheet(button_style); self.always_on_top_button.setCheckable(True)
        left_panel_layout.addWidget(self.always_on_top_button)
        self.open_screenshot_tool_button = QPushButton("Screenshot Tool"); self.open_screenshot_tool_button.setStyleSheet(button_style)
        left_panel_layout.addWidget(self.open_screenshot_tool_button)
        self.exit_button = QPushButton("Exit GUI"); self.exit_button.setStyleSheet(button_style)
        left_panel_layout.addWidget(self.exit_button)
        left_panel_layout.addStretch()

        main_layout.addWidget(left_panel)

        right_splitter = QSplitter(Qt.Orientation.Vertical)

        self.scene = QGraphicsScene()
        self.view = FlowchartView(self.scene, self)
        self.view.setRenderHint(self.view.renderHints().Antialiasing)
        self.view.setFocusPolicy(Qt.FocusPolicy.StrongFocus)  # Enable keyboard events
        right_splitter.addWidget(self.view)
        
        bottom_widget = QWidget()
        bottom_layout = QHBoxLayout(bottom_widget)
        bottom_layout.setContentsMargins(0,0,0,0)
        bottom_splitter = QSplitter(Qt.Orientation.Horizontal)
        
        self.variables_group_box = QGroupBox("Global Variables")
        variables_layout = QVBoxLayout()
        self.variables_list = QListWidget()
        variables_layout.addWidget(self.variables_list)
        var_buttons_layout = QHBoxLayout()
        self.add_var_button = QPushButton("Add"); var_buttons_layout.addWidget(self.add_var_button)
        self.edit_var_button = QPushButton("Edit"); var_buttons_layout.addWidget(self.edit_var_button)
        self.delete_var_button = QPushButton("Delete"); var_buttons_layout.addWidget(self.delete_var_button)
        self.clear_vars_button = QPushButton("Reset"); var_buttons_layout.addWidget(self.clear_vars_button)
        variables_layout.addLayout(var_buttons_layout)
        self.variables_group_box.setLayout(variables_layout)
        bottom_splitter.addWidget(self.variables_group_box)

        log_group_box = QGroupBox("Execution Log")
        log_layout = QVBoxLayout()
        self.log_console = QTextEdit(); self.log_console.setReadOnly(True)
        log_layout.addWidget(self.log_console)
        self.clear_log_button = QPushButton("Clear Log"); self.clear_log_button.clicked.connect(self.log_console.clear); log_layout.addWidget(self.clear_log_button)
        log_group_box.setLayout(log_layout)
        bottom_splitter.addWidget(log_group_box)
        
        bottom_splitter.setSizes([350, 800])
        bottom_layout.addWidget(bottom_splitter)
        
        right_splitter.addWidget(bottom_widget)
        right_splitter.setSizes([self.height() - 250, 250])
        
        main_layout.addWidget(right_splitter)
        self.centralWidget().setLayout(main_layout)

        self.add_var_button.clicked.connect(self.add_variable)
        self.edit_var_button.clicked.connect(self.edit_variable)
        self.delete_var_button.clicked.connect(self.delete_variable)
        self.clear_vars_button.clicked.connect(self.reset_all_variable_values)
        self.gui_communicator.log_message_signal.connect(self._log_to_console)
        self._update_variables_list_display()

    def _log_to_console(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_console.append(f"[{timestamp}] {message}")

    def _get_item_by_id(self, item_id):
        for item in self.scene.items():
            if isinstance(item, FlowchartItem) and hasattr(item, 'step_id') and item.step_id == item_id:
                return item
        return None

# In flowchart.py

    def add_step(self):
        module_dialog = ModuleSelectionDialog(self)
        if not module_dialog.exec(): return
            
        method_info = module_dialog.selected_method_info
        if not method_info: return

        # Correctly unpack the full method info tuple: (module_name, class_name, method_name, params)
        module_name, class_name, method_name, params = method_info 
        param_dialog = ParameterInputDialog(f"{class_name}.{method_name}", params, list(self.global_variables.keys()), self)
        if not param_dialog.exec(): return
        
        config, assign_to = param_dialog.get_parameters_config()
        if config is None: return

        if assign_to and assign_to not in self.global_variables:
            self.global_variables[assign_to] = None
            self._update_variables_list_display()

        step_data = {'type': 'step', 'class': class_name, 'method': method_name, 'module_name': module_name, 'config': config, 'assign_to': assign_to, 'status': ''}
        
        new_item = StepItem(step_data=step_data)
        
        ordered_items = [self._get_item_by_id(sid) for sid in self.flow_steps if self._get_item_by_id(sid)]
        insertion_dialog = StepInsertionDialog(ordered_items, self)
        
        if insertion_dialog.exec():
            target_id, mode = insertion_dialog.get_insertion_point()
            
            self.scene.addItem(new_item)

            if target_id is None:
                self.flow_steps.append(new_item.step_id)
            else:
                try:
                    target_idx = self.flow_steps.index(target_id)
                    target_item = self._get_item_by_id(target_id)
                    
                    # --- START FIX: Determine Branching from Target Item (FIXED LOGIC) ---
                    branch_to_assign = None
                    if_id_to_assign = None
                    
                    # 1. Inherit from a regular step already in a branch (covers most cases)
                    if target_item and target_item.step_data.get('if_branch'):
                        branch_to_assign = target_item.step_data['if_branch']
                        if_id_to_assign = target_item.step_data['if_id']
                    
                    # 2. Derive branch from block markers if target doesn't have an 'if_branch' tag
                    elif target_item and target_item.step_data.get('type') in ['IF_START', 'ELSE', 'IF_END']:
                        target_type = target_item.step_data['type']
                        target_if_id = target_item.step_data.get('if_id')

                        if target_type == 'IF_START' and mode == 'after':
                            branch_to_assign = 'true' # Inserting after IF_START begins the TRUE branch
                            if_id_to_assign = target_if_id
                            
                        elif target_type == 'ELSE':
                            if target_if_id:
                                if mode == 'before':
                                    branch_to_assign = 'true'  # Inserting before ELSE means last step of the TRUE branch
                                    if_id_to_assign = target_if_id
                                elif mode == 'after':
                                    branch_to_assign = 'false' # Inserting after ELSE begins the FALSE branch
                                    if_id_to_assign = target_if_id

                        elif target_type == 'IF_END' and mode == 'before':
                            branch_to_assign = 'false' # Inserting before IF_END means last step of the FALSE branch
                            if_id_to_assign = target_if_id

                    # 3. Apply assignment and inherit parent properties
                    if branch_to_assign and if_id_to_assign:
                        new_item.step_data['if_branch'] = branch_to_assign
                        new_item.step_data['if_id'] = if_id_to_assign
                        
                        # Also inherit parent branch properties for nested IF blocks 
                        if target_item and target_item.step_data.get('parent_if_branch'):
                            new_item.step_data['parent_if_branch'] = target_item.step_data['parent_if_branch']
                            new_item.step_data['parent_if_id'] = target_item.step_data['parent_if_id']
                    # --- END FIX ---

                    if mode == 'after':
                        self.flow_steps.insert(target_idx + 1, new_item.step_id)
                    else: # 'before'
                        self.flow_steps.insert(target_idx, new_item.step_id)
                except ValueError:
                    self.flow_steps.append(new_item.step_id)

            self._redraw_flowchart()

# In flowchart.py: FlowchartApp.add_decision_branch

    def add_decision_branch(self):
        dialog = ConditionalConfigDialog(self.global_variables, self)
        if not dialog.exec(): return

        config = dialog.get_config()
        if not config: return
        
        if_id = f"if_{uuid.uuid4().hex[:6]}"
        if_data = {"type": "IF_START", "if_id": if_id, "condition_config": config}
        
        # --- FIX: Use a simple, zero-argument placeholder from an existing module ---
        # Assuming 'my_modules.py' is the simplest placeholder module available.
        PLACEHOLDER_MODULE = 'Bot_module.Gui_Automate'
        PLACEHOLDER_CLASS = 'Bot_utility'
        PLACEHOLDER_METHOD = 'wait_ms' 
        # Configuration to pass '1' (millisecond) to the 'milliseconds' parameter
        WAIT_MS_CONFIG = {
            'milliseconds': {
                'type': 'hardcoded',
                'value': 1  # Placeholder value of 1 millisecond
            }
        }
        
        true_data = {
            'type': 'step', 
            'class': PLACEHOLDER_CLASS, 
            'method': PLACEHOLDER_METHOD, 
            'module_name': PLACEHOLDER_MODULE, 
            'if_branch': 'true', 
            'if_id': if_id, 
            'config': WAIT_MS_CONFIG, 
            'assign_to': None
        }
        
        false_data = {
            'type': 'step', 
            'class': PLACEHOLDER_CLASS, 
            'method': PLACEHOLDER_METHOD, 
            'module_name': PLACEHOLDER_MODULE,
            'if_branch': 'false', 
            'if_id': if_id, 
            'config': WAIT_MS_CONFIG, 
            'assign_to': None
        }
        # --- END FIX ---
        
        end_if_data = {"type": "IF_END", "if_id": if_id}

        if_item = DecisionItem(text="IF...", step_data=if_data)
        true_item = StepItem(text="True Branch Start (wait_ms)", step_data=true_data)
        false_item = StepItem(text="False Branch Start (wait_ms)", step_data=false_data)
        end_if_item = StepItem(step_data=end_if_data)

        self.scene.addItem(if_item)
        self.scene.addItem(true_item)
        self.scene.addItem(false_item)
        self.scene.addItem(end_if_item)
        
        # --- Show Insertion Dialog for the whole block ---
        ordered_items = [self._get_item_by_id(sid) for sid in self.flow_steps if self._get_item_by_id(sid)]
        insertion_dialog = StepInsertionDialog(ordered_items, self)
        
        if insertion_dialog.exec():
            target_id, mode = insertion_dialog.get_insertion_point()
            new_block_ids = [if_item.step_id, true_item.step_id, false_item.step_id, end_if_item.step_id]

            # Inherit parent branch properties if inserting into an existing branch
            if target_id is not None:
                target_item = self._get_item_by_id(target_id)
                if target_item and target_item.step_data.get('if_branch'):
                    parent_branch = target_item.step_data.get('if_branch')
                    parent_if_id = target_item.step_data.get('if_id')
                    
                    # Set parent branch info for the entire nested IF block
                    if_item.step_data['parent_if_branch'] = parent_branch
                    if_item.step_data['parent_if_id'] = parent_if_id
                    if_item.step_data['if_branch'] = parent_branch
                    if_item.step_data['outer_if_id'] = parent_if_id
                    
                    true_item.step_data['parent_if_branch'] = parent_branch
                    true_item.step_data['parent_if_id'] = parent_if_id
                    
                    false_item.step_data['parent_if_branch'] = parent_branch
                    false_item.step_data['parent_if_id'] = parent_if_id
                    
                    end_if_item.step_data['if_branch'] = parent_branch
                    end_if_item.step_data['outer_if_id'] = parent_if_id

            if target_id is None:
                self.flow_steps.extend(new_block_ids)
            else:
                try:
                    idx = self.flow_steps.index(target_id)
                    if mode == 'after':
                        self.flow_steps[idx + 1:idx + 1] = new_block_ids
                    else: # 'before'
                         self.flow_steps[idx:idx] = new_block_ids
                except ValueError:
                    self.flow_steps.extend(new_block_ids) # Fallback

            self._redraw_flowchart()
        else:
             # If insertion is cancelled, remove the items we just added
             self.scene.removeItem(if_item)
             self.scene.removeItem(true_item)
             self.scene.removeItem(false_item)
             self.scene.removeItem(end_if_item)

    def _redraw_flowchart(self):
        # ... (Unchanged from previous turn - responsible for layout and drawing connections)
        all_items_on_scene = [item for item in self.scene.items() if isinstance(item, FlowchartItem)]
        item_map = {item.step_id: item for item in all_items_on_scene}
        
        for item in self.scene.items():
            if isinstance(item, Connector):
                self.scene.removeItem(item)
        
        # Clear all connector lists before redrawing connections
        for item in all_items_on_scene:
            item.connectors.clear()


        y_pos = 50
        last_item = None
        
        def process_if_block(if_item, start_y, x_center):
            """Recursively process an IF block and return the END_IF item and next y position"""
            if_id = if_item.step_data['if_id']
            if_idx = self.flow_steps.index(if_item.step_id)
            
            # Find all items belonging to this specific IF block
            true_branch_ids = []
            false_branch_ids = []
            end_if_idx = -1
            nesting_level = 0
            
            for j in range(if_idx + 1, len(self.flow_steps)):
                s_id = self.flow_steps[j]
                s_item = item_map.get(s_id)
                if not s_item: continue

                s_type = s_item.step_data.get('type')
                s_if_id = s_item.step_data.get('if_id')
                
                # Logic to handle nested IF blocks correctly
                if s_type == 'IF_START':
                    # Only increment nesting if the inner IF is inside the current branch's scope
                    if s_item.step_data.get('parent_if_id') == if_id or (s_if_id == if_id and nesting_level == 0):
                        nesting_level += 1
                        
                elif s_type == 'IF_END':
                    # Only decrement nesting if the END_IF matches the last nested IF
                    if nesting_level > 0 and s_if_id != if_id and s_item.step_data.get('parent_if_id') == if_id:
                        nesting_level -= 1
                    elif nesting_level == 0 and s_if_id == if_id:
                        end_if_idx = j
                        break
                
                # Check for item inclusion based on nesting and if_id
                if nesting_level == 0 and s_if_id == if_id and s_type != 'IF_END':
                    if s_item.step_data.get('if_branch') == 'true':
                        true_branch_ids.append(s_id)
                    elif s_item.step_data.get('if_branch') == 'false':
                        false_branch_ids.append(s_id)
                elif nesting_level > 0 and s_type != 'IF_START':
                    # Collect steps belonging to nested blocks for correct sequential positioning within the branch
                    parent_if_id = s_item.step_data.get('parent_if_id')
                    if parent_if_id == if_id:
                        parent_branch = s_item.step_data.get('parent_if_branch')
                        if parent_branch == 'true':
                            true_branch_ids.append(s_id)
                        elif parent_branch == 'false':
                            false_branch_ids.append(s_id)


            # After the loop, true_branch_ids and false_branch_ids should contain all steps/blocks in their order
            if end_if_idx == -1:
                print(f"Error: Could not find matching END_IF for {if_id}")
                return None, start_y
            
            end_if_item = item_map[self.flow_steps[end_if_idx]]
            
            # Position IF_START
            if_item.setPos(x_center, start_y)
            
            branch_y_start = start_y + (DECISION_HEIGHT/2) + VERTICAL_SPACING
            
            # Process True branch
            current_y_true = branch_y_start
            last_true_item = if_item
            true_x = x_center - HORIZONTAL_SPACING
            
            i = 0
            while i < len(true_branch_ids):
                true_id = true_branch_ids[i]
                true_item = item_map[true_id]
                
                if true_item.step_data.get('type') == 'IF_START':
                    # Nested IF block in TRUE branch
                    nested_end_if, next_y = process_if_block(true_item, current_y_true, true_x)
                    self._add_connector(last_true_item, true_item, "True" if last_true_item == if_item else "", TRUE_BRANCH_COLOR_LINE)
                    last_true_item = nested_end_if if nested_end_if else true_item
                    current_y_true = next_y
                    
                    # Advance index past the end of the nested block in the current branch list
                    nested_if_id = true_item.step_data['if_id']
                    nested_if_idx_in_flow_steps = self.flow_steps.index(true_id)
                    
                    temp_j = nested_if_idx_in_flow_steps + 1
                    while temp_j < len(self.flow_steps):
                        check_id = self.flow_steps[temp_j]
                        check_item = item_map.get(check_id)
                        if check_item and check_item.step_data.get('type') == 'IF_END' and check_item.step_data.get('if_id') == nested_if_id:
                            # Found the END_IF, now figure out how many items this skips in true_branch_ids
                            skip_count = 0
                            for k in range(i, len(true_branch_ids)):
                                if true_branch_ids[k] == check_id:
                                    skip_count = k - i
                                    break
                            i += skip_count # Move index past the nested END_IF
                            break
                        temp_j += 1
                
                elif true_item.step_data.get('type') == 'IF_END':
                    i += 1 # Skip END_IF blocks as they're handled by their IF_START
                else:
                    true_item.setPos(true_x, current_y_true)
                    self._add_connector(last_true_item, true_item, "True" if last_true_item == if_item else "", TRUE_BRANCH_COLOR_LINE)
                    last_true_item = true_item
                    current_y_true += STEP_HEIGHT + VERTICAL_SPACING
                    i += 1
            
            # Process False branch
            current_y_false = branch_y_start
            last_false_item = if_item
            false_x = x_center + HORIZONTAL_SPACING
            
            i = 0
            while i < len(false_branch_ids):
                false_id = false_branch_ids[i]
                false_item = item_map[false_id]
                
                if false_item.step_data.get('type') == 'IF_START':
                    # Nested IF block in FALSE branch
                    nested_end_if, next_y = process_if_block(false_item, current_y_false, false_x)
                    self._add_connector(last_false_item, false_item, "False" if last_false_item == if_item else "", FALSE_BRANCH_COLOR_LINE)
                    last_false_item = nested_end_if if nested_end_if else false_item
                    current_y_false = next_y
                    
                    # Advance index past the end of the nested block in the current branch list
                    nested_if_id = false_item.step_data['if_id']
                    nested_if_idx_in_flow_steps = self.flow_steps.index(false_id)
                    
                    temp_j = nested_if_idx_in_flow_steps + 1
                    while temp_j < len(self.flow_steps):
                        check_id = self.flow_steps[temp_j]
                        check_item = item_map.get(check_id)
                        if check_item and check_item.step_data.get('type') == 'IF_END' and check_item.step_data.get('if_id') == nested_if_id:
                            # Found the END_IF, now figure out how many items this skips in false_branch_ids
                            skip_count = 0
                            for k in range(i, len(false_branch_ids)):
                                if false_branch_ids[k] == check_id:
                                    skip_count = k - i
                                    break
                            i += skip_count # Move index past the nested END_IF
                            break
                        temp_j += 1
                
                elif false_item.step_data.get('type') == 'IF_END':
                    i += 1 # Skip END_IF blocks as they're handled by their IF_START
                else:
                    false_item.setPos(false_x, current_y_false)
                    self._add_connector(last_false_item, false_item, "False" if last_false_item == if_item else "", FALSE_BRANCH_COLOR_LINE)
                    last_false_item = false_item
                    current_y_false += STEP_HEIGHT + VERTICAL_SPACING
                    i += 1
            
            # Determine max branch height for END_IF positioning
            max_branch_y = max(current_y_true, current_y_false)
            
            # Position END_IF
            end_if_y_pos = max_branch_y
            end_if_item.setPos(x_center, end_if_y_pos)
            
            # Connect last items of both branches to END_IF
            self._add_connector(last_true_item, end_if_item, "", TRUE_BRANCH_COLOR_LINE)
            self._add_connector(last_false_item, end_if_item, "", FALSE_BRANCH_COLOR_LINE)
            
            return end_if_item, end_if_y_pos + STEP_HEIGHT + VERTICAL_SPACING
        
        # Main processing loop
        i = 0
        while i < len(self.flow_steps):
            step_id = self.flow_steps[i]
            item = item_map.get(step_id)
            if not item: 
                i += 1
                continue

            # Skip items that are part of IF blocks (they'll be processed recursively)
            if item.step_data.get('if_branch') or item.step_data.get('type') == 'IF_END':
                i += 1
                continue

            if item.step_data.get('type') == 'IF_START' and not item.step_data.get('outer_if_id'):
                # Top-level IF block
                x_center = self.view.mapToScene(self.view.viewport().rect().center()).x()
                if last_item:
                    self._add_connector(last_item, item)
                
                end_if_item, y_pos = process_if_block(item, y_pos, x_center)
                last_item = end_if_item
                
                # Skip to after the END_IF
                if end_if_item:
                    end_if_idx = self.flow_steps.index(end_if_item.step_id)
                    i = end_if_idx + 1
                else:
                    i += 1
            else:
                # Regular step (not in any branch)
                x_pos = self.view.mapToScene(self.view.viewport().rect().center()).x()
                item.setPos(x_pos, y_pos)
                if last_item:
                    self._add_connector(last_item, item)
                last_item = item
                y_pos += item.boundingRect().height() + VERTICAL_SPACING
                i += 1
        
        # Update step numbers and details for all items
        for i, step_id in enumerate(self.flow_steps):
            item = item_map.get(step_id)
            if item:
                item.set_step_details(i + 1, item.step_data.get('status', ''))
                
        # Adjust scene rect to fit all items
        self.view.setSceneRect(self.scene.itemsBoundingRect().adjusted(-50, -50, 50, 50))


    def _add_connector(self, start_item, end_item, label="", color=QColor(LINE_COLOR)):
        connector = Connector(start_item, end_item, label, color)
        self.scene.addItem(connector)
        connector.update_position()
        
        # Register connector with both items to enable movement tracking
        start_item.add_connector(connector)
        end_item.add_connector(connector)
        
        return connector

    def clear_all(self):
        reply = QMessageBox.question(self, "Confirm Clear",
                                     "Are you sure you want to remove everything?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self._reset_flowchart_state()

    def _reset_flowchart_state(self):
        self.scene.clear()
        self.flow_steps.clear()
        self.global_variables.clear()
        self._update_variables_list_display()
        self.start_item = None
        self.btn_connect_manual.setChecked(False)
        self.log_console.clear()
        self.log_console.append("Flowchart state reset.")
        self.execute_one_step_button.setEnabled(False)


    def toggle_connection_mode(self, checked):
        self.connection_mode = checked
        if checked:
            self.start_item = None
            self.setCursor(Qt.CursorShape.CrossCursor)
        else:
            self.setCursor(Qt.CursorShape.ArrowCursor)
            if self.start_item:
                pen = self.start_item.pen()
                if self.start_item.step_data.get('if_branch') == 'true':
                    pen.setColor(TRUE_BRANCH_COLOR_LINE)
                    pen.setStyle(Qt.PenStyle.DashLine)
                elif self.start_item.step_data.get('if_branch') == 'false':
                    pen.setColor(FALSE_BRANCH_COLOR_LINE)
                    pen.setStyle(Qt.PenStyle.DashLine)
                else:
                    pen.setColor(QColor(SHAPE_BORDER_COLOR))
                    pen.setStyle(Qt.PenStyle.SolidLine)
                pen.setWidth(2)
                self.start_item.setPen(pen)
                self.start_item = None

    def _update_variables_list_display(self):
        self.variables_list.clear()
        if not self.global_variables:
            self.variables_list.addItem("No global variables defined.")
            return
        for name, value in self.global_variables.items():
            value_str = repr(value)
            if len(value_str) > 60: value_str = value_str[:57] + "..."
            self.variables_list.addItem(f"{name} = {value_str}")
    
    def add_variable(self):
        dialog = GlobalVariableDialog(parent=self)
        if dialog.exec():
            name, value = dialog.get_variable_data()
            if name:
                if name in self.global_variables:
                    QMessageBox.warning(self, "Duplicate Variable", f"A variable named '{name}' already exists.")
                    return
                # Convert string representation of literal Python values
                try:
                    self.global_variables[name] = ast.literal_eval(value)
                except (ValueError, SyntaxError):
                    self.global_variables[name] = value

                self._update_variables_list_display()
    
    def edit_variable(self):
        # ... (Unchanged from previous turn)
        selected_item = self.variables_list.currentItem()
        if not selected_item or "No global variables" in selected_item.text():
            QMessageBox.information(self, "No Selection", "Please select a variable to edit.")
            return
        
        var_name = selected_item.text().split(' = ')[0]
        dialog = GlobalVariableDialog(var_name, self.global_variables.get(var_name, ""), self)
        
        if dialog.exec():
            new_name, new_value = dialog.get_variable_data()
            if new_name:
                if new_name != var_name and new_name in self.global_variables:
                    QMessageBox.warning(self, "Duplicate Variable", f"A variable named '{new_name}' already exists.")
                    return
                if new_name != var_name and var_name in self.global_variables:
                    del self.global_variables[var_name]
                
                try:
                    self.global_variables[new_name] = ast.literal_eval(new_value)
                except (ValueError, SyntaxError):
                    self.global_variables[new_name] = new_value

                self._update_variables_list_display()

    def delete_variable(self):
        # ... (Unchanged from previous turn)
        selected_item = self.variables_list.currentItem()
        if not selected_item or "No global variables" in selected_item.text():
            QMessageBox.information(self, "No Selection", "Please select a variable to delete.")
            return

        var_name = selected_item.text().split(' = ')[0]
        reply = QMessageBox.question(self, "Confirm Delete",
                                     f"Are you sure you want to delete the variable '{var_name}'?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            if var_name in self.global_variables:
                del self.global_variables[var_name]
            self._update_variables_list_display()

    def reset_all_variable_values(self):
        # ... (Unchanged from previous turn)
        if not self.global_variables:
            QMessageBox.information(self, "Info", "There are no global variables to reset.")
            return
        
        reply = QMessageBox.question(self, "Confirm Reset",
                                     "Are you sure you want to reset all variable values to None?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            for var_name in self.global_variables:
                self.global_variables[var_name] = None
            self._update_variables_list_display()

    def _get_bot_names(self) -> list:
        """Returns a sorted list of existing bot names (without extension)."""
        os.makedirs(BOT_STEPS_DIR, exist_ok=True)
        return sorted([os.path.splitext(f)[0] for f in os.listdir(BOT_STEPS_DIR) if f.endswith(".csv")])

    def _get_flat_steps_data(self) -> list:
        """Converts the list of step_ids into a linear list of step_data dictionaries."""
        flat_steps_data = []
        item_map = {item.step_id: item for item in self.scene.items() if isinstance(item, FlowchartItem)}
        
        for step_id in self.flow_steps:
            item = item_map.get(step_id)
            if item and item.step_data:
                # Need to use a deep copy or copy to prevent modifying the original dict in FlowchartItem
                step_data_copy = item.step_data.copy()
                flat_steps_data.append(step_data_copy)
                
        return flat_steps_data

    def save_bot_steps_dialog(self) -> None:
        """Opens a dialog to name the bot, then saves the steps and variables to a CSV file."""
        # ... (Unchanged from previous turn)
        
        flat_steps = self._get_flat_steps_data()

        if not flat_steps and not self.global_variables:
            QMessageBox.information(self, "Nothing to Save", "The flowchart and global variables are empty.")
            return

        existing_bots = self._get_bot_names()
        dialog = SaveBotDialog(existing_bots, self)
        
        if dialog.exec() == QDialog.DialogCode.Accepted:
            bot_name = dialog.get_bot_name()
            if not bot_name: return # Dialog already showed an error message

            file_path = os.path.join(BOT_STEPS_DIR, f"{bot_name}.csv")

            # Check for overwrite
            if bot_name in existing_bots:
                 reply = QMessageBox.question(self, "Confirm Overwrite",
                                              f"A bot named '{bot_name}' already exists. Overwrite it?",
                                              QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                 if reply == QMessageBox.StandardButton.No:
                    return

            try:
                # Only serialize primitive types (str, int, float, bool, list, dict, None)
                variables_to_save = {}
                for var_name, var_value in self.global_variables.items():
                    # Attempt to dump/load to check serializability. If it fails, save as None.
                    try:
                        json.dumps(var_value)
                        variables_to_save[var_name] = var_value
                    except (TypeError, OverflowError):
                        variables_to_save[var_name] = None 

                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    
                    # 1. Write variables
                    writer.writerow(["__GLOBAL_VARIABLES__"])
                    for var_name, var_value in variables_to_save.items():
                        writer.writerow([var_name, json.dumps(var_value)])
                    
                    # 2. Write steps
                    writer.writerow(["__BOT_STEPS__"])
                    writer.writerow(["StepType", "DataJSON"])
                    for step_data_dict in flat_steps:
                        # Ensure keys like 'original_listbox_row_index' are removed before final save
                        step_data_to_save = step_data_dict.copy()
                        # NOTE: For simplicity in flowchart view, we're not dealing with deep cleaning internal keys
                        
                        writer.writerow([step_data_to_save["type"], json.dumps(step_data_to_save)])
                
                QMessageBox.information(self, "Save Successful", f"Bot saved to:\n{file_path}")
            
            except Exception as e:
                QMessageBox.critical(self, "Save Error", f"Failed to save bot steps:\n{e}")

    def load_bot_steps(self, bot_name: str) -> None:
        """Loads a saved bot's steps and variables from a CSV file and rebuilds the flowchart."""
        self._reset_flowchart_state() # Clear current state
        file_path = os.path.join(BOT_STEPS_DIR, f"{bot_name}.csv")

        if not os.path.exists(file_path):
            QMessageBox.critical(self, "Load Error", f"Bot file not found: {file_path}")
            return

        try:
            current_section = None
            step_data_list = []

            with open(file_path, 'r', newline='', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)

                for row in reader:
                    if not row: continue

                    if row[0] == "__GLOBAL_VARIABLES__":
                        current_section = "VARIABLES"
                        continue
                    elif row[0] == "__BOT_STEPS__":
                        current_section = "STEPS"
                        next(reader, None) # Skip StepType,DataJSON header
                        continue

                    if current_section == "VARIABLES" and len(row) >= 2:
                        var_name = row[0]
                        var_value_str = row[1]
                        try:
                            # Value is stored as JSON string in CSV
                            self.global_variables[var_name] = json.loads(var_value_str)
                        except (json.JSONDecodeError, TypeError):
                            self.global_variables[var_name] = var_value_str # Fallback to string if parsing fails

                    elif current_section == "STEPS" and len(row) >= 2:
                        data_json_str = row[1]
                        try:
                            step_data = json.loads(data_json_str)
                            step_data_list.append(step_data)
                        except json.JSONDecodeError as e:
                            self.log_console.append(f"Error decoding step data: {e} in row: {row}")
                            QMessageBox.warning(self, "Data Error", f"Skipping corrupted step in CSV: {e}")

            # Recreate Flowchart Items
            self._update_variables_list_display()
            self.log_console.append(f"Successfully loaded bot '{bot_name}'. Rebuilding flowchart...")
            
            for i, data in enumerate(step_data_list):
                step_type = data.get('type')
                
                # Add a temporary index for the worker to track progress, as required by main_app's logic
                data["original_listbox_row_index"] = i 

                if step_type == 'IF_START':
                    new_item = DecisionItem(step_data=data)
                elif step_type == 'IF_END' or step_type == 'step':
                    new_item = StepItem(step_data=data)
                else:
                    self.log_console.append(f"Warning: Unknown step type '{step_type}'. Skipping.")
                    continue

                self.scene.addItem(new_item)
                self.flow_steps.append(new_item.step_id)

            self._redraw_flowchart()
            self.log_console.append("Flowchart rebuilt successfully.")
            
        except Exception as e:
            self._reset_flowchart_state()
            QMessageBox.critical(self, "Load Error", f"Failed to load bot steps from '{bot_name}':\n{e}")

    def open_bot_loader(self):
        """Opens the new dialog to load saved bots and connects the selection signal."""
        dialog = BotLoaderDialog(BOT_STEPS_DIR, parent=self)
        # Connect the dialog's signal to the loading method
        dialog.bot_selected.connect(self.load_bot_steps)
        dialog.exec()
        
    def _validate_block_structure_on_execution(self) -> bool:
        """Checks for unclosed or mismatched blocks before execution."""
        open_blocks = []
        
        # NOTE: Using the simple list of step IDs and retrieving data to replicate the check.
        flat_steps = self._get_flat_steps_data() 

        for step_data in flat_steps:
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

    def execute_all_steps(self) -> None:
        """Starts the ExecutionWorker to run all steps."""
        if self.is_bot_running:
            QMessageBox.warning(self, "Execution in Progress", "A bot is already running. Please wait for it to complete.")
            return

        steps_to_execute = self._get_flat_steps_data()
        if not steps_to_execute:
            QMessageBox.information(self, "No Steps", "No steps have been added.")
            return
        
        # Add temporary indices for the worker to track progress
        for i, step_data in enumerate(steps_to_execute):
             step_data["original_listbox_row_index"] = i

        if not self._validate_block_structure_on_execution():
            return
        
        self.is_bot_running = True
        self.set_ui_enabled_state(False)
        self.progress_bar.setValue(0)
        self.progress_bar.show()
        self._clear_status_for_all_items()
        
        self.worker = ExecutionWorker(
            steps_to_execute, 
            MODULE_DIR, 
            self.gui_communicator, 
            self.global_variables,
            single_step_mode=False
        )
        self._connect_worker_signals()
        self.worker.start()

    def execute_one_step(self) -> None:
        """Starts the ExecutionWorker to run a single step."""
        if self.is_bot_running:
            QMessageBox.warning(self, "Execution in Progress", "A bot is already running. Please wait for it to complete.")
            return

        # Find the index of the last item in the flow_steps list.
        # This acts as the currently selected step in a simple Flowchart.
        if not self.flow_steps:
             QMessageBox.information(self, "No Steps", "No steps have been added.")
             return

        selected_id = self.flow_steps[-1]
        
        # Get the step's index in the flattened list (needed for worker start index)
        steps_to_execute = self._get_flat_steps_data()
        current_row = -1
        for i, step_data in enumerate(steps_to_execute):
             step_data["original_listbox_row_index"] = i
             # Find the index of the step with the selected ID
             if step_data.get('step_id') == selected_id:
                 current_row = i
                 break

        if current_row == -1:
            QMessageBox.critical(self, "Error", "Could not find the last drawn step.")
            return
        
        if not self._validate_block_structure_on_execution():
            return

        self.is_bot_running = True
        self.set_ui_enabled_state(False)
        self.progress_bar.setValue(0)
        self.progress_bar.show()
        self._clear_status_for_all_items()
        
        self.worker = ExecutionWorker(
            steps_to_execute, 
            MODULE_DIR, 
            self.gui_communicator, 
            self.global_variables, 
            single_step_mode=True, 
            selected_start_index=current_row
        )
        self._connect_worker_signals()
        self.worker.start()

    def _connect_worker_signals(self) -> None:
        if self.worker:
            try: self.worker.execution_started.disconnect()
            except TypeError: pass
            try: self.worker.execution_progress.disconnect()
            except TypeError: pass
            try: self.worker.execution_item_started.disconnect()
            except TypeError: pass
            try: self.worker.execution_item_finished.disconnect()
            except TypeError: pass
            try: self.worker.execution_error.disconnect()
            except TypeError: pass
            try: self.worker.execution_finished_all.disconnect()
            except TypeError: pass
            
            self.worker.execution_progress.connect(self.progress_bar.setValue)
            self.worker.execution_item_started.connect(self.update_execution_tree_item_status_started)
            self.worker.execution_item_finished.connect(self.update_execution_tree_item_status_finished)
            self.worker.execution_error.connect(self.update_execution_tree_item_status_error)
            self.worker.execution_finished_all.connect(self.on_execution_finished)

    def update_execution_tree_item_status_started(self, step_data_dict: Dict[str, Any], original_listbox_row_index: int) -> None:
        item_to_highlight = self._get_item_by_id(step_data_dict.get('step_id'))
        if item_to_highlight:
            item_to_highlight.set_step_details(original_listbox_row_index + 1, "RUNNING...")
            self.view.ensureVisible(item_to_highlight, 50, 50)
            self.view.viewport().update()

    def update_execution_tree_item_status_finished(self, step_data_dict: Dict[str, Any], message: str, original_listbox_row_index: int) -> None:
        item_to_highlight = self._get_item_by_id(step_data_dict.get('step_id'))
        if item_to_highlight:
            item_to_highlight.set_step_details(original_listbox_row_index + 1, "COMPLETED")
            self.view.viewport().update()
        self._update_variables_list_display()

    def update_execution_tree_item_status_error(self, step_data_dict: Dict[str, Any], error_message: str, original_listbox_row_index: int) -> None:
        item_to_highlight = self._get_item_by_id(step_data_dict.get('step_id'))
        if item_to_highlight:
            item_to_highlight.set_step_details(original_listbox_row_index + 1, "ERROR!")
            self.view.viewport().update()
        self._update_variables_list_display()

    def on_execution_finished(self, context: ExecutionContext, stopped_by_error: bool, next_step_index_to_select: int) -> None:
        self.is_bot_running = False
        self.progress_bar.setValue(100)
        self.set_ui_enabled_state(True)
        
        if stopped_by_error:
            QMessageBox.critical(self, "Execution Halted", "Execution stopped due to an error.")
        else:
            self.log_console.append("Execution finished successfully.")
        
        if next_step_index_to_select != -1 and 0 <= next_step_index_to_select < len(self.flow_steps):
            target_step_id = self._get_flat_steps_data()[next_step_index_to_select].get('step_id')
            target_item = self._get_item_by_id(target_step_id)
            if target_item:
                # We simply focus the next step as if it were a visual indicator
                self.log_console.append(f"Next step to execute: Step {next_step_index_to_select + 1}")
                self.view.ensureVisible(target_item, 50, 50)

    def _clear_status_for_all_items(self):
        """Resets the status of all flowchart items."""
        for item in self.scene.items():
            if isinstance(item, FlowchartItem):
                # The step number is correctly reset in _redraw_flowchart when re-run.
                item.set_step_details(self.flow_steps.index(item.step_id) + 1, "")

    def set_ui_enabled_state(self, enabled: bool) -> None:
        widgets_to_toggle = [
            self.execute_all_button, self.add_loop_button, self.add_conditional_button, 
            self.save_steps_button, self.clear_selected_button, self.remove_all_steps_button,
            self.add_var_button, self.edit_var_button, self.delete_var_button, 
            self.clear_vars_button, self.open_screenshot_tool_button, 
            self.group_steps_button, self.btn_open_bot, self.btn_connect_manual,
        ]
        for widget in widgets_to_toggle:
            widget.setEnabled(enabled)

        # Enable 'Execute 1 Step' only when UI is generally enabled
        self.execute_one_step_button.setEnabled(enabled and len(self.flow_steps) > 0)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FlowchartApp()
    window.show()
    sys.exit(app.exec())
