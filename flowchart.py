import sys
import os
import inspect
import importlib
import uuid
import json
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QGraphicsView, QGraphicsScene, QWidget,
    QInputDialog, QGraphicsPolygonItem, QGraphicsTextItem,
    QMessageBox, QPushButton, QVBoxLayout, QHBoxLayout, QFrame, QDialog,
    QTreeWidget, QTreeWidgetItem, QDialogButtonBox, QSplitter, QGroupBox,
    QListWidget, QTextEdit, QProgressBar, QCheckBox, QLineEdit, QFormLayout,
    QFileDialog, QComboBox, QRadioButton, QLabel, QListWidgetItem,
    QGridLayout, QTreeWidgetItemIterator, QGraphicsPathItem
)
from PyQt6.QtGui import QPolygonF, QBrush, QPen, QFont, QColor, QPainterPath
from PyQt6.QtCore import QPointF, Qt, QRectF, QVariant, pyqtSignal

# --- Configuration ---
STEP_WIDTH = 250
STEP_HEIGHT = 120
DECISION_WIDTH = 280
DECISION_HEIGHT = 100
VERTICAL_SPACING = 80
HORIZONTAL_SPACING = 300
GRID_SIZE = 20  # Define the size of the grid squares

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


# --- Dummy Classes for Compatibility ---
class GuiCommunicator:
    update_module_info_signal = pyqtSignal(str)

class ExecutionContext:
    pass

# --- Global Variable Dialog ---
class GlobalVariableDialog(QDialog):
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

# --- Conditional Config Dialog ---
class ConditionalConfigDialog(QDialog):
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

# --- Parameter Input Dialog ---
class ParameterInputDialog(QDialog):
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

# --- Step Insertion Dialog ---
class StepInsertionDialog(QDialog):
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

# --- Module Selection Dialog ---
class ModuleSelectionDialog(QDialog):
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


# --- Custom Graphics Items ---
class Connector(QGraphicsPathItem):
    def __init__(self, start_item, end_item, label="", color=QColor(LINE_COLOR), parent=None):
        super().__init__(parent)
        self.start_item = start_item
        self.end_item = end_item
        self.setPen(QPen(color, 2))
        self.setZValue(-1)

        self.label = QGraphicsTextItem(label, self)
        self.label.setDefaultTextColor(color)
        self.label.setFont(QFont("Arial", 10, QFont.Weight.Bold))

    def update_position(self):
        start_rect = self.start_item.boundingRect()
        end_rect = self.end_item.boundingRect()
        
        # Determine base connection points in Scene coordinates
        is_true_branch = self.label.toPlainText() == "True"
        is_false_branch = self.label.toPlainText() == "False"
        
        if isinstance(self.start_item, DecisionItem):
            if is_true_branch:
                # Left side of diamond
                start_point = self.start_item.mapToScene(QPointF(-DECISION_WIDTH/2, 0))
            elif is_false_branch:
                # Right side of diamond
                start_point = self.start_item.mapToScene(QPointF(DECISION_WIDTH/2, 0))
            else:
                # Bottom point of diamond
                start_point = self.start_item.mapToScene(QPointF(0, DECISION_HEIGHT/2))
        else:
            # Bottom center of rectangle
            start_point = self.start_item.mapToScene(QPointF(0, start_rect.height() / 2))
        
        if isinstance(self.end_item, DecisionItem) and not (is_true_branch or is_false_branch):
            # Top point of diamond (default entry)
            end_point = self.end_item.mapToScene(QPointF(0, -DECISION_HEIGHT/2))
        else:
            # Top center of rectangle/step
            end_point = self.end_item.mapToScene(QPointF(0, -end_rect.height() / 2))

        # --- Corrected Path Drawing Logic for Moveable Shapes ---
        
        path = QPainterPath()
        path.moveTo(start_point)
        
        dy = end_point.y() - start_point.y()
        # dx = end_point.x() - start_point.x() # Not strictly needed for the path logic below

        # Determine if this is a connection from the side of a DecisionItem
        is_side_branch_from_decision = isinstance(self.start_item, DecisionItem) and (is_true_branch or is_false_branch)
        
        # The key to keeping lines attached is using an elbow (VHV or HVH) path
        # that doesn't rely on fixed grid coordinates.
        if is_side_branch_from_decision:
            # HVH path: Horizontal-Vertical-Horizontal
            # 1. Move horizontally out from the DecisionItem to clear the shape.
            #    DECISION_WIDTH/2 is the max horizontal size. We add a 10px buffer.
            h_clearance = DECISION_WIDTH / 2 + 10
            x_offset = h_clearance * (1 if is_false_branch else -1)
            
            waypoint_x = start_point.x() + x_offset
            
            path.lineTo(waypoint_x, start_point.y()) # Move horizontally out
            
            # 2. Move vertically to align with the target's Y position
            path.lineTo(waypoint_x, end_point.y())
            
            # 3. Move horizontally into the target item
            path.lineTo(end_point)
        else:
            # VHV path: Vertical-Horizontal-Vertical (Standard Top/Bottom Connection)
            # 1. Move vertically halfway to the target for a clean elbow
            mid_y = start_point.y() + dy / 2
            
            path.lineTo(start_point.x(), mid_y)
            
            # 2. Move horizontally to align with the target's X position
            path.lineTo(end_point.x(), mid_y)
            
            # 3. Move vertically into the target
            path.lineTo(end_point)

        self.setPath(path)
        
        # Adjust label position
        if is_true_branch or is_false_branch:
            # Position label near the start point for branch labels
            label_x = start_point.x() + (-40 if is_true_branch else 10)
            label_y = start_point.y() + 5
        else:
            # Position label in the middle for other connections
            label_x = (start_point.x() + end_point.x()) / 2 - 20
            label_y = (start_point.y() + end_point.y()) / 2 - 10
        
        self.label.setPos(label_x, label_y)


class FlowchartItem(QGraphicsPolygonItem):
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
    def __init__(self, text="Step", step_data=None, parent=None):
        rect = QPolygonF([
            QPointF(-STEP_WIDTH / 2, -STEP_HEIGHT / 2), QPointF(STEP_WIDTH / 2, -STEP_HEIGHT / 2),
            QPointF(STEP_WIDTH / 2, STEP_HEIGHT / 2), QPointF(-STEP_WIDTH / 2, STEP_HEIGHT / 2),
        ])
        super().__init__(rect, text, step_data, parent)

class DecisionItem(FlowchartItem):
    def __init__(self, text="If...", step_data=None, parent=None):
        diamond = QPolygonF([
            QPointF(0, -DECISION_HEIGHT / 2), QPointF(DECISION_WIDTH / 2, 0),
            QPointF(0, DECISION_HEIGHT / 2), QPointF(-DECISION_WIDTH / 2, 0),
        ])
        super().__init__(diamond, text, step_data, parent)

# --- Custom QGraphicsView ---
class FlowchartView(QGraphicsView):
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
                        # Create the connection
                        connector = Connector(self.main_window.start_item, end_item)
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
            start_rect = self.main_window.start_item.boundingRect()
            
            # Use the already determined connection point for the start of the line
            start_point = QPointF(0, start_rect.height() / 2) # Default bottom center in local coords
            if isinstance(self.main_window.start_item, DecisionItem):
                is_true = self.main_window.start_item.step_data.get('if_branch') == 'true'
                is_false = self.main_window.start_item.step_data.get('if_branch') == 'false'
                if is_true:
                    start_point = QPointF(-DECISION_WIDTH/2, 0)
                elif is_false:
                    start_point = QPointF(DECISION_WIDTH/2, 0)
                else:
                    start_point = QPointF(0, DECISION_HEIGHT/2)
            
            start_point_scene = self.main_window.start_item.mapToScene(start_point)
            
            path = QPainterPath()
            path.moveTo(start_point_scene)
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
        self.global_variables = {}
        self.flow_steps = []

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
        self.execute_all_button = QPushButton("Execute All Steps"); self.execute_all_button.setStyleSheet(button_style)
        left_panel_layout.addWidget(self.execute_all_button)
        self.execute_one_step_button = QPushButton("Execute 1 Step"); self.execute_one_step_button.setStyleSheet(button_style); self.execute_one_step_button.setEnabled(False)
        left_panel_layout.addWidget(self.execute_one_step_button)
        self.add_loop_button = QPushButton("Add Loop"); self.add_loop_button.setStyleSheet(button_style)
        left_panel_layout.addWidget(self.add_loop_button)
        self.add_conditional_button = QPushButton("Add Conditional"); self.add_conditional_button.setStyleSheet(button_style)
        left_panel_layout.addWidget(self.add_conditional_button)
        self.save_steps_button = QPushButton("Save Bot"); self.save_steps_button.setStyleSheet(button_style)
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
        self.clear_log_button = QPushButton("Clear Log"); log_layout.addWidget(self.clear_log_button)
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
        self._update_variables_list_display()

    def _get_item_by_id(self, item_id):
        for item in self.scene.items():
            if isinstance(item, FlowchartItem) and hasattr(item, 'step_id') and item.step_id == item_id:
                return item
        return None

    def add_step(self):
        module_dialog = ModuleSelectionDialog(self)
        if not module_dialog.exec(): return
            
        method_info = module_dialog.selected_method_info
        if not method_info: return

        _, class_name, method_name, params = method_info
        param_dialog = ParameterInputDialog(f"{class_name}.{method_name}", params, list(self.global_variables.keys()), self)
        if not param_dialog.exec(): return
        
        config, assign_to = param_dialog.get_parameters_config()
        if config is None: return

        if assign_to and assign_to not in self.global_variables:
            self.global_variables[assign_to] = None
            self._update_variables_list_display()

        step_data = {'class': class_name, 'method': method_name, 'config': config, 'assign_to': assign_to, 'status': ''}
        
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
                    
                    # Inherit branch properties if inserting into a branch
                    if target_item and target_item.step_data.get('if_branch'):
                        new_item.step_data['if_branch'] = target_item.step_data['if_branch']
                        new_item.step_data['if_id'] = target_item.step_data['if_id']
                        
                        # Also inherit parent branch properties for nested branches
                        if target_item.step_data.get('parent_if_branch'):
                            new_item.step_data['parent_if_branch'] = target_item.step_data['parent_if_branch']
                            new_item.step_data['parent_if_id'] = target_item.step_data['parent_if_id']

                    if mode == 'after':
                        self.flow_steps.insert(target_idx + 1, new_item.step_id)
                    else: # 'before'
                        self.flow_steps.insert(target_idx, new_item.step_id)
                except ValueError:
                    self.flow_steps.append(new_item.step_id)

            self._redraw_flowchart()

    def add_decision_branch(self):
        dialog = ConditionalConfigDialog(self.global_variables, self)
        if not dialog.exec(): return

        config = dialog.get_config()
        if not config: return
        
        if_id = f"if_{uuid.uuid4().hex[:6]}"
        if_data = {"type": "IF_START", "if_id": if_id, "condition_config": config}
        true_data = {'type': 'step', 'class': 'Bot_utility', 'method': 'wait_ms', 'if_branch': 'true', 'if_id': if_id, 'config': {'value': {'type': 'hardcoded', 'value': 1}}, 'assign_to': None}
        false_data = {'type': 'step', 'class': 'Bot_utility', 'method': 'wait_ms', 'if_branch': 'false', 'if_id': if_id, 'config': {'value': {'type': 'hardcoded', 'value': 1}}, 'assign_to': None}
        end_if_data = {"type": "IF_END", "if_id": if_id}

        if_item = DecisionItem(text="IF...", step_data=if_data)
        true_item = StepItem(text="True Branch Start (Placeholder)", step_data=true_data)
        false_item = StepItem(text="False Branch Start (Placeholder)", step_data=false_data)
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
                
                if s_type == 'IF_START':
                    s_outer_if_id = s_item.step_data.get('outer_if_id')
                    if s_outer_if_id == if_id and nesting_level == 0:
                        parent_branch = s_item.step_data.get('if_branch')
                        if parent_branch == 'true':
                            true_branch_ids.append(s_id)
                        elif parent_branch == 'false':
                            false_branch_ids.append(s_id)
                    nesting_level += 1
                elif s_type == 'IF_END':
                    if nesting_level == 0 and s_if_id == if_id:
                        end_if_idx = j
                        break
                    elif nesting_level > 0:
                        s_outer_if_id = s_item.step_data.get('outer_if_id')
                        if s_outer_if_id == if_id and nesting_level == 1:
                            parent_branch = s_item.step_data.get('if_branch')
                            if parent_branch == 'true':
                                true_branch_ids.append(s_id)
                            elif parent_branch == 'false':
                                false_branch_ids.append(s_id)
                        nesting_level -= 1
                elif nesting_level == 0 and s_if_id == if_id:
                    if s_item.step_data.get('if_branch') == 'true':
                        true_branch_ids.append(s_id)
                    elif s_item.step_data.get('if_branch') == 'false':
                        false_branch_ids.append(s_id)
                elif nesting_level > 0:
                    parent_if_id = s_item.step_data.get('parent_if_id')
                    if parent_if_id == if_id:
                        parent_branch = s_item.step_data.get('parent_if_branch')
                        if parent_branch == 'true':
                            true_branch_ids.append(s_id)
                        elif parent_branch == 'false':
                            false_branch_ids.append(s_id)
            
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
                    nested_end_if, current_y_true = process_if_block(true_item, current_y_true, true_x)
                    self._add_connector(last_true_item, true_item, "True" if last_true_item == if_item else "", TRUE_BRANCH_COLOR_LINE)
                    last_true_item = nested_end_if if nested_end_if else true_item
                    
                    # Skip all items that belong to this nested IF block
                    nested_if_id = true_item.step_data['if_id']
                    while i < len(true_branch_ids):
                        check_item = item_map[true_branch_ids[i]]
                        # Stop when we reach the END_IF of the nested block
                        if (check_item.step_data.get('type') == 'IF_END' and 
                            check_item.step_data.get('if_id') == nested_if_id):
                            i += 1
                            break
                        i += 1
                elif true_item.step_data.get('type') == 'IF_END':
                    # Skip END_IF blocks as they're handled by their IF_START
                    i += 1
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
                    nested_end_if, current_y_false = process_if_block(false_item, current_y_false, false_x)
                    self._add_connector(last_false_item, false_item, "False" if last_false_item == if_item else "", FALSE_BRANCH_COLOR_LINE)
                    last_false_item = nested_end_if if nested_end_if else false_item
                    
                    # Skip all items that belong to this nested IF block
                    nested_if_id = false_item.step_data['if_id']
                    while i < len(false_branch_ids):
                        check_item = item_map[false_branch_ids[i]]
                        # Stop when we reach the END_IF of the nested block
                        if (check_item.step_data.get('type') == 'IF_END' and 
                            check_item.step_data.get('if_id') == nested_if_id):
                            i += 1
                            break
                        i += 1
                elif false_item.step_data.get('type') == 'IF_END':
                    # Skip END_IF blocks as they're handled by their IF_START
                    i += 1
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
            self.scene.clear()
            self.flow_steps.clear()
            self.start_item = None
            self.btn_connect_manual.setChecked(False)

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
                self.main_window.start_item.setPen(pen)
                self.main_window.start_item = None

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
                self.global_variables[name] = value
                self._update_variables_list_display()
    
    def edit_variable(self):
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
                
                self.global_variables[new_name] = new_value
                self._update_variables_list_display()

    def delete_variable(self):
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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FlowchartApp()
    window.show()
    sys.exit(app.exec())
