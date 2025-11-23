# File: Bot_module/code_module.py

import sys
import os
import inspect
import importlib.util
from typing import Optional, List, Dict, Any

# --- PyQt6 Imports ---
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QWidget, QMessageBox, QLabel,
    QPlainTextEdit, QHBoxLayout, QSplitter,
    QDialogButtonBox, QTextEdit, QPushButton, QFileDialog,
    QTreeWidget, QTreeWidgetItem, QListWidget, QListWidgetItem
)
from PyQt6.QtGui import QSyntaxHighlighter, QTextCharFormat, QColor, QFont, QTextDocument, QPainter
from PyQt6.QtCore import Qt, QRect, QSize, QVariant

# --- Main App Imports (Fallback for standalone testing) ---
try:
    from my_lib.shared_context import ExecutionContext
except ImportError:
    print("Warning: Could not import main app libraries. Using fallbacks for Code Executor.")
    class ExecutionContext:
        def add_log(self, message: str): print(f"LOG: {message}")
        def get_variable(self, name: str, default: Any = None) -> Any:
            print(f"Fallback: Getting variable '{name}'")
            return default
        def set_variable(self, name: str, value: Any):
            print(f"Fallback: Setting variable '{name}' to {value}")

class _PythonHighlighter(QSyntaxHighlighter):
    """Syntax highlighter for Python code."""
    def __init__(self, parent: QTextDocument):
        super().__init__(parent)
        self.highlighting_rules = []
        keyword_format = QTextCharFormat()
        keyword_format.setForeground(QColor("#000080")) # Dark Blue
        keyword_format.setFontWeight(QFont.Weight.Bold)
        keywords = [
            "and", "as", "assert", "break", "class", "continue", "def", "del",
            "elif", "else", "except", "False", "finally", "for", "from", "global",
            "if", "import", "in", "is", "lambda", "None", "nonlocal", "not",
            "or", "pass", "raise", "return", "True", "try", "while", "with", "yield"
        ]
        self.highlighting_rules += [(f"\\b{keyword}\\b", keyword_format) for keyword in keywords]

        string_format = QTextCharFormat(); string_format.setForeground(QColor("#008000")) # Green
        self.highlighting_rules.append(("'[^']*'", string_format))
        self.highlighting_rules.append(('"[^"]*"', string_format))

        comment_format = QTextCharFormat(); comment_format.setForeground(QColor("#808080")); comment_format.setFontItalic(True) # Grey Italic
        self.highlighting_rules.append(("#[^\n]*", comment_format))

        number_format = QTextCharFormat(); number_format.setForeground(QColor("#A52A2A")) # Brown
        self.highlighting_rules.append(("\\b[0-9]+\\.?[0-9]*\\b", number_format))

        function_format = QTextCharFormat(); function_format.setForeground(QColor("#800080")) # Purple
        self.highlighting_rules.append(("(\\w+)\\s*\\(", function_format))

    def highlightBlock(self, text):
        from PyQt6.QtCore import QRegularExpression
        for pattern, format in self.highlighting_rules:
            expression = QRegularExpression(pattern)
            it = expression.globalMatch(text)
            while it.hasNext():
                match = it.next()
                self.setFormat(match.capturedStart(), match.capturedLength(), format)

# --- Widget for Line Numbers ---
class _LineNumberArea(QWidget):
    def __init__(self, editor: QPlainTextEdit):
        super().__init__(editor)
        self.code_editor = editor

    def sizeHint(self) -> QSize:
        return QSize(self.code_editor.lineNumberAreaWidth(), 0)

    def paintEvent(self, event: "QPaintEvent") -> None:
        self.code_editor.lineNumberAreaPaintEvent(event)

# --- Code Editor with Line Numbers and Tab Settings ---
class _CodeEditor(QPlainTextEdit):
    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.line_number_area = _LineNumberArea(self)

        self.blockCountChanged.connect(self.updateLineNumberAreaWidth)
        self.updateRequest.connect(self.updateLineNumberArea)
        self.cursorPositionChanged.connect(self.highlightCurrentLine)

        self.updateLineNumberAreaWidth(0)
        self.highlightCurrentLine()

        font_metrics = self.fontMetrics()
        space_width = font_metrics.horizontalAdvance(' ')
        self.setTabStopDistance(5 * space_width)

    def lineNumberAreaWidth(self) -> int:
        digits = 1
        count = max(1, self.blockCount())
        while count >= 10:
            count //= 10
            digits += 1
        space = 10 + self.fontMetrics().horizontalAdvance('9') * digits
        return space

    def updateLineNumberAreaWidth(self, new_block_count: int):
        self.setViewportMargins(self.lineNumberAreaWidth(), 0, 0, 0)

    def updateLineNumberArea(self, rect: QRect, dy: int):
        if dy:
            self.line_number_area.scroll(0, dy)
        else:
            self.line_number_area.update(0, rect.y(), self.line_number_area.width(), rect.height())

        if rect.contains(self.viewport().rect()):
            self.updateLineNumberAreaWidth(0)

    def resizeEvent(self, event: "QResizeEvent") -> None:
        super().resizeEvent(event)
        cr = self.contentsRect()
        self.line_number_area.setGeometry(QRect(cr.left(), cr.top(), self.lineNumberAreaWidth(), cr.height()))

    def lineNumberAreaPaintEvent(self, event: "QPaintEvent"):
        painter = QPainter(self.line_number_area)
        painter.fillRect(event.rect(), QColor("#f0f0f0"))
        block = self.firstVisibleBlock()
        block_number = block.blockNumber()
        top = self.blockBoundingGeometry(block).translated(self.contentOffset()).top()
        bottom = top + self.blockBoundingRect(block).height()
        while block.isValid() and top <= event.rect().bottom():
            if block.isVisible() and bottom >= event.rect().top():
                number = str(block_number + 1)
                painter.setPen(Qt.GlobalColor.darkGray)
                painter.drawText(0, int(top), self.line_number_area.width() - 5, self.fontMetrics().height(),
                                 Qt.AlignmentFlag.AlignRight, number)
            block = block.next()
            top = bottom
            bottom = top + self.blockBoundingRect(block).height()
            block_number += 1

    def highlightCurrentLine(self):
        extra_selections = []
        if not self.isReadOnly():
            selection = QTextEdit.ExtraSelection()
            line_color = QColor(Qt.GlobalColor.yellow).lighter(160)
            selection.format.setBackground(line_color)
            selection.format.setProperty(QTextCharFormat.Property.FullWidthSelection, True)
            selection.cursor = self.textCursor()
            selection.cursor.clearSelection()
            extra_selections.append(selection)
        self.setExtraSelections(extra_selections)


# --- REWRITTEN DIALOG WITH THREE-PANEL LAYOUT ---
class _CodeExecutorDialog(QDialog):
    def __init__(self, global_variables: List[str], py_file_dir: str, parent: Optional[QWidget] = None, initial_config: Optional[Dict[str, Any]] = None, **kwargs):
        super().__init__(parent)
        self.setWindowTitle("Python Code Executor")
        self.setMinimumSize(1100, 700)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowMaximizeButtonHint)

        self.global_variables = global_variables
        self.imported_module_path = None
        self.py_file_dir = py_file_dir # Store the default directory for scripts

        # --- Main Layout ---
        main_layout = QVBoxLayout(self)
        info_label = QLabel("Write code in the center. Use tools from the side panels by double-clicking.")
        info_label.setStyleSheet("font-style: italic; color: #555;")
        main_layout.addWidget(info_label)

        # --- Main 3-Panel Splitter ---
        main_splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(main_splitter, 1)

        # --- Panel 1: Left - Module Inspector ---
        inspector_panel = QWidget()
        inspector_layout = QVBoxLayout(inspector_panel)
        inspector_layout.setContentsMargins(5, 0, 5, 0)
        inspector_layout.setSpacing(10)
        
        inspector_layout.addWidget(QLabel("<b>Tools</b>"))
        
        # --- NEW: Button to load script from file ---
        self.load_script_button = QPushButton("Load Script...")
        self.load_script_button.clicked.connect(self._load_script_from_file)
        inspector_layout.addWidget(self.load_script_button)
        
        inspector_layout.addWidget(QLabel("<b>Module Inspector</b>"))
        self.import_button = QPushButton("Import .py File...")
        self.import_button.clicked.connect(self._import_module)
        inspector_layout.addWidget(self.import_button)
        
        self.imported_file_label = QLabel("No module imported.")
        self.imported_file_label.setStyleSheet("font-style: italic; color: #777;")
        self.imported_file_label.setWordWrap(True)
        inspector_layout.addWidget(self.imported_file_label)

        self.module_tree = QTreeWidget()
        self.module_tree.setHeaderHidden(True)
        self.module_tree.itemDoubleClicked.connect(self._insert_from_tree)
        inspector_layout.addWidget(self.module_tree)
        main_splitter.addWidget(inspector_panel)

        # --- Panel 2: Center - Code Editor ---
        self.code_editor = _CodeEditor()
        self.code_editor.setFont(QFont("Courier New", 10))
        self.code_editor.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.highlighter = _PythonHighlighter(self.code_editor.document())
        main_splitter.addWidget(self.code_editor)

        # --- Panel 3: Right - Variable Lists ---
        variables_panel = QWidget()
        variables_layout = QVBoxLayout(variables_panel)
        variables_layout.setContentsMargins(0, 0, 0, 0)
        variables_layout.setSpacing(10)
        
        variables_layout.addWidget(QLabel("<b>Input Variables</b> (Double-click to use)"))
        self.input_vars_list = QListWidget()
        self.input_vars_list.itemDoubleClicked.connect(self._insert_get_variable)
        variables_layout.addWidget(self.input_vars_list)
        
        variables_layout.addWidget(QLabel("<b>Output Variables</b> (Double-click to set)"))
        self.output_vars_list = QListWidget()
        self.output_vars_list.itemDoubleClicked.connect(self._insert_set_variable)
        variables_layout.addWidget(self.output_vars_list)
        main_splitter.addWidget(variables_panel)

        main_splitter.setSizes([250, 600, 250])

        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        self._populate_variable_lists()
        if initial_config:
            self._populate_from_initial_config(initial_config)

    def _populate_variable_lists(self):
        self.input_vars_list.clear()
        self.output_vars_list.clear()
        self.input_vars_list.addItem("context")
        str_variables = sorted([v for v in self.global_variables if isinstance(v, str)])
        self.input_vars_list.addItems(str_variables)
        self.output_vars_list.addItems(str_variables)

    def _insert_get_variable(self, item: QListWidgetItem):
        variable_name = item.text()
        text_to_insert = "context" if variable_name == 'context' else f"context.get_variable('{variable_name}')"
        self.code_editor.insertPlainText(text_to_insert)
        self.code_editor.setFocus()

    def _insert_set_variable(self, item: QListWidgetItem):
        variable_name = item.text()
        text_to_insert = f"context.set_variable('{variable_name}', )"
        cursor = self.code_editor.textCursor()
        cursor.insertText(text_to_insert)
        cursor.movePosition(cursor.MoveOperation.Left, n=1)
        self.code_editor.setTextCursor(cursor)
        self.code_editor.setFocus()

    def _load_script_from_file(self):
        """Opens a file dialog to load a Python script into the editor."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "Load Python Script", 
            self.py_file_dir,  # Start in the default script directory
            "Python Files (*.py);;All Files (*)"
        )
        
        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                script_content = f.read()
            self.code_editor.setPlainText(script_content)
            
        except Exception as e:
            QMessageBox.critical(self, "Load Error", f"Failed to load script file:\n\n{e}")

    def _import_module(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Import Python Module", self.py_file_dir, "Python Files (*.py)")
        if not file_path:
            return

        self.imported_module_path = file_path
        self.module_tree.clear()
        
        try:
            module_name = os.path.splitext(os.path.basename(file_path))[0]
            spec = importlib.util.spec_from_file_location(module_name, file_path)
            if not spec or not spec.loader:
                raise ImportError(f"Could not create module spec for {file_path}")
            
            module = importlib.util.module_from_spec(spec)
            sys.modules[module_name] = module
            spec.loader.exec_module(module)

            self.imported_file_label.setText(f"Imported: {os.path.basename(file_path)}")

            for name, obj in inspect.getmembers(module, inspect.isclass):
                if obj.__module__ == module_name:
                    class_item = QTreeWidgetItem(self.module_tree, [name])
                    class_item.setData(0, Qt.ItemDataRole.UserRole, QVariant(("class", name)))

                    for method_name, method_obj in inspect.getmembers(obj, inspect.isfunction):
                        if not method_name.startswith('_'):
                            sig = inspect.signature(method_obj)
                            params = [p.name for p in sig.parameters.values() if p.name != 'self']
                            param_str = ", ".join(params)
                            method_item = QTreeWidgetItem(class_item, [f"{method_name}({param_str})"])
                            method_item.setData(0, Qt.ItemDataRole.UserRole, QVariant(("method", name, method_name, params)))
            self.module_tree.expandAll()
        except Exception as e:
            self.imported_file_label.setText(f"Error importing: {os.path.basename(file_path)}")
            QMessageBox.critical(self, "Import Error", f"Failed to import and inspect module:\n\n{e}")
        finally:
            if 'module_name' in locals() and module_name in sys.modules:
                del sys.modules[module_name]

    def _insert_from_tree(self, item: QTreeWidgetItem, column: int):
        item_data_raw = item.data(0, Qt.ItemDataRole.UserRole)
        item_data = item_data_raw.value() if isinstance(item_data_raw, QVariant) else item_data_raw

        if not isinstance(item_data, tuple):
            return

        item_type = item_data[0]
        text_to_insert = ""
        cursor_offset = 0

        if item_type == 'class':
            _, class_name = item_data
            text_to_insert = f"{class_name}()"
            cursor_offset = 0
        elif item_type == 'method':
            _, class_name, method_name, params = item_data
            param_str = ", ".join(params)
            text_to_insert = f"{method_name}({param_str})"
            cursor_offset = -1 if param_str else 0

        if text_to_insert:
            cursor = self.code_editor.textCursor()
            cursor.insertText(text_to_insert)
            if cursor_offset != 0:
                cursor.movePosition(cursor.MoveOperation.Left, n=abs(cursor_offset))
            self.code_editor.setTextCursor(cursor)
            self.code_editor.setFocus()

    def _populate_from_initial_config(self, config: Dict[str, Any]):
        self.code_editor.setPlainText(config.get("code_string", ""))
        
    def get_executor_method_name(self) -> str:
        return "_execute_python_code"

    def get_config_data(self) -> Optional[Dict[str, Any]]:
        code_string = self.code_editor.toPlainText()
        if not code_string.strip():
            QMessageBox.warning(self, "Input Error", "Code cannot be empty.")
            return None
        return {
            "code_string": code_string,
            "imported_module_path": self.imported_module_path
        }

    def get_assignment_variable(self) -> Optional[str]:
        return None

# --- Public-Facing Module Class for the Code Executor ---
class Code_Executor:
    def __init__(self, context: Optional[ExecutionContext] = None):
        self.context = context

    def _log(self, message: str):
        if self.context:
            self.context.add_log(message)
        else:
            print(message)

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        self._log("Opening Python Code Executor configuration...")

        # --- NEW: Ensure the py_file directory exists ---
        base_dir = os.path.dirname(os.path.abspath(sys.argv[0] if hasattr(sys, 'frozen') else __file__))
        py_file_dir = os.path.join(base_dir, "py_file")
        os.makedirs(py_file_dir, exist_ok=True)
        # --- END NEW ---

        return _CodeExecutorDialog(
            global_variables,
            py_file_dir=py_file_dir, # Pass the directory path to the dialog
            parent=parent_window,
            **kwargs
        )

    def _execute_python_code(self, context: ExecutionContext, config_data: dict) -> Any:
        self.context = context
        code_to_run = config_data.get("code_string", "")
        imported_module_path = config_data.get("imported_module_path")
        self._log("Executing custom Python code...")

        local_scope = {'context': context}
        original_sys_path = sys.path[:]
        try:
            if imported_module_path and os.path.exists(imported_module_path):
                self._log(f"Loading module from: {imported_module_path}")
                module_dir = os.path.dirname(imported_module_path)
                if module_dir not in sys.path:
                    sys.path.insert(0, module_dir)
                
                module_name = os.path.splitext(os.path.basename(imported_module_path))[0]
                spec = importlib.util.spec_from_file_location(module_name, imported_module_path)
                if not spec or not spec.loader:
                     raise ImportError(f"Could not load spec for module {module_name}")
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                for name, obj in inspect.getmembers(module, inspect.isclass):
                    if obj.__module__ == module_name:
                        local_scope[name] = obj
                self._log(f"Made classes from '{module_name}' available to script.")

            if hasattr(context, 'global_variables_ref') and isinstance(context.global_variables_ref, dict):
                local_scope.update(context.global_variables_ref)

            exec(code_to_run, globals(), local_scope)

            if hasattr(context, 'global_variables_ref'):
                for var_name, value in local_scope.items():
                    if var_name not in ['__builtins__', 'context'] and var_name in context.global_variables_ref:
                         context.set_variable(var_name, value)

            lines = [line for line in code_to_run.strip().split('\n') if line.strip()]
            final_result = None
            if lines:
                last_line = lines[-1].strip()
                try:
                    import ast
                    parsed_code = ast.parse(last_line)
                    is_simple_expression = (isinstance(parsed_code, ast.Module) and
                                            len(parsed_code.body) == 1 and
                                            isinstance(parsed_code.body[0], ast.Expr))
                    if is_simple_expression:
                        final_result = eval(last_line, globals(), local_scope)
                except (SyntaxError, NameError, TypeError):
                    final_result = None

            self._log(f"Code execution finished. Final expression result: {type(final_result).__name__}")
            return final_result

        except Exception as e:
            error_message = f"Error during Python code execution: {type(e).__name__}: {e}"
            self._log(f"FATAL ERROR: {error_message}")
            raise
        finally:
            sys.path = original_sys_path