# File: Bot_module/code_module.py

import sys
from typing import Optional, List, Dict, Any

# --- PyQt6 Imports ---
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QWidget, QMessageBox, QLabel,
    QPlainTextEdit, QListWidget, QHBoxLayout, QSplitter,
    QListWidgetItem, QDialogButtonBox,QTextEdit  
)
from PyQt6.QtGui import QSyntaxHighlighter, QTextCharFormat, QColor, QFont, QTextDocument, QPainter
from PyQt6.QtCore import Qt, QRect, QSize

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

# --- NEW: Widget for Line Numbers ---
class _LineNumberArea(QWidget):
    """Widget to display line numbers for the code editor."""
    def __init__(self, editor: QPlainTextEdit):
        super().__init__(editor)
        self.code_editor = editor

    def sizeHint(self) -> QSize:
        return QSize(self.code_editor.lineNumberAreaWidth(), 0)

    def paintEvent(self, event: "QPaintEvent") -> None:
        self.code_editor.lineNumberAreaPaintEvent(event)

# --- NEW: Code Editor with Line Numbers and Tab Settings ---
class _CodeEditor(QPlainTextEdit):
    """A QPlainTextEdit with line numbers and custom tab width."""
    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.line_number_area = _LineNumberArea(self)

        self.blockCountChanged.connect(self.updateLineNumberAreaWidth)
        self.updateRequest.connect(self.updateLineNumberArea)
        self.cursorPositionChanged.connect(self.highlightCurrentLine)

        self.updateLineNumberAreaWidth(0)
        self.highlightCurrentLine()
        
        # --- ENHANCEMENT: Set tab width to 5 spaces ---
        # Calculate the width of a single space character in the current font
        font_metrics = self.fontMetrics()
        space_width = font_metrics.horizontalAdvance(' ')
        # Set the tab stop distance to 5 times the width of a space
        self.setTabStopDistance(5 * space_width)

    # ... (Keep all other methods in _CodeEditor: lineNumberAreaWidth, updateLineNumberAreaWidth, etc. They are correct)
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

# --- REWRITTEN DIALOG WITH NEW UI AND FUNCTIONALITY ---
class _CodeExecutorDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None, initial_config: Optional[Dict[str, Any]] = None, **kwargs):
        super().__init__(parent)
        self.setWindowTitle("Python Code Executor")
        self.setMinimumSize(950, 700)
        
        # --- ENHANCEMENT: Add maximize button to the dialog window ---
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowMaximizeButtonHint)
        
        self.global_variables = global_variables
        
        # --- Main Layout ---
        main_layout = QVBoxLayout(self)
        
        info_label = QLabel("Write Python code in the editor. Double-click variables from the right panel to use them.")
        info_label.setStyleSheet("font-style: italic; color: #555;")
        main_layout.addWidget(info_label)

        # Main splitter: Code Editor on Left | Variables on Right
        main_splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(main_splitter, 1)

        # --- Left Side: Code Editor (Uses the modified _CodeEditor class) ---
        self.code_editor = _CodeEditor()
        self.code_editor.setFont(QFont("Courier New", 10))
        self.code_editor.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.highlighter = _PythonHighlighter(self.code_editor.document())
        main_splitter.addWidget(self.code_editor)

        # --- Right Side: Variable Lists ---
        variables_panel = QWidget()
        variables_layout = QVBoxLayout(variables_panel)
        variables_layout.setContentsMargins(0, 0, 0, 0)
        variables_layout.setSpacing(10)

        # Input Variables
        variables_layout.addWidget(QLabel("<b>Input Variables</b> (Double-click to get value)"))
        self.input_vars_list = QListWidget()
        self.input_vars_list.itemDoubleClicked.connect(self._insert_get_variable)
        variables_layout.addWidget(self.input_vars_list)
        
        # Output Variables
        variables_layout.addWidget(QLabel("<b>Output Variables</b> (Double-click to set value)"))
        self.output_vars_list = QListWidget()
        self.output_vars_list.itemDoubleClicked.connect(self._insert_set_variable)
        variables_layout.addWidget(self.output_vars_list)
        
        main_splitter.addWidget(variables_panel)
        main_splitter.setSizes([700, 250])

        # --- Dialog Buttons ---
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        # --- Populate Data ---
        self._populate_variable_lists()
        if initial_config:
            self._populate_from_initial_config(initial_config)
            
    def _populate_variable_lists(self):
        """Populates the input and output variable lists."""
        # BUG FIX: Ensure only strings are sorted and added to the list widget.
        # This prevents the TypeError if the global_variables list contains non-string items.
        str_variables = sorted([v for v in self.global_variables if isinstance(v, str)])
        self.input_vars_list.addItems(str_variables)
        self.output_vars_list.addItems(str_variables)

    def _insert_get_variable(self, item: QListWidgetItem):
        """Inserts context.get_variable('var_name') into the code editor."""
        variable_name = item.text()
        text_to_insert = f"context.get_variable('{variable_name}')"
        self.code_editor.insertPlainText(text_to_insert)
        self.code_editor.setFocus()

    def _insert_set_variable(self, item: QListWidgetItem):
        """Inserts context.set_variable('var_name', ) into the code editor."""
        variable_name = item.text()
        text_to_insert = f"context.set_variable('{variable_name}', )"
        
        cursor = self.code_editor.textCursor()
        cursor.insertText(text_to_insert)
        
        # Move cursor back one position to be inside the parentheses
        cursor.movePosition(cursor.MoveOperation.Left, n=1)
        self.code_editor.setTextCursor(cursor)
        self.code_editor.setFocus()
        
    def _populate_from_initial_config(self, config: Dict[str, Any]):
        """Populates the UI from a saved configuration."""
        self.code_editor.setPlainText(config.get("code_string", ""))
        
    def get_executor_method_name(self) -> str:
        """Returns the name of the method that will execute the code."""
        return "_execute_python_code"

    def get_config_data(self) -> Optional[Dict[str, Any]]:
        """Returns the configuration data to be saved."""
        code_string = self.code_editor.toPlainText()
        if not code_string.strip():
            QMessageBox.warning(self, "Input Error", "Code cannot be empty.")
            return None
        # The assignment variable is no longer needed here as per the new design
        return {"code_string": code_string}

    def get_assignment_variable(self) -> Optional[str]:
        """This method is now obsolete with the new design but kept for compatibility."""
        return None

#
# --- Public-Facing Module Class for the Code Executor ---
#
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
        return _CodeExecutorDialog(
            global_variables,
            parent_window,
            **kwargs
        )

    def _execute_python_code(self, context: ExecutionContext, config_data: dict) -> Any:
        """
        Executes the user's Python code in a controlled environment.
        The return value of the last expression is returned, but using context.set_variable() is preferred.
        """
        self.context = context
        code_to_run = config_data.get("code_string", "")
        self._log("Executing custom Python code...")
        
        local_scope = {'context': context}
        
        # Add existing global variables to the scope for direct access.
        if hasattr(context, 'global_variables_ref') and isinstance(context.global_variables_ref, dict):
            local_scope.update(context.global_variables_ref)

        try:
            exec(code_to_run, globals(), local_scope)
            
            # After execution, update the main context with any changes from the local scope.
            if hasattr(context, 'global_variables_ref'):
                for var_name, value in local_scope.items():
                    if var_name not in ['__builtins__', 'context'] and var_name in context.global_variables_ref:
                         context.set_variable(var_name, value)

            # Determine the result from the last expression, if any.
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
            # Note: This return value is often ignored if you use `context.set_variable`.
            # For your old system, this used to be assigned to the output variable.
            return final_result

        except Exception as e:
            error_message = f"Error during Python code execution: {type(e).__name__}: {e}"
            self._log(f"FATAL ERROR: {error_message}")
            raise
