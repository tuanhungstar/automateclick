import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QGraphicsView, QGraphicsScene, QWidget,
    QInputDialog, QGraphicsPolygonItem, QGraphicsTextItem, QGraphicsLineItem,
    QMessageBox, QPushButton, QVBoxLayout, QHBoxLayout, QFrame
)
from PyQt6.QtGui import QPolygonF, QBrush, QPen, QFont, QColor
from PyQt6.QtCore import QPointF, Qt, QRectF

# --- Configuration ---
STEP_WIDTH = 140
STEP_HEIGHT = 70
DECISION_WIDTH = 160
DECISION_HEIGHT = 80
VERTICAL_SPACING = 50
HORIZONTAL_SPACING = 80

# --- NEW: Bright Mode Style Configuration ---
BG_COLOR = "#F0F0F0"
LEFT_PANEL_COLOR = "#E0E0E0"
CANVAS_COLOR = "#FFFFFF"
BUTTON_COLOR = "#DCDCDC"
BUTTON_TEXT_COLOR = "#000000"
SHAPE_FILL_COLOR = "#FFFFFF"
SHAPE_BORDER_COLOR = "#333333"
LINE_COLOR = "#333333"
HIGHLIGHT_COLOR = "#0078D7" # A bright blue for highlighting

# --- Custom Graphics Items ---

class Connector(QGraphicsLineItem):
    """A line that connects two FlowchartItems."""
    def __init__(self, start_item, end_item, parent=None):
        super().__init__(parent)
        self.start_item = start_item
        self.end_item = end_item
        self.setPen(QPen(QColor(LINE_COLOR), 2))
        self.setZValue(-1)

    def update_position(self):
        start_pos = self.start_item.scenePos()
        end_pos = self.end_item.scenePos()
        self.setLine(start_pos.x(), start_pos.y(), end_pos.x(), end_pos.y())

class FlowchartItem(QGraphicsPolygonItem):
    """Base class for all flowchart shapes."""
    def __init__(self, polygon, text="Default", parent=None):
        super().__init__(polygon, parent)
        self.connectors = []
        self.setBrush(QBrush(QColor(SHAPE_FILL_COLOR)))
        self.setPen(QPen(QColor(SHAPE_BORDER_COLOR), 2))
        self.setAcceptHoverEvents(True)
        self.setFlag(QGraphicsPolygonItem.GraphicsItemFlag.ItemIsMovable)
        self.setFlag(QGraphicsPolygonItem.GraphicsItemFlag.ItemIsSelectable)
        self.setFlag(QGraphicsPolygonItem.GraphicsItemFlag.ItemSendsGeometryChanges)

        self.text_item = QGraphicsTextItem(text, self)
        font = QFont("Arial", 11)
        # Black text for light background
        self.text_item.setDefaultTextColor(QColor(BUTTON_TEXT_COLOR))
        self.text_item.setFont(font)
        self.text_item.setTextWidth(self.boundingRect().width() * 0.9)
        self.center_text()

    def center_text(self):
        text_rect = self.text_item.boundingRect()
        new_x = -text_rect.width() / 2
        new_y = -text_rect.height() / 2
        self.text_item.setPos(new_x, new_y)

    def add_connector(self, connector):
        self.connectors.append(connector)

    def itemChange(self, change, value):
        if change == QGraphicsPolygonItem.GraphicsItemChange.ItemPositionHasChanged:
            for connector in self.connectors:
                connector.update_position()
        return super().itemChange(change, value)
    
    def mouseDoubleClickEvent(self, event):
        text, ok = QInputDialog.getText(None, "Edit Text", "Enter new text:", text=self.text_item.toPlainText())
        if ok and text:
            self.text_item.setPlainText(text)
            self.center_text()
        super().mouseDoubleClickEvent(event)

    def hoverEnterEvent(self, event):
        self.setPen(QPen(QColor(HIGHLIGHT_COLOR), 3))
        super().hoverEnterEvent(event)

    def hoverLeaveEvent(self, event):
        self.setPen(QPen(QColor(SHAPE_BORDER_COLOR), 2))
        super().hoverLeaveEvent(event)

class StepItem(FlowchartItem):
    def __init__(self, text="Step", parent=None):
        rect = QPolygonF([
            QPointF(-STEP_WIDTH / 2, -STEP_HEIGHT / 2), QPointF(STEP_WIDTH / 2, -STEP_HEIGHT / 2),
            QPointF(STEP_WIDTH / 2, STEP_HEIGHT / 2), QPointF(-STEP_WIDTH / 2, STEP_HEIGHT / 2),
        ])
        super().__init__(rect, text, parent)

class DecisionItem(FlowchartItem):
    def __init__(self, text="If...", parent=None):
        diamond = QPolygonF([
            QPointF(0, -DECISION_HEIGHT / 2), QPointF(DECISION_WIDTH / 2, 0),
            QPointF(0, DECISION_HEIGHT / 2), QPointF(-DECISION_WIDTH / 2, 0),
        ])
        super().__init__(diamond, text, parent)

# --- Custom QGraphicsView ---
class FlowchartView(QGraphicsView):
    def __init__(self, scene, parent=None):
        super().__init__(scene, parent)
        self.main_window = parent
        self.setStyleSheet(f"background-color: {CANVAS_COLOR}; border: 1px solid #C0C0C0;")

    def mousePressEvent(self, event):
        if self.main_window.connection_mode and event.button() == Qt.MouseButton.LeftButton:
            item = self.itemAt(event.pos())
            if isinstance(item, FlowchartItem):
                if self.main_window.start_item is None:
                    self.main_window.start_item = item
                    item.setPen(QPen(QColor(HIGHLIGHT_COLOR), 3))
                else:
                    if self.main_window.start_item != item:
                        connector = Connector(self.main_window.start_item, item)
                        self.main_window.start_item.add_connector(connector)
                        item.add_connector(connector)
                        self.scene().addItem(connector)
                        connector.update_position()
                    self.main_window.start_item.setPen(QPen(QColor(SHAPE_BORDER_COLOR), 2))
                    self.main_window.start_item = None
        else:
            super().mousePressEvent(event)

# --- Main Application Window ---
class FlowchartApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Flowchart Creator")
        self.setGeometry(100, 100, 1200, 800)
        self.setStyleSheet(f"background-color: {BG_COLOR};")

        self.connection_mode = False
        self.start_item = None
        self.last_item = None
        self.step_counter = 0

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        left_panel = QFrame()
        left_panel.setFixedWidth(200)
        left_panel.setStyleSheet(f"background-color: {LEFT_PANEL_COLOR};")
        left_panel_layout = QVBoxLayout(left_panel)
        left_panel_layout.setContentsMargins(10, 10, 10, 10)
        left_panel_layout.setSpacing(10)

        button_style = f"""
            QPushButton {{
                background-color: {BUTTON_COLOR};
                color: {BUTTON_TEXT_COLOR};
                border: 1px solid #B0B0B0;
                padding: 10px;
                font-size: 13px;
                border-radius: 5px;
            }}
            QPushButton:hover {{
                background-color: #E6E6E6;
                border-color: #999999;
            }}
            QPushButton:pressed {{
                background-color: #C0C0C0;
            }}
            QPushButton:checkable:checked {{
                background-color: {HIGHLIGHT_COLOR};
                color: #FFFFFF;
                border-color: #005A9E;
            }}
        """
        
        btn_add_step = QPushButton("Add Step")
        btn_add_step.setStyleSheet(button_style)
        btn_add_step.clicked.connect(self.add_step)
        left_panel_layout.addWidget(btn_add_step)

        btn_add_decision = QPushButton("Add Decision Branch")
        btn_add_decision.setStyleSheet(button_style)
        btn_add_decision.clicked.connect(self.add_decision_branch)
        left_panel_layout.addWidget(btn_add_decision)
        
        self.btn_connect_manual = QPushButton("Connect Manually")
        self.btn_connect_manual.setStyleSheet(button_style)
        self.btn_connect_manual.setCheckable(True)
        self.btn_connect_manual.toggled.connect(self.toggle_connection_mode)
        left_panel_layout.addWidget(self.btn_connect_manual)
        
        btn_clear = QPushButton("Clear All")
        btn_clear.setStyleSheet(button_style)
        btn_clear.clicked.connect(self.clear_all)
        left_panel_layout.addWidget(btn_clear)
        
        left_panel_layout.addStretch()

        self.scene = QGraphicsScene()
        self.view = FlowchartView(self.scene, self)
        self.view.setRenderHint(self.view.renderHints().Antialiasing)
        self.view.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.view.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        main_layout.addWidget(left_panel)
        main_layout.addWidget(self.view)

    def _get_next_position_y(self, previous_item, new_item_height=0):
        if not previous_item:
            return 50
        prev_pos = previous_item.scenePos()
        prev_height = previous_item.boundingRect().height()
        return prev_pos.y() + (prev_height / 2) + VERTICAL_SPACING + (new_item_height / 2)

    def _add_and_connect_item(self, item_class, text, x_pos, y_pos, connect_from_item=None):
        self.step_counter += 1
        shape_text = f"Step {self.step_counter}: {text}"
        new_item = item_class(shape_text)
        new_item.setPos(x_pos, y_pos)
        self.scene.addItem(new_item)
        if connect_from_item:
            connector = Connector(connect_from_item, new_item)
            connect_from_item.add_connector(connector)
            new_item.add_connector(connector)
            self.scene.addItem(connector)
            connector.update_position()
        self.view.ensureVisible(new_item)
        return new_item

    def add_step(self):
        initial_x = self.view.mapToScene(self.view.viewport().rect().center()).x()
        initial_y = self._get_next_position_y(self.last_item, STEP_HEIGHT)
        new_item = self._add_and_connect_item(StepItem, "New Step", initial_x, initial_y, self.last_item)
        self.last_item = new_item
        
    def add_decision_branch(self):
        initial_x = self.view.mapToScene(self.view.viewport().rect().center()).x()
        initial_y = self._get_next_position_y(self.last_item, DECISION_HEIGHT)
        decision_item = self._add_and_connect_item(DecisionItem, "New Decision", initial_x, initial_y, self.last_item)
        y_for_parallel_steps = self._get_next_position_y(decision_item, STEP_HEIGHT)
        left_step_x = initial_x - (STEP_WIDTH / 2) - (HORIZONTAL_SPACING / 2)
        right_step_x = initial_x + (STEP_WIDTH / 2) + (HORIZONTAL_SPACING / 2)
        step_left = self._add_and_connect_item(StepItem, "New Step", left_step_x, y_for_parallel_steps, decision_item)
        step_right = self._add_and_connect_item(StepItem, "New Step", right_step_x, y_for_parallel_steps, decision_item)
        lowest_y_of_parallel = max(step_left.scenePos().y() + step_left.boundingRect().height() / 2,
                                   step_right.scenePos().y() + step_right.boundingRect().height() / 2)
        final_step_y = lowest_y_of_parallel + VERTICAL_SPACING + (STEP_HEIGHT / 2)
        final_step = self._add_and_connect_item(StepItem, "New Step", initial_x, final_step_y)
        connector_left_to_final = Connector(step_left, final_step)
        step_left.add_connector(connector_left_to_final)
        final_step.add_connector(connector_left_to_final)
        self.scene.addItem(connector_left_to_final)
        connector_left_to_final.update_position()
        connector_right_to_final = Connector(step_right, final_step)
        step_right.add_connector(connector_right_to_final)
        final_step.add_connector(connector_right_to_final)
        self.scene.addItem(connector_right_to_final)
        connector_right_to_final.update_position()
        self.last_item = final_step
        self.view.ensureVisible(final_step)

    def clear_all(self):
        reply = QMessageBox.question(self, "Confirm Clear",
                                     "Are you sure you want to remove everything?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.scene.clear()
            self.last_item = None
            self.start_item = None
            self.step_counter = 0
            self.btn_connect_manual.setChecked(False)

    def toggle_connection_mode(self, checked):
        self.connection_mode = checked
        if checked:
            self.start_item = None
            self.setCursor(Qt.CursorShape.CrossCursor)
        else:
            self.setCursor(Qt.CursorShape.ArrowCursor)
            if self.start_item:
                self.start_item.setPen(QPen(QColor(SHAPE_BORDER_COLOR), 2))
                self.start_item = None

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FlowchartApp()
    window.show()
    sys.exit(app.exec())
