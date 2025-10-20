import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QGraphicsView, QGraphicsScene, QToolBar,
    QInputDialog, QGraphicsPolygonItem, QGraphicsTextItem, QGraphicsLineItem
)
from PyQt6.QtGui import QPolygonF, QBrush, QPen, QAction, QFont
from PyQt6.QtCore import QPointF, Qt, QRectF

# --- Configuration ---
STEP_WIDTH = 120
STEP_HEIGHT = 60
DECISION_WIDTH = 150
DECISION_HEIGHT = 75
DEFAULT_BRUSH = QBrush(Qt.GlobalColor.cyan)
HOVER_BRUSH = QBrush(Qt.GlobalColor.lightGray)
LINE_PEN = QPen(Qt.GlobalColor.black, 2)
HIGHLIGHT_PEN = QPen(Qt.GlobalColor.red, 3)

# --- Custom Graphics Items ---

class Connector(QGraphicsLineItem):
    """A line that connects two FlowchartItems."""
    def __init__(self, start_item, end_item, parent=None):
        super().__init__(parent)
        self.start_item = start_item
        self.end_item = end_item
        self.setPen(LINE_PEN)
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
        self.setBrush(DEFAULT_BRUSH)
        self.setPen(QPen(Qt.GlobalColor.black, 2))
        self.setAcceptHoverEvents(True)
        self.setFlag(QGraphicsPolygonItem.GraphicsItemFlag.ItemIsMovable)
        self.setFlag(QGraphicsPolygonItem.GraphicsItemFlag.ItemIsSelectable)
        self.setFlag(QGraphicsPolygonItem.GraphicsItemFlag.ItemSendsGeometryChanges)

        self.text_item = QGraphicsTextItem(text, self)
        font = QFont("Arial", 10)
        self.text_item.setFont(font)
        self.center_text()

    def center_text(self):
        text_rect = self.text_item.boundingRect()
        polygon_rect = self.boundingRect()
        self.text_item.setPos(
            (polygon_rect.width() - text_rect.width()) / 2,
            (polygon_rect.height() - text_rect.height()) / 2
        )

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
        self.setBrush(HOVER_BRUSH)
        super().hoverEnterEvent(event)

    def hoverLeaveEvent(self, event):
        self.setBrush(DEFAULT_BRUSH)
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

# --- KEY CHANGE 1: Custom QGraphicsView ---
# We create a custom view to handle mouse events for connecting items.
class FlowchartView(QGraphicsView):
    def __init__(self, scene, parent=None):
        super().__init__(scene, parent)
        self.main_window = parent # Store reference to the main window to access its state

    def mousePressEvent(self, event):
        # Handle connection logic only if in connection mode
        if self.main_window.connection_mode and event.button() == Qt.MouseButton.LeftButton:
            item = self.itemAt(event.pos()) # Use event.pos() which is in the view's coordinates
            
            if isinstance(item, FlowchartItem):
                if self.main_window.start_item is None:
                    # This is the first item clicked
                    self.main_window.start_item = item
                    item.setPen(HIGHLIGHT_PEN)
                    print(f"Start item selected: {item.text_item.toPlainText()}")
                else:
                    # This is the second item, create the connector
                    if self.main_window.start_item != item:
                        connector = Connector(self.main_window.start_item, item)
                        self.main_window.start_item.add_connector(connector)
                        item.add_connector(connector)
                        self.scene().addItem(connector)
                        connector.update_position()
                        print(f"Connected to: {item.text_item.toPlainText()}")
                    
                    # Reset for the next connection
                    self.main_window.start_item.setPen(QPen(Qt.GlobalColor.black, 2))
                    self.main_window.start_item = None
                    # Optionally, turn off connection mode after one connection
                    # self.main_window.findChild(QAction, "connect_action").setChecked(False)
            
        else:
            # If not in connection mode, pass the event to the default handler
            # This allows for normal item selection and moving.
            super().mousePressEvent(event)

# --- Main Application Window ---
class FlowchartApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyQt6 Flowchart Creator")
        self.setGeometry(100, 100, 1000, 750)

        self.connection_mode = False
        self.start_item = None

        self.scene = QGraphicsScene()
        self.scene.setSceneRect(0, 0, 980, 730)
        
        # --- KEY CHANGE 2: Use the custom view ---
        self.view = FlowchartView(self.scene, self)
        self.view.setRenderHint(self.view.renderHints().Antialiasing)
        self.setCentralWidget(self.view)
        
        toolbar = QToolBar("Main Toolbar")
        self.addToolBar(toolbar)

        action_step = QAction("Add Step", self)
        action_step.triggered.connect(self.add_step)
        toolbar.addAction(action_step)

        action_decision = QAction("Add Decision", self)
        action_decision.triggered.connect(self.add_decision)
        toolbar.addAction(action_decision)
        
        action_connect = QAction("Connect", self)
        action_connect.setCheckable(True)
        action_connect.triggered.connect(self.toggle_connection_mode)
        action_connect.setObjectName("connect_action") # Give it a name to find it later
        toolbar.addAction(action_connect)

    def add_step(self):
        item = StepItem("New Step")
        item.setPos(100, 100)
        self.scene.addItem(item)
        
    def add_decision(self):
        item = DecisionItem("New Decision")
        item.setPos(200, 200)
        self.scene.addItem(item)
    
    def toggle_connection_mode(self, checked):
        self.connection_mode = checked
        if checked:
            self.start_item = None
            self.setCursor(Qt.CursorShape.CrossCursor)
        else:
            self.setCursor(Qt.CursorShape.ArrowCursor)
            if self.start_item:
                self.start_item.setPen(QPen(Qt.GlobalColor.black, 2))
                self.start_item = None

    # --- KEY CHANGE 3: The mousePressEvent is REMOVED from QMainWindow ---
    # The logic is now correctly placed in FlowchartView.

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FlowchartApp()
    window.show()
    sys.exit(app.exec())
