import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QGraphicsView, QGraphicsScene, QToolBar,
    QInputDialog, QGraphicsPolygonItem, QGraphicsTextItem, QGraphicsLineItem,
    QMessageBox
)
from PyQt6.QtGui import QPolygonF, QBrush, QPen, QAction, QFont
from PyQt6.QtCore import QPointF, Qt, QRectF

# --- Configuration ---
STEP_WIDTH = 120
STEP_HEIGHT = 60
DECISION_WIDTH = 150
DECISION_HEIGHT = 75
VERTICAL_SPACING = 50  # The space between shapes
HORIZONTAL_SPACING = 80 # Space between parallel shapes
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

# --- Custom QGraphicsView ---
class FlowchartView(QGraphicsView):
    def __init__(self, scene, parent=None):
        super().__init__(scene, parent)
        self.main_window = parent

    def mousePressEvent(self, event):
        if self.main_window.connection_mode and event.button() == Qt.MouseButton.LeftButton:
            item = self.itemAt(event.pos())
            if isinstance(item, FlowchartItem):
                if self.main_window.start_item is None:
                    self.main_window.start_item = item
                    item.setPen(HIGHLIGHT_PEN)
                else:
                    if self.main_window.start_item != item:
                        connector = Connector(self.main_window.start_item, item)
                        self.main_window.start_item.add_connector(connector)
                        item.add_connector(connector)
                        self.scene().addItem(connector)
                        connector.update_position()
                    self.main_window.start_item.setPen(QPen(Qt.GlobalColor.black, 2))
                    self.main_window.start_item = None
        else:
            super().mousePressEvent(event)

# --- Main Application Window ---
class FlowchartApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyQt6 Flowchart Creator")
        self.setGeometry(100, 100, 1000, 750)

        # --- App State ---
        self.connection_mode = False
        self.start_item = None
        self.last_item = None
        self.step_counter = 0

        self.scene = QGraphicsScene()
        
        self.view = FlowchartView(self.scene, self)
        self.view.setRenderHint(self.view.renderHints().Antialiasing)
        self.setCentralWidget(self.view)

        self.view.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.view.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        toolbar = QToolBar("Main Toolbar")
        self.addToolBar(toolbar)

        action_step = QAction("Add Step", self)
        action_step.triggered.connect(self.add_step)
        toolbar.addAction(action_step)

        action_decision = QAction("Add Decision Branch", self) # Changed text
        action_decision.triggered.connect(self.add_decision_branch) # Changed connected method
        toolbar.addAction(action_decision)
        
        action_connect = QAction("Connect Manually", self)
        action_connect.setCheckable(True)
        action_connect.triggered.connect(self.toggle_connection_mode)
        toolbar.addAction(action_connect)

        action_clear = QAction("Clear All", self)
        action_clear.triggered.connect(self.clear_all)
        toolbar.addAction(action_clear)

    def _get_next_position_y(self, previous_item, new_item_height=0):
        """Calculates the Y position for an item directly below the previous_item."""
        if not previous_item:
            return 50  # Default starting Y if no previous item
        
        prev_pos = previous_item.scenePos()
        prev_height = previous_item.boundingRect().height()
        return prev_pos.y() + (prev_height / 2) + VERTICAL_SPACING + (new_item_height / 2)

    def _add_and_connect_item(self, item_class, text, x_pos, y_pos, connect_from_item=None):
        """Helper to create, position, add, and optionally connect an item."""
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
        """Adds a single Step item below the previous one or at top-center."""
        new_item_height = STEP_HEIGHT # Default for StepItem
        initial_x = self.view.mapToScene(self.view.viewport().rect().center()).x()
        initial_y = self._get_next_position_y(self.last_item, new_item_height)

        new_item = self._add_and_connect_item(StepItem, "New Step", initial_x, initial_y, self.last_item)
        self.last_item = new_item
        
    def add_decision_branch(self):
        """Adds a Decision item, two steps below it, and a final step to merge."""
        
        # 1. Place the Decision item
        new_decision_height = DECISION_HEIGHT
        initial_x = self.view.mapToScene(self.view.viewport().rect().center()).x()
        initial_y = self._get_next_position_y(self.last_item, new_decision_height)

        decision_item = self._add_and_connect_item(DecisionItem, "New Decision", initial_x, initial_y, self.last_item)
        
        # Calculate Y for the two parallel steps
        y_for_parallel_steps = self._get_next_position_y(decision_item, STEP_HEIGHT)

        # 2. Place the two Step items below the Decision
        # Calculate X positions for the two steps (left and right of center)
        left_step_x = initial_x - (STEP_WIDTH / 2) - (HORIZONTAL_SPACING / 2)
        right_step_x = initial_x + (STEP_WIDTH / 2) + (HORIZONTAL_SPACING / 2)

        step_left = self._add_and_connect_item(StepItem, "New Step", left_step_x, y_for_parallel_steps, decision_item)
        step_right = self._add_and_connect_item(StepItem, "New Step", right_step_x, y_for_parallel_steps, decision_item)

        # 3. Place the final merging Step item
        # It should be below both parallel steps. Find the lowest point and add spacing.
        lowest_y_of_parallel = max(step_left.scenePos().y() + step_left.boundingRect().height() / 2,
                                   step_right.scenePos().y() + step_right.boundingRect().height() / 2)
        
        final_step_y = lowest_y_of_parallel + VERTICAL_SPACING + (STEP_HEIGHT / 2)
        
        final_step = self._add_and_connect_item(StepItem, "New Step", initial_x, final_step_y)

        # 4. Connect the two parallel steps to the final merging step
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

        self.last_item = final_step # The final step is now the last item for subsequent additions
        self.view.ensureVisible(final_step) # Ensure the whole branch is visible

    def clear_all(self):
        """Clears the entire scene and resets the state."""
        reply = QMessageBox.question(self, "Confirm Clear",
                                     "Are you sure you want to remove everything?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.scene.clear()
            self.last_item = None
            self.start_item = None
            self.connection_mode = False
            self.step_counter = 0

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

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FlowchartApp()
    window.show()
    sys.exit(app.exec())
