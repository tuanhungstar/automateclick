# shared_context.py
import time
import random 
from datetime import datetime
from PyQt6.QtCore import QObject, pyqtSignal


class GuiCommunicator(QObject):
    """
    A QObject subclass used to emit signals from non-GUI threads/objects
    back to the main GUI thread. This is crucial for thread-safe UI updates.
    """
    log_message_signal = pyqtSignal(str)
    update_module_info_signal = pyqtSignal(str)
    update_click_signal = pyqtSignal(str)

class ExecutionContext:
    """
    This class serves as a shared data container that is passed to each bot step method
    during execution. It allows steps to store and retrieve data/results,
    thereby enabling communication and state management across a sequence of steps.
    """
    def __init__(self):
        self.results = {}
        self.data = {}
        self.logs = []
        self.start_time = datetime.now()
        self.gui_communicator = None
        self.click_image_base_dir = ""

    def set_gui_communicator(self, communicator_obj: GuiCommunicator):
        """
        Sets the GuiCommunicator instance that can be used to emit signals
        back to the main GUI thread from within bot step methods.
        """
        self.gui_communicator = communicator_obj

    def get_result(self, key):
        """Retrieves a result stored by a previous step."""
        return self.results.get(key)

    def set_result(self, key, value):
        """Stores a result from the current step."""
        self.results[key] = value

    def get_data(self, key):
        """Retrieves general data stored by a previous step."""
        return self.data.get(key)

    def set_data(self, key, value):
        """Stores general data for subsequent steps."""
        self.data[key] = value
# --- ADD THESE TWO NEW METHODS ---
    def set_click_image_base_dir(self, path: str):
        """Sets the base path to the Click_image folder."""
        self.click_image_base_dir = path

    def get_click_image_base_dir(self) -> str:
        """Gets the base path to the Click_image folder."""
        return self.click_image_base_dir
    # --- END OF ADDITIONS ---
    
    def send_click_status(self,click_status) -> str:
        """send click image status to Label_Info1"""
        
        if self.gui_communicator:
            self.gui_communicator.update_click_signal.emit(click_status)        
        

    # --- END OF ADDITIONS ---
    
    def add_log(self, message):
        """
        Adds a log entry to the context and optionally emits a signal
        to the GUI for real-time log display.
        """
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"{timestamp} - {message}"
        self.logs.append(log_entry)
        
        if self.gui_communicator:
            self.gui_communicator.log_message_signal.emit(log_entry)
            # MODIFIED: Send the full log_entry directly to update_module_info_signal
            # This allows MainWindow to just display it without adding redundant prefixes.
            self.gui_communicator.update_module_info_signal.emit(message)
    
    def __str__(self):
        """
        Returns a string representation of the context, useful for debugging
        and displaying in the GUI (e.g., in a message box).
        """
        return f"Context(results={self.results}, data={self.data}, logs_count={len(self.logs)})"
