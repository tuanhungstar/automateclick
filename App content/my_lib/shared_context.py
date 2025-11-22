# shared_context.py
import time
import random 
from datetime import datetime
from PyQt6.QtCore import QObject, pyqtSignal
from typing import Any, Dict


class GuiCommunicator(QObject):
    """
    A QObject subclass used to emit signals from non-GUI threads/objects
    back to the main GUI thread. This is crucial for thread-safe UI updates.
    """
    log_message_signal = pyqtSignal(str)
    update_module_info_signal = pyqtSignal(str)
    update_click_signal = pyqtSignal(str)
    
    # --- ADD THESE TWO NEW SIGNALS ---
    hide_gui_signal = pyqtSignal()
    show_gui_signal = pyqtSignal()
    # -----------------------------------


class ExecutionContext:
    """
    This class serves as a shared data container that is passed to each bot step method
    during execution. It allows steps to store and retrieve data/results,
    and provides a bridge to the GUI and global variables.
    """
    def __init__(self):
        # --- Existing properties ---
        self.results = {}
        self.data = {}
        self.logs = []
        self.start_time = datetime.now()
        self.gui_communicator = None
        self.click_image_base_dir = ""
        
        # --- NEW: Reference to the main app's global variables ---
        self.global_variables_ref: Dict[str, Any] = {}

    def set_gui_communicator(self, communicator_obj: GuiCommunicator):
        """
        Sets the GuiCommunicator instance that can be used to emit signals
        back to the main GUI thread from within bot step methods.
        """
        self.gui_communicator = communicator_obj

    # --- NEW: Method to link the context to the main app's variables ---
    def set_global_variables_ref(self, variables: Dict[str, Any]):
        """
        Allows the ExecutionWorker to pass a reference to the global variables.
        This should be called once at the start of an execution run.
        """
        self.global_variables_ref = variables
    
    # --- NEW: Method for modules to get a variable at runtime ---
    def get_variable(self, name: str, default: Any = None) -> Any:
        """
        Retrieves a value from the bot's global variables dictionary by name.
        """
        return self.global_variables_ref.get(name, default)

    def set_variable(self, name: str, value: Any):
        """
        Sets or updates a value in the bot's global variables dictionary.
        This is the proper way for modules to assign values back.
        """
        # This directly modifies the dictionary that was passed by reference
        # from the main application's worker thread.
        self.global_variables_ref[name] = value
        
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
        
    def set_click_image_base_dir(self, path: str):
        """Sets the base path to the Click_image folder."""
        self.click_image_base_dir = path

    def get_click_image_base_dir(self) -> str:
        """Gets the base path to the Click_image folder."""
        return self.click_image_base_dir
    
    def send_click_status(self,click_status) -> str:
        """send click image status to Label_Info1"""
        if self.gui_communicator:
            self.gui_communicator.update_click_signal.emit(click_status)        
        
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
            self.gui_communicator.update_module_info_signal.emit(message)

    # --- ADD THESE TWO NEW METHODS ---
    def hide_main_gui(self):
        """Emits a signal to tell the main application to hide its window."""
        if self.gui_communicator:
            self.gui_communicator.hide_gui_signal.emit()

    def show_main_gui(self):
        """Emits a signal to tell the main application to show its window."""
        if self.gui_communicator:
            self.gui_communicator.show_gui_signal.emit()
    # ---------------------------------
            
    def __str__(self):
        """
        Returns a string representation of the context, useful for debugging
        and displaying in the GUI (e.g., in a message box).
        """
        return f"Context(results={self.results}, data={self.data}, logs_count={len(self.logs)})"
