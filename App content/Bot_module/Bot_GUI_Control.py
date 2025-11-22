import time
from my_lib.shared_context import ExecutionContext

class Bot_GUI_Control:
    """
    A collection of methods to control the main application's graphical user interface (GUI).
    Allows the bot to hide, show, or check the state of the main window.
    """
    def __init__(self, context: ExecutionContext = None):
        """
        Initializes the Bot GUI Control module.

        Args:
            context (ExecutionContext): The execution context provided by the application,
                                        used to communicate with the main GUI.
        """
        self.context = context

    def hide_main_gui(self):
        """
        Hides the main application window.

        This is useful for running background tasks or when the bot needs to
        perform actions on the screen without the main window being in the way.
        """
        if self.context:
            self.context.hide_main_gui()
            # It's good practice to add a small delay to ensure the window has time to hide
            time.sleep(0.5)
            self.context.add_log("Main GUI has been hidden.")
            return "Main GUI hidden successfully."
        else:
            print("CONTEXT NOT AVAILABLE: Cannot hide GUI.")
            return "Error: ExecutionContext not available."

    def show_main_gui(self):
        """
        Shows the main application window if it was previously hidden.

        This is useful after a background task is complete or when the bot
        needs to bring the main window back into focus.
        """
        if self.context:
            self.context.show_main_gui()
            # It's good practice to add a small delay to ensure the window has time to appear
            time.sleep(0.5)
            self.context.add_log("Main GUI has been shown.")
            return "Main GUI shown successfully."
        else:
            print("CONTEXT NOT AVAILABLE: Cannot show GUI.")
            return "Error: ExecutionContext not available."
