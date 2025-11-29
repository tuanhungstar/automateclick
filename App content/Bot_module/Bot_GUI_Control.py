import time
from my_lib.shared_context import ExecutionContext
import socket
import getpass

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
    def get_computer_name(self):
      """
      Gets the computer's network name (hostname).
      
      Returns:
          str: The hostname of the local machine.
      """
      try:
        hostname = socket.gethostname()
        return hostname
      except Exception as e:
        print(f"An error occurred: {e}")
        return None
        
    def get_login_user(self):
        """
        Gets the current login user ID.
        
        Returns:
            str: The username of the currently logged-in user.
        """
        try:
            # Primary method: using getpass module
            username = getpass.getuser()
            
            if self.context:
                self.context.add_log(f"Current login user: {username}")
            
            return username
        except Exception as e:
            # Fallback method: using environment variables
            try:
                username = os.environ.get('USERNAME') or os.environ.get('USER')
                if username:
                    if self.context:
                        self.context.add_log(f"Current login user (from env): {username}")
                    return username
                else:
                    raise Exception("Could not determine username from environment variables")
            except Exception as fallback_error:
                error_msg = f"Error getting login user: {e}, Fallback error: {fallback_error}"
                if self.context:
                    self.context.add_log(error_msg)
                print(error_msg)
                return None
        
class Bot_Flow_Control:
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
        
# In Bot_GUI_Control.py, inside the Bot_GUI_Control class

    def JumpTo(self, target_step_number: int):
        """
        Instructs the bot runner to immediately jump to a specific step number.

        This step does not execute any other action. It provides a special
        return value that the execution engine intercepts to change the
        flow of the workflow.

        Args:
            target_step_number (int): The 1-based step number to jump to. 
                                      For example, to jump to "Step 5", use the number 5.

        Returns:
            str: A special formatted string that signals a jump instruction.
                 Example: "_JUMP_TO_STEP_::10"
        """
        if not isinstance(target_step_number, int) or target_step_number < 1:
            # You can add a log to the context if you want
            if self.context:
                self.context.add_log(f"ERROR: Invalid target step for JumpTo: '{target_step_number}'. Must be a positive number.")
            # We still raise an error to stop execution, as this is a critical logic failure.
            raise ValueError("Target step number must be a positive integer.")

        # This special string is the "signal" to the main execution worker.
        return f"_JUMP_TO_STEP_::{target_step_number}"
        
    def stop_workflow(self):
        raise Exception("User Reqeust stop workflow")
        return 
    
    def do_nothing(self):
        return 'do nothing'