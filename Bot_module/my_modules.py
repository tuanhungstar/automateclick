# Bot_module/my_modules.py
import sys
import json
import win32com.client
import win32gui
import os
import clipboard
from datetime import datetime, timedelta, date
import pandas as pd
import numpy as np
import time
import io
import base64
from PIL import ImageGrab
import PIL.Image
import sys
import json
import subprocess
import pyautogui
import re
from PIL.ImageQt import ImageQt
from pymsgbox import *
import pygetwindow as gw
import keyboard
import openpyxl
import datetime
from datetime import datetime
from my_lib.shared_context import ExecutionContext as Context
#import Gui_Automate
#import Mouse_Key_Automate
#import File_ReadWrite
#import Variable_manipulation
import clipboard

class youmodule2:
    """A generic base class for bot steps."""
    def __init__(self, context: Context):
        self.context = context # Store the shared context
        self.log_prefix = f"[{self.__class__.__name__}]"
        self.context.add_log(f"{self.log_prefix} Initialized with context.")

    def stop_workflow(self):

        raise Exception("User Reqeust stop workflow")
        return 
    def do_nothing(self):

        return 'do nothing'

    def clipboard_value(self):
        ''' Return the value in clipboard '''
        return clipboard.paste()

    def blocking_reason(self, similarity):

        '''
            block_reason = f"Block due to similarity check{str(similarity)} on {current_date}"

        '''
        import datetime
        format_string="%d-%m-%Y"

        current_date = datetime.date.today().strftime(format_string)
        similarity=str(int(similarity)*100 )
        block_reason = f"Block due to similarity check {similarity}%  on {current_date}"
        return block_reason
    def delete_file(self,file_link):

        try:
            os.remove(file_link)
            return f'Deleted file {file_link}'
        except FileNotFoundError:
            self.context.add_log(f"Error: File not found at {file_link}")
            
        except OSError as e:
            self.context.add_log(f"Error deleting file {file_link}: {e}") 
        return f'Not able to deleted file {file_link}'
               
               
    def wait_for_file_os(self,file_link: str, timeout: int = 30) -> bool:
        """
        Waits for a file to exist using the os module.
        
        Args:
            file_link (str): The string path of the file to wait for.
            timeout (int): Maximum time to wait in seconds.

        Returns:
            bool: True if the file was found, False on timeout.
        """
        poll_interval: float = 0.5
        start_time = time.monotonic()
        
        while time.monotonic() - start_time < timeout:
            if os.path.exists(file_link) and os.path.isfile(file_link):
                self.context.add_log(f"File found: {file_link}")

                return True
            time.sleep(poll_interval)

        self.context.add_log(f"Timeout reached. File not found: {file_link}")
        return False            

# In the MainWindow class, add or replace this method:

    def create_and_run_subprocess_bat(self, python_exe_link: str, py_script_link: str):
        """
        Creates a .bat file to run the given script with the specified Python interpreter
        and executes it using subprocess. The paths are provided as parameters.

        Args:
            python_exe_link (str): The absolute path to the Python executable (e.g., "C:/Python39/python.exe").
            py_script_link (str): The absolute path to the .py script to execute.
        """
        # 1. Validate the input paths
        if not python_exe_link or not os.path.exists(python_exe_link):
            error_message = f"Invalid Python executable path provided: {python_exe_link}"
            self.context.add_log(error_message)
            #QMessageBox.warning(self, "Path Error", error_message)
            return

        if not py_script_link or not os.path.exists(py_script_link):
            error_message = f"Invalid Python script path provided: {py_script_link}"
            self.context.add_log(error_message)
            #QMessageBox.warning(self, "Path Error", error_message)
            return

        # 2. Define the .bat file path and content
        try:
            # Create the 'bat_subprocess' directory if it doesn't exist
            self.base_directory = os.path.dirname(os.path.abspath(__file__))
            bat_folder = os.path.join(self.base_directory, "bat_subprocess")
            os.makedirs(bat_folder, exist_ok=True)

            # Generate a unique name for the .bat file to avoid conflicts
            script_name = os.path.splitext(os.path.basename(py_script_link))[0]
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            bat_filename = f"run_{script_name}_{timestamp}.bat"
            bat_filepath = os.path.join(bat_folder, bat_filename)

            # The content of the .bat file. Quotes are crucial for paths with spaces.
            # @echo off prevents commands from being printed to the console.
            # The "start" command with the /B flag runs the process in the same window,
            # but without this flag, it opens a new one, which is what we want.
            bat_content = f'@echo off\n'
            bat_content += f'echo Starting script: {py_script_link}\n'
            bat_content += f'"{python_exe_link}" "{py_script_link}"\n'
            bat_content += f'echo Script finished.\n'

            # 3. Write the .bat file
            with open(bat_filepath, "w", encoding="utf-8") as f:
                f.write(bat_content)
            
            self.context.add_log(f"Successfully created batch file at: {bat_filepath}")

            # 4. Execute the .bat file using subprocess
            # Using CREATE_NEW_CONSOLE ensures it opens in its own, new command prompt window.
            # This is non-blocking, so your main application will continue to run.
            subprocess.Popen([bat_filepath], creationflags=subprocess.CREATE_NEW_CONSOLE)

            self.context.add_log(f"Executing '{bat_filename}' in a new console window...")

        except Exception as e:
            error_message = f"Failed to create or run subprocess: {e}"
            self.context.add_log(error_message)
            #QMessageBox.critical(self, "Subprocess Error", error_message)
