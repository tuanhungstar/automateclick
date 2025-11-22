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
from dateutil.relativedelta import relativedelta # <-- IMPORT THIS
from calendar import monthrange # To find the last day of a month

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

# --- ENHANCED METHOD ---
    def create_and_run_subprocess_bat(self, python_exe_link: str, py_script_link: str, arg1: str = None, arg2: str = None):
        """
        Creates a .bat file to run a Python script with optional arguments and executes it.

        The .bat file is created in a 'bat_subprocess' subdirectory and is executed in a 
        new console window, allowing the main application to continue running without blocking.

        Args:
            python_exe_link (str): The absolute path to the Python executable 
                                   (e.g., "C:/Python39/python.exe").
            py_script_link (str): The absolute path to the .py script to execute.
            arg1 (str, optional): The first optional argument to pass to the script. Defaults to None.
            arg2 (str, optional): The second optional argument to pass to the script. Defaults to None.
        """
        # 1. Validate the input paths
        if not python_exe_link or not os.path.exists(python_exe_link):
            error_message = f"Invalid Python executable path provided: {python_exe_link}"
            self.context.add_log(error_message)
            return

        if not py_script_link or not os.path.exists(py_script_link):
            error_message = f"Invalid Python script path provided: {py_script_link}"
            self.context.add_log(error_message)
            return

        # 2. Construct the command to be executed
        try:
            # Build the base command with quoted paths for safety
            command = f'"{python_exe_link}" "{py_script_link}"'
            
            # Append arguments if they are provided.
            # Arguments are also quoted to handle spaces correctly.
            if arg1 is not None:
                command += f' "{arg1}"'
            if arg2 is not None:
                command += f' "{arg2}"'

            # 3. Define the .bat file path and content
            # Create the 'bat_subprocess' directory if it doesn't exist
            self.base_directory = os.path.dirname(os.path.abspath(__file__))
            bat_folder = os.path.join(self.base_directory, "bat_subprocess")
            os.makedirs(bat_folder, exist_ok=True)

            # Generate a unique name for the .bat file
            script_name = os.path.splitext(os.path.basename(py_script_link))[0]
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            bat_filename = f"run_{script_name}_{timestamp}.bat"
            bat_filepath = os.path.join(bat_folder, bat_filename)

            # The content of the .bat file.
            bat_content = f'@echo off\n'
            bat_content += f'echo Starting script: {py_script_link}\n'
            bat_content += f'echo With command: {command}\n'
            bat_content += f'{command}\n' # The actual command to run
            bat_content += f'echo Script finished.\n'
            # 'pause' can be useful for debugging to see output before the window closes
            # bat_content += f'pause\n' 

            # 4. Write the .bat file
            with open(bat_filepath, "w", encoding="utf-8") as f:
                f.write(bat_content)
            
            self.context.add_log(f"Successfully created batch file at: {bat_filepath}")
            self.context.add_log(f"Arguments passed: arg1='{arg1}', arg2='{arg2}'")

            # 5. Execute the .bat file using subprocess
            # CREATE_NEW_CONSOLE ensures it opens in a new command prompt window (non-blocking).
            subprocess.Popen([bat_filepath], creationflags=subprocess.CREATE_NEW_CONSOLE)

            self.context.add_log(f"Executing '{bat_filename}' in a new console window...")

        except Exception as e:
            error_message = f"Failed to create or run subprocess: {e}"
            self.context.add_log(error_message)

    def get_left_string(self, input_string: str, num_chars: int) -> str:
        """
        Extracts a specified number of characters from the left side of a string.

        Args:
            input_string (str): The original string.
            num_chars (int): The number of characters to extract from the left.

        Returns:
            str: The extracted substring, or an empty string if an error occurs.
        """
        try:
            # Ensure inputs are the correct type
            if not isinstance(input_string, str):
                input_string = str(input_string)
            num_chars = int(num_chars)

            # Handle negative num_chars as an invalid case
            if num_chars < 0:
                self.context.add_log(f"Error: Number of characters ({num_chars}) cannot be negative.")
                return ""

            return input_string[:num_chars]

        except (ValueError, TypeError) as e:
            self.context.add_log(f"Error in get_left_string: Invalid input provided. Details: {e}")
            return ""

    def get_right_string(self, input_string: str, num_chars: int) -> str:
        """
        Extracts a specified number of characters from the right side of a string.

        Args:
            input_string (str): The original string.
            num_chars (int): The number of characters to extract from the right.

        Returns:
            str: The extracted substring, or an empty string if an error occurs.
        """
        try:
            if not isinstance(input_string, str):
                input_string = str(input_string)
            num_chars = int(num_chars)

            if num_chars < 0:
                self.context.add_log(f"Error: Number of characters ({num_chars}) cannot be negative.")
                return ""
            
            # If num_chars is 0, return an empty string. Otherwise, use negative slicing.
            return input_string[-num_chars:] if num_chars > 0 else ""

        except (ValueError, TypeError) as e:
            self.context.add_log(f"Error in get_right_string: Invalid input provided. Details: {e}")
            return ""

    def get_mid_string(self, input_string: str, start_pos: int, num_chars: int) -> str:
        """
        Extracts a substring from the middle, given a 1-based start position and length.

        Args:
            input_string (str): The original string.
            start_pos (int): The starting position (1-based index) of the substring.
            num_chars (int): The number of characters to extract.

        Returns:
            str: The extracted substring, or an empty string if an error occurs.
        """
        try:
            if not isinstance(input_string, str):
                input_string = str(input_string)
            start_pos = int(start_pos)
            num_chars = int(num_chars)

            if start_pos < 1:
                self.context.add_log(f"Error: Start position ({start_pos}) must be 1 or greater.")
                return ""
            
            if num_chars < 0:
                self.context.add_log(f"Error: Number of characters ({num_chars}) cannot be negative.")
                return ""

            # Convert 1-based start_pos to 0-based index
            start_index = start_pos - 1
            end_index = start_index + num_chars
            
            return input_string[start_index:end_index]

        except (ValueError, TypeError) as e:
            self.context.add_log(f"Error in get_mid_string: Invalid input provided. Details: {e}")
            return ""



    def create_past_months_df(self,number_month, date_format: str = '%m/%d/%Y') -> pd.DataFrame:
        """
        Creates a DataFrame with "Year-Month", "FirstDayOfMonth", and "LastDayOfMonth".
        
        The data consists of number_month rows, covering the current month and the 12 preceding months.

        Args:
            date_format (str, optional): A format string to apply to the date columns
                                         (e.g., "%d-%m-%Y", "%Y/%m/%d"). 
                                         If None, columns will be datetime.date objects.
                                         Defaults to None.

        Returns:
            pd.DataFrame: A DataFrame with the three date-related columns, or an empty
                          DataFrame on error.
        """
        try:
            log_message = "Generating DataFrame for the last number_month months"
            if date_format:
                log_message += f" with date format '{date_format}'."
            else:
                log_message += " with date objects."
            self.context.add_log(log_message)
            
            today = date.today()
            date_records = []
            
            for i in range(number_month):
                target_date = today - relativedelta(months=i)
                first_day = target_date.replace(day=1)
                _, num_days = monthrange(first_day.year, first_day.month)
                last_day = first_day.replace(day=num_days)
                
                # Apply formatting if a format string is provided
                if date_format:
                    first_day_val = first_day.strftime(date_format)
                    last_day_val = last_day.strftime(date_format)
                else:
                    # Otherwise, use the date objects themselves
                    first_day_val = first_day
                    last_day_val = last_day

                date_records.append({
                    "Year-Month": first_day.strftime('%Y-%m'),
                    "FirstDayOfMonth": first_day_val,
                    "LastDayOfMonth": last_day_val,
                    "Year": first_day.strftime('%Y'),
                    "Month": first_day.strftime('%m')
                })

            df = pd.DataFrame(date_records)
            
            self.context.add_log(f"Successfully created DataFrame with {len(df)} entries.")
            return df
            
        except ValueError as e:
            # This will catch invalid format strings
            self.context.add_log(f"Error: Invalid date_format string provided ('{date_format}'). Details: {e}")
            return pd.DataFrame()
        except Exception as e:
            self.context.add_log(f"An error occurred while creating the months DataFrame: {e}")
            return pd.DataFrame()
class mic_help_class:
    """A generic base class for bot steps."""
    def __init__(self, context: Context):
        self.context = context # Store the shared context
        self.log_prefix = f"[{self.__class__.__name__}]"
        self.context.add_log(f"{self.log_prefix} Initialized with context.")
    def filter_df_rule(self,df):
        df_result=df
        df_result=df_result[(df_result['Status'] == 'IN_USE')|(df_result['Status'] == 'READY')].copy()   
        df_result=df_result[df_result['Type'] == 'DIRECT'].copy()  
        df_result['Valid From'] = pd.to_datetime(df_result['Valid From'] )
        df_result['Valid From'] =  df_result['Valid From'].dt.date
        
        return df_result.reset_index(drop=True)