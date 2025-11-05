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

class change_variable_value:
    """A class for basic variable manipulation tasks."""
    def __init__(self, context: Context):
        """Initializes the variable_manipulation class."""
        self.context = context # Store the shared context
        pass
    def assign_value2variable(self,value):
        """Assigns a value to a variable by returning it.

        This method acts as a simple pass-through, which can be useful in
        automation workflows to explicitly set a variable's value.

        Args:
            value (any): The input value of any data type.

        Returns:
            any: The same value that was provided as input.
        """
        return value

    def math_add_variable (self, var1,var2):

        return int(var1) + int(var2)
