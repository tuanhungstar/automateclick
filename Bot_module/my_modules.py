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
import Gui_Automate
import Mouse_Key_Automate
import File_ReadWrite
import Variable_manipulation


class youmodule2:
    """A generic base class for bot steps."""
    def __init__(self, context: Context):
        self.context = context # Store the shared context
        self.log_prefix = f"[{self.__class__.__name__}]"
        self.context.add_log(f"{self.log_prefix} Initialized with context.")

    def stop_workflow(self):

        raise Exception("User Reqeust stop workflow")
        return 

class yourmodule1:
    """A generic base class for bot steps."""
    def __init__(self, context: Context):
        self.context = context # Store the shared context
        self.log_prefix = f"[{self.__class__.__name__}]"
        self.context.add_log(f"{self.log_prefix} Initialized with context.")

    def stop_workflow(self):
        Mouse_Key_Automate.Keyboard(self.context).moveMouse(x, y)
        raise Exception("User Reqeust stop workflow")
        return 

    def test_add_method(self,x,y):
        Mouse_Key_Automate.Keyboard(self.context).moveMouse(x, y)
        
        return Variable_manipulation.change_variable_value(self.context).math_add_variable(x, y)
