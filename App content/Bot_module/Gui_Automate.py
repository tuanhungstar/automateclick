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
import pyperclipimg
from typing import Optional, List, Dict, Any

# --- PyQt6 Imports ---
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QDialogButtonBox,
    QComboBox, QWidget, QGroupBox, QMessageBox, QLabel, QHBoxLayout, QRadioButton,
    QCheckBox, QDoubleSpinBox
)
from PyQt6.QtCore import Qt
    
from my_lib.shared_context import ExecutionContext as Context

#
# --- HELPER: The GUI Dialog for Bot Utility ---
#
class _BotUtilityDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None,
                 initial_variable: Optional[str] = None):
        super().__init__(parent)
        self.setWindowTitle("Bot Utility Configuration")
        self.setMinimumWidth(500)
        self.global_variables = global_variables

        main_layout = QVBoxLayout(self)

        # 1. Action Selection
        action_group = QGroupBox("Action Selection")
        action_layout = QFormLayout(action_group)
        self.method_combo = QComboBox()
        self.method_combo.addItems([
            "advance_action (AI Detection Action)",
            "left_click (Image Recognition)",
            "right_click (Image Recognition)",
            "double_click (Image Recognition)",
            "activate_window",
            "maximize_window",
            "close_window"
        ])
        self.method_combo.currentTextChanged.connect(self._on_method_changed)
        action_layout.addRow("Method:", self.method_combo)
        main_layout.addWidget(action_group)

        # 2. Advance Action Configuration
        self.advance_group = QGroupBox("Advance Action Settings")
        advance_layout = QFormLayout(self.advance_group)
        
        self.ai_var_combo = QComboBox()
        self.ai_var_combo.addItems(["-- Select Variable --"] + [str(v) for v in self.global_variables])
        
        self.label_input = QLineEdit()
        self.label_input.setPlaceholderText("e.g., Liên hệ, Button Name")
        
        self.action_type_combo = QComboBox()
        self.action_type_combo.addItems(["left click", "right click", "double click", "move mouse", "check if button appear", "check if button disappear"])
        
        advance_layout.addRow("AI Response (Variable):", self.ai_var_combo)
        advance_layout.addRow("Target Label:", self.label_input)
        advance_layout.addRow("Action Type:", self.action_type_combo)
        
        self.click_all_check = QCheckBox("Click All Matching Elements")
        self.click_delay_spin = QDoubleSpinBox()
        self.click_delay_spin.setRange(0.0, 60.0)
        self.click_delay_spin.setValue(0.5)
        self.click_delay_spin.setSuffix(" sec")
        
        advance_layout.addRow(self.click_all_check)
        advance_layout.addRow("Delay between clicks:", self.click_delay_spin)
        
        main_layout.addWidget(self.advance_group)

        # 3. Standard Image Action Configuration
        self.image_group = QGroupBox("Image Action Settings")
        image_layout = QFormLayout(self.image_group)
        self.image_name_input = QLineEdit()
        self.image_name_input.setPlaceholderText("Image name in Click_image folder")
        image_layout.addRow("Image Name:", self.image_name_input)
        main_layout.addWidget(self.image_group)

        # 4. Window Action Configuration
        self.window_group = QGroupBox("Window Action Settings")
        window_layout = QFormLayout(self.window_group)
        self.window_title_input = QLineEdit()
        window_layout.addRow("Window Title:", self.window_title_input)
        main_layout.addWidget(self.window_group)

        # Buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # Connections
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        self._on_method_changed(self.method_combo.currentText())
        if initial_config: self._populate_from_initial_config(initial_config)

    def _on_method_changed(self, method_name):
        self.advance_group.setVisible("advance_action" in method_name)
        self.image_group.setVisible("_click" in method_name)
        self.window_group.setVisible("_window" in method_name or "advance_action" in method_name)

    def _populate_from_initial_config(self, config):
        method = config.get("method")
        if method: self.method_combo.setCurrentText(method)
        
        if "advance_action" in method:
            self.ai_var_combo.setCurrentText(config.get("ai_response_var", "-- Select Variable --"))
            self.label_input.setText(config.get("target_label", ""))
            self.action_type_combo.setCurrentText(config.get("action_type", "left click"))
            self.click_all_check.setChecked(config.get("click_all", False))
            self.click_delay_spin.setValue(config.get("delay", 0.5))
        elif "_click" in method:
            self.image_name_input.setText(config.get("image_to_click", ""))
        elif "_window" in method:
            self.window_title_input.setText(config.get("title", ""))

    def get_executor_method_name(self) -> str:
        method_text = self.method_combo.currentText()
        return method_text.split(" ")[0]

    def get_config_data(self) -> Optional[Dict[str, Any]]:
        method = self.get_executor_method_name()
        config = {"method": self.method_combo.currentText()}
        
        if method == "advance_action":
            ai_var = self.ai_var_combo.currentText()
            if ai_var == "-- Select Variable --":
                QMessageBox.warning(self, "Input Error", "Please select an AI response variable.")
                return None
            config.update({
                "ai_response": ai_var,
                "target_label": self.label_input.text().strip(),
                "action_type": self.action_type_combo.currentText(),
                "window_title": self.window_title_input.text().strip(),
                "click_all": self.click_all_check.isChecked(),
                "delay": self.click_delay_spin.value()
            })
        elif "_click" in method:
            config["image_to_click"] = self.image_name_input.text().strip()
        elif "_window" in method:
            config["title"] = self.window_title_input.text().strip()
            
        return config

    def get_assignment_variable(self) -> Optional[str]:
        return None # Usually these actions return True/False or None

class Bot_utility:
    '''
    A utility class containing a collection of reusable functions for GUI automation,
    window management, and file operations.
    '''
    def __init__(self, context: Context):
        """Initializes the Bot_utility class.

        Args:
            context (Context): A shared context object for logging and state management
                               across different bot components.
        """
        self.context = context # Store the shared context
        self.log_prefix = f"[{self.__class__.__name__}]"
        self.context.add_log(f"{self.log_prefix} Initialized with context.")
        
        self.click_image_folder_path = context.get_click_image_base_dir()
        pass

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], 
                           initial_config: Optional[Dict[str, Any]] = None, 
                           initial_variable: Optional[str] = None) -> QDialog:
        """
        Configures the Bot Utility actions.
        """
        return _BotUtilityDialog(
            global_variables=global_variables, 
            parent=parent_window,
            initial_config=initial_config,
            initial_variable=initial_variable
        )
        
    def advance_action(self, ai_response, target_label, action_type="left click", window_title=None, click_all=False, delay=0.5):
        """
        Advanced action based on AI detection results.
        """
        self.context.add_log(f"{self.log_prefix} advance_action called. Label: '{target_label}', Action: '{action_type}', Click All: {click_all}")
        try:
            # 0. Resolve variable if name provided
            if isinstance(ai_response, str):
                self.context.add_log(f"{self.log_prefix} Attempting to resolve variable: {ai_response}")
                resolved = self.context.get_variable(ai_response)
                if resolved is not None:
                    ai_response = resolved
                    self.context.add_log(f"{self.log_prefix} Variable '{ai_response}' resolved successfully.")
                else:
                    self.context.add_log(f"{self.log_prefix} Variable '{ai_response}' not found or empty.")

            # 1. Parse JSON from various input types
            data = None
            if isinstance(ai_response, pd.DataFrame):
                self.context.add_log(f"{self.log_prefix} Processing DataFrame input.")
                if not ai_response.empty:
                    val = ai_response.iloc[0, 0]
                    if isinstance(val, str):
                        try: 
                            data = json.loads(val)
                            self.context.add_log(f"{self.log_prefix} Parsed JSON from first cell.")
                        except: data = val
                    else: data = val
            elif isinstance(ai_response, str):
                try: 
                    data = json.loads(ai_response)
                    self.context.add_log(f"{self.log_prefix} Parsed JSON from string.")
                except: data = ai_response
            elif isinstance(ai_response, dict):
                data = ai_response
                self.context.add_log(f"{self.log_prefix} Using dictionary input.")

            if data is None:
                self.context.add_log(f"{self.log_prefix} ERROR: Could not parse AI response (Data is None).")
                return False

            if isinstance(data, str):
                try: 
                    data = json.loads(data)
                    self.context.add_log(f"{self.log_prefix} Parsed nested JSON string.")
                except: pass

            def find_key_recursive(obj, target_key):
                if isinstance(obj, dict):
                    if target_key in obj: return obj[target_key]
                    for v in obj.values():
                        res = find_key_recursive(v, target_key)
                        if res: return res
                elif isinstance(obj, list):
                    for item in obj:
                        res = find_key_recursive(item, target_key)
                        if res: return res
                return None

            elements = find_key_recursive(data, "results") or find_key_recursive(data, "elements") or find_key_recursive(data, "objects")
            
            if not elements or not isinstance(elements, list):
                if isinstance(data, list):
                    elements = data
                else:
                    self.context.add_log(f"{self.log_prefix} ERROR: No elements found in AI response. Data keys: {list(data.keys()) if isinstance(data, dict) else 'Not a dict'}")
                    return False
            
            self.context.add_log(f"{self.log_prefix} Searching for label '{target_label}' in {len(elements)} elements.")
            # 2. Find target label(s)
            matching_elements = []
            for el in elements:
                label = el.get("label") or el.get("keyword") or ""
                # If target_label is empty, match ALL elements found
                if not target_label or str(label).lower().strip() == str(target_label).lower().strip():
                    matching_elements.append(el)
            
            if not click_all and matching_elements:
                matching_elements = [matching_elements[0]]
                
            # 3. Handle 'check' actions
            if action_type == "check if button appear":
                res = len(matching_elements) > 0
                self.context.add_log(f"{self.log_prefix} Check appear '{target_label}': {res}")
                return res
            if action_type == "check if button disappear":
                res = len(matching_elements) == 0
                self.context.add_log(f"{self.log_prefix} Check disappear '{target_label}': {res}")
                return res
            
            # 4. Handle interaction actions
            if not matching_elements:
                self.context.add_log(f"{self.log_prefix} Label '{target_label}' not found.")
                return False
            
            self.context.add_log(f"{self.log_prefix} Found {len(matching_elements)} matches for '{target_label}'. Click All={click_all}")
            
            success_count = 0
            for i, target_el in enumerate(matching_elements):
                if i > 0 and delay > 0:
                    self.context.add_log(f"{self.log_prefix} Waiting {delay}s before next click...")
                    time.sleep(delay)
                
                pix = target_el.get("bbox_pixels")
                if not pix:
                    bbox = target_el.get("bbox") or target_el.get("box")
                    if isinstance(bbox, list) and len(bbox) == 4:
                        pix = {"xmin": bbox[0], "ymin": bbox[1], "xmax": bbox[2], "ymax": bbox[3]}
                
                if not pix:
                    self.context.add_log(f"{self.log_prefix} No coordinates found for element {i+1}.")
                    continue
                
                # Calculate center in image coordinates
                img_x = (pix["xmin"] + pix["xmax"]) / 2
                img_y = (pix["ymin"] + pix["ymax"]) / 2
                
                # Calculate screen coordinates using window offset
                offset_x, offset_y = 0, 0
                if window_title:
                    wins = gw.getWindowsWithTitle(window_title)
                    if wins:
                        win = wins[0]
                        offset_x, offset_y = win.left, win.top
                        if win.isMaximized:
                            offset_x += 8
                            offset_y += 8
                    else:
                        self.context.add_log(f"{self.log_prefix} WARNING: Window '{window_title}' not found.")
                
                screen_x = img_x + offset_x
                screen_y = img_y + offset_y
                
                act = str(action_type).lower().strip()
                try:
                    if act == "move mouse":
                        pyautogui.moveTo(screen_x, screen_y, duration=0.2)
                    elif act == "left click":
                        pyautogui.click(screen_x, screen_y)
                    elif act == "right click":
                        pyautogui.rightClick(screen_x, screen_y)
                    elif act == "double click":
                        pyautogui.doubleClick(screen_x, screen_y)
                    else:
                        continue
                    
                    success_count += 1
                except Exception as click_err:
                    self.context.add_log(f"{self.log_prefix} Interaction {i+1} failed: {click_err}")

            return success_count > 0
        except Exception as e:
            self.context.add_log(f"{self.log_prefix} Error in advance_action: {e}")
            return False
        
    def click_ai_element(self, ai_response_json, label_to_click):
        """Clicks an element identified by AI coordinates in a JSON response.
        
        Args:
            ai_response_json (str or dict): The JSON response from the AI server.
            label_to_click (str): The label of the element to click.
        """
        try:
            data = None
            # 1. Handle DataFrame (common output from bot modules)
            if isinstance(ai_response_json, pd.DataFrame):
                if not ai_response_json.empty:
                    val = ai_response_json.iloc[0, 0]
                    if isinstance(val, str):
                        try: data = json.loads(val)
                        except: data = val # Might be raw string
                    else: data = val
            
            # 2. Handle string
            elif isinstance(ai_response_json, str):
                try: data = json.loads(ai_response_json)
                except: pass
            
            # 3. Handle dict
            elif isinstance(ai_response_json, dict):
                data = ai_response_json

            if data is None:
                self.context.add_log(f"{self.log_prefix} Could not parse AI response (type: {type(ai_response_json)})")
                return False

            # Sometimes the response is a string inside a dict, or nested
            if isinstance(data, str):
                try: data = json.loads(data)
                except: pass

            def find_key_recursive(obj, target_key):
                if isinstance(obj, dict):
                    if target_key in obj: return obj[target_key]
                    for v in obj.values():
                        res = find_key_recursive(v, target_key)
                        if res: return res
                elif isinstance(obj, list):
                    for item in obj:
                        res = find_key_recursive(item, target_key)
                        if res: return res
                return None

            elements = find_key_recursive(data, "elements") or find_key_recursive(data, "objects")
            
            if not elements or not isinstance(elements, list):
                self.context.add_log(f"{self.log_prefix} No 'elements' or 'objects' list found in AI response: {data}")
                return False
                
            # 1:1 Mapping (No scaling as user specified 100% scale)
            scale_x, scale_y = 1.0, 1.0
            
            for el in elements:
                if el.get("label", "").lower() == label_to_click.lower():
                    # Check for bbox_pixels first
                    pix = el.get("bbox_pixels")
                    if pix:
                        xmin = pix.get("xmin", 0)
                        ymin = pix.get("ymin", 0)
                        xmax = pix.get("xmax", xmin)
                        ymax = pix.get("ymax", ymin)
                        click_x = (xmin + xmax) / 2
                        click_y = (ymin + ymax) / 2
                    else:
                        # Fallback to direct bbox or coordinator
                        bbox = el.get("bbox") or el.get("box") or el.get("coordinator")
                        if isinstance(bbox, list):
                            if len(bbox) == 4:
                                click_x = (bbox[1] + bbox[3]) / 2 # Assuming [ymin, xmin, ymax, xmax] standard
                                click_y = (bbox[0] + bbox[2]) / 2
                            elif len(bbox) == 2:
                                click_x, click_y = bbox[0], bbox[1]
                            else: continue
                        else: continue
                    
                    self.context.add_log(f"{self.log_prefix} Clicking AI element '{label_to_click}' at ({click_x:.1f}, {click_y:.1f}) [1:1 Scale]")
                    pyautogui.click(click_x, click_y)
                    return True
            
            self.context.add_log(f"{self.log_prefix} Label '{label_to_click}' not found in AI response.")
            return False
        except Exception as e:
            self.context.add_log(f"{self.log_prefix} Error clicking AI element: {str(e)}")
            return False
        
    def _base64_pgn(self,text):
        """Decodes a base64 encoded string into a PIL Image object.

        This is a private helper method used to convert base64 image data from
        text files into a format usable by image recognition libraries.

        Args:
            text (str): A base64 encoded string representing an image.

        Returns:
            PIL.Image.Image: An image object that can be processed by libraries
                             like pyautogui.
        """
        image_data = base64.b64decode(text)
        img_file = PIL.Image.open(io.BytesIO(image_data))    
        return img_file

    def text_located(self,im):
        """Locates an image on the screen and moves the mouse cursor to it.

        Args:
            im (str): A base64 encoded string of the image to find on the screen.

        Returns:
            pyautogui.Point or None: The coordinates (x, y) of the center of the
                                     found image as a Point object, or None if the
                                     image is not found.
        """
        img_file = self._base64_pgn(im)
        location =None
        try:
            location= pyautogui.locateCenterOnScreen(img_file,grayscale=True, confidence=0.92)
        except:
            pass
            
        if location!=None:
            pyautogui.moveTo(location.x, location.y)
        return location
        
    def _check_item_existing(self,image_to_click):
        """Checks if a specific UI element (as an image) exists on the screen.

        This private helper method reads a base64 encoded image from a specified
        file and searches for it on the screen.

        Args:
            image_to_click (str): The name of the file (without extension) located in the
                                'Click_image/' directory, which contains the image data.

        Returns:
            bool: True if the image is found on the screen, False otherwise.
        """
        full_path_without_ext = os.path.join(self.click_image_folder_path, image_to_click)
        full_path_with_ext = full_path_without_ext + ".txt"
        with open(full_path_with_ext) as json_file:
            image_file = json.load(json_file)
        for key, data in image_file.items():
            img_obj = self._base64_pgn(data)
            location=None
            try:
                location= pyautogui.locateCenterOnScreen(img_obj,grayscale=True, confidence=0.98)
            except:
                pass
            if location!=None:
                return True
        return False
        
    def check_picture_list(self, picture_list_str: str) -> bool:
        """
        Checks if any picture in a comma-separated list of picture names exists on the screen.

        Args:
            picture_list_str (str): A comma-separated string of picture names (without extensions).
                                    These names should correspond to files in the 'Click_image/' directory.

        Returns:
            bool: True if any picture from the list is found on the screen, False otherwise.
        """
        picture_names = [name.strip() for name in picture_list_str.split(',')]
        
        for picture_name in picture_names:
            if self._check_item_existing(picture_name):
                self.context.add_log(f"Found '{picture_name}' from the picture list.")
                return True  # Break and return True as soon as one picture is found
        
        self.context.add_log(f"None of the pictures in the list '{picture_list_str}' were found.")
        return False
        
    def left_click(self,image_to_click,offset_x=0,offset_y=0,confidence=0.92,stop_if_not_found=False):
        """Finds a UI element (as an image) on the screen and performs a left click.

        Args:
            image_to_click (str): The name of the image file (without extension) to find and click.
            offset_x (int, optional): The horizontal offset in pixels from the image's center
                                      to click. Defaults to 0.
            offset_y (int, optional): The vertical offset in pixels from the image's center
                                      to click. Defaults to 0.
            confidence (float, optional): The confidence level for the image recognition.
                                          Defaults to 0.92.

        Returns:
            str: A status message, 'left_click_done' on success or a failure message
                 if the image could not be found.
        """
        location=None
        file_name= image_to_click
        full_path_without_ext = os.path.join(self.click_image_folder_path, image_to_click)
        full_path_with_ext = full_path_without_ext + ".txt"
        confidence = float(confidence)
        with open(full_path_with_ext) as json_file:
            
            image_file = json.load(json_file)
        for key, data in image_file.items():
            img_obj = self._base64_pgn(data)
            try:
                location= pyautogui.locateCenterOnScreen(img_obj,grayscale=True, confidence=confidence)
                if location!=None:
                    pyautogui.click(location.x + int(offset_x), location.y + int(offset_y))
                    self.context.add_log(f"Image found: {file_name}")
                    self.context.send_click_status(f"Image found: {file_name}")
                    return 'left_click_done'
            except:
                pass
        
        if location is None:
            self.context.send_click_status(f"Image not found: {file_name}")
            self.context.add_log(f"Image not found: {file_name}")

        if stop_if_not_found:
            raise Exception ("Image not found")
        return 'left_click_done fail: {}'.format(file_name)
        
    def right_click(self,image_to_click,offset_x=0,offset_y=0,confidence=0.92):
        """Finds a UI element (as an image) on the screen and performs a right click.

        Args:
            image_to_click (str): The name of the image file (without extension) to find and click.
            offset_x (int, optional): The horizontal offset in pixels from the image's center
                                      to click. Defaults to 0.
            offset_y (int, optional): The vertical offset in pixels from the image's center
                                      to click. Defaults to 0.
            confidence (float, optional): The confidence level for the image recognition.
                                          Defaults to 0.92.

        Returns:
            str: A status message, 'left_click_done' on success or a failure message
                 if the image could not be found.
        """
        location=None
        file_name= image_to_click
        full_path_without_ext = os.path.join(self.click_image_folder_path, image_to_click)
        full_path_with_ext = full_path_without_ext + ".txt"

        with open(full_path_with_ext) as json_file:
            
            image_file = json.load(json_file)
        for key, data in image_file.items():
            img_obj = self._base64_pgn(data)
            try:
                location= pyautogui.locateCenterOnScreen(img_obj,grayscale=True, confidence=confidence)
                if location!=None:
                    pyautogui.rightClick(location.x + int(offset_x), location.y + int(offset_y))
                    self.context.add_log(f"Image found: {file_name}")
                    return 'right_click_done'
            except:
                pass
        if location is None:
            self.context.add_log(f"Image not found: {file_name}")
        return 'right_click_done fail: {}'.format(file_name)
        
    def double_click(self,image_to_click,offset_x=0,offset_y=0,confidence=0.92):
        """Finds a UI element (as an image) on the screen and performs a right click.

        Args:
            image_to_click (str): The name of the image file (without extension) to find and click.
            offset_x (int, optional): The horizontal offset in pixels from the image's center
                                      to click. Defaults to 0.
            offset_y (int, optional): The vertical offset in pixels from the image's center
                                      to click. Defaults to 0.
            confidence (float, optional): The confidence level for the image recognition.
                                          Defaults to 0.92.

        Returns:
            str: A status message, 'left_click_done' on success or a failure message
                 if the image could not be found.
        """
        location=None
        file_name= image_to_click
        full_path_without_ext = os.path.join(self.click_image_folder_path, image_to_click)
        full_path_with_ext = full_path_without_ext + ".txt"

        with open(full_path_with_ext) as json_file:
            
            image_file = json.load(json_file)
        for key, data in image_file.items():
            img_obj = self._base64_pgn(data)
            try:
                location= pyautogui.locateCenterOnScreen(img_obj,grayscale=True, confidence=confidence)
                if location!=None:
                    pyautogui.doubleClick(location.x + int(offset_x), location.y + int(offset_y))
                    self.context.add_log(f"Image found: {file_name}")
                    return 'double_click_done'
            except:
                pass
        if location is None:
            self.context.add_log(f"Image not found: {file_name}")
        return 'double_click fail: {}'.format(file_name)      
    def get_image_location(self,image_to_click,offset_x=0,offset_y=0,confidence=0.92):
        """ Find and return image location
        """
        location=None
        file_name= image_to_click
        full_path_without_ext = os.path.join(self.click_image_folder_path, image_to_click)
        full_path_with_ext = full_path_without_ext + ".txt"

        with open(full_path_with_ext) as json_file:
            
            image_file = json.load(json_file)
        for key, data in image_file.items():
            img_obj = self._base64_pgn(data)
            try:
                location= pyautogui.locateCenterOnScreen(img_obj,grayscale=True, confidence=confidence)
                if location!=None:
                    self.context.add_log(f"Image found: {file_name}")
                    self.context.send_click_status(f"Image found: {file_name}")
                    return location
            except:
                pass
        if location is None:
            self.context.send_click_status(f"Image not found: {file_name}")
            self.context.add_log(f"Image not found: {file_name}")
        return None

    def image_action_advanced(self, window_title, image_to_click, action_type="Left Click", waiting_time=0.5, timeout=10, exit_if_found=True):
        """
        Unified Advanced Action:
        1. Activate window (if title provided)
        2. Wait 500ms stabilization
        3. Perform specified action (Click, Wait, etc.)
        """
        if window_title:
            self.activate_window(window_title)
            time.sleep(0.5)
            
        all_image_files = [img.strip() for img in image_to_click.split(',') if img.strip()]
        found_any = False
        
        for img_name in all_image_files:
            res = 'fail'
            if action_type == "Left Click":
                res = self.left_click(img_name, confidence=0.92)
            elif action_type == "Right Click":
                res = self.right_click(img_name, confidence=0.92)
            elif action_type == "Double Click":
                res = self.double_click(img_name, confidence=0.92)
            elif action_type == "Wait Appear":
                exists = self.check_image_exits(img_name, timeout=timeout)
                res = 'done' if exists else 'fail'
            elif action_type == "Wait Disappear":
                disappeared = self.wait_image_disappear(img_name, timeout=timeout)
                res = 'done' if disappeared else 'fail'
            
            # Check for various success strings
            if res in ['left_click_done', 'right_click_done', 'double_click_done', 'done']:
                found_any = True
                if exit_if_found:
                    return 'done'
                if action_type in ["Left Click", "Right Click", "Double Click"]:
                    time.sleep(float(waiting_time))
        
        return 'done' if found_any else 'fail'
        
    def left_click_cross_2_images(self,image_to_click_x,image_to_click_y,offset_x=0,offset_y=0,confidence=0.92):
        """ Find and return image location
        """
        location=None
        x_position=None
        file_name= image_to_click_x
        full_path_without_ext = os.path.join(self.click_image_folder_path, image_to_click_x)
        full_path_with_ext = full_path_without_ext + ".txt"

        with open(full_path_with_ext) as json_file:
            
            image_file = json.load(json_file)
        for key, data in image_file.items():
            img_obj = self._base64_pgn(data)
            try:
                location= pyautogui.locateCenterOnScreen(img_obj,grayscale=True, confidence=confidence)
                if location!=None:
                    self.context.add_log(f"Image found: {file_name}")
                    self.context.send_click_status(f"Image found: {file_name}")
                    x_position = location.x
                    break
            except:
                pass
        if x_position is None:
            self.context.send_click_status(f"Image not found: {file_name}")
            self.context.add_log(f"Image not found: {file_name}")
            return None      



        location=None
        y_position=None
        file_name= image_to_click_y
        full_path_without_ext = os.path.join(self.click_image_folder_path, image_to_click_y)
        full_path_with_ext = full_path_without_ext + ".txt"

        with open(full_path_with_ext) as json_file:
            
            image_file = json.load(json_file)
        for key, data in image_file.items():
            img_obj = self._base64_pgn(data)
            try:
                location= pyautogui.locateCenterOnScreen(img_obj,grayscale=True, confidence=confidence)
                if location!=None:
                    self.context.add_log(f"Image found: {file_name}")
                    self.context.send_click_status(f"Image found: {file_name}")
                    y_position = location.y
                    break
            except:
                pass
        if y_position is None:
            self.context.send_click_status(f"Image not found: {file_name}")
            self.context.add_log(f"Image not found: {file_name}")
            return None      

        if x_position is not None and y_position is not None:
        
            pyautogui.click(x_position + int(offset_x),y_position+int(offset_y))
            self.context.add_log(f"left_click_cross_2_images :{image_to_click_x} and {image_to_click_y} is done")
            return "done"


        

    def activate_window(self,title):
        """Brings a window with a matching title to the foreground.

        The method repeatedly tries to find and activate the window until it succeeds
        or a timeout is reached.

        Args:
            title (str): A partial or full title of the window to activate.

        Returns:
            str: 'done' if the window was successfully activated, or 'fail' if the
                 window could not be found after ~1000 attempts.
        """
        loop=0
        full_tile =''
        while title not in full_tile:
            try:
                need_avtive = gw.getWindowsWithTitle(title)[0]
                #need_avtive.maximize()
                need_avtive.activate()
                full_tile = gw.getActiveWindow().title
                if title in full_tile:
                    #print(full_tile)
                    return 'done'
            except:
                #print ('try times:' , str(loop))
                loop +=1
                if loop >1000:
                    return 'fail'

    def maximize_window(self,title):
        """Maximizes a window with a matching title and brings it to the foreground.

        Args:
            title (str): A partial or full title of the window to maximize.

        Returns:
            str: 'done' if the window was successfully maximized, or 'fail' if it
                 could not be found after ~1000 attempts.
        """
        loop=0
        full_tile =''
        while title not in full_tile:
            try:
                need_avtive = gw.getWindowsWithTitle(title)[0]
                need_avtive.maximize()
                need_avtive.activate()
                full_tile = gw.getActiveWindow().title
                if title in full_tile:
                    #print(full_tile)
                    return 'done'
            except:
                #print ('try times:' , str(loop))
                loop +=1
                if loop >1000:
                    return 'fail'    
                    
    def close_window(self,title):
        """Finds a window by its title and attempts to close it.

        Args:
            title (str): A partial or full title of the window to close.

        Returns:
            str: A status message indicating success ('closed window') or failure
                 (e.g., "Can't close {title} window").
        """
        loop=0
        full_tile =''
        wd_closed =None
        while title not in full_tile:
            try:
                need_avtive = gw.getWindowsWithTitle(title)[0]
                need_avtive.close()
                all_window = gw.getAllTitles()
                wd_closed =None
                for each_window in all_window:
                    if title in full_tile:
                        wd_closed =False
                        break
                    else:
                        wd_closed =True
                if wd_closed==True:
                    return 'closed window'
            except:
                #print ('try times:' , str(loop))
                loop +=1
                if loop >1000 and wd_closed!=True:
                    return "Can't close {} window".format(title)    
    
    def waiting_window_close(self,window_name,timeout=5):
        """Waits for a window with a specific title to close.

        This method monitors the open windows and returns once the specified
        window is no longer present, or when the timeout is reached.

        Args:
            window_name (str): The partial or full title of the window to monitor.
            timeout (int, optional): The maximum time in seconds to wait. Defaults to 5.

        Returns:
            str: 'window_closed' if the window closes within the timeout period,
                 'window_not_close' otherwise.
        """
        start_time = time.time()
        
        while int(time.time()-start_time) < timeout:
            all_window = gw.getAllTitles()
            result =  'window_closed'
            for each_window in all_window:
                if window_name in each_window:
                    result =  'window_not_close'
                    break # exit for look, get list of windows again
            if result !=  'window_not_close':
                return 'window_closed'  
        return 'window_not_close'
    
    def check_win_title_exits(self,title,timeout=5):
        """Checks if a window with a specific title exists, waiting up to a timeout.

        Args:
            title (str): The partial or full title of the window to search for.
            timeout (int, optional): The maximum time in seconds to wait. Defaults to 5.

        Returns:
            bool: True if a window with the matching title is found within the
                  timeout, False otherwise.
        """
        win = None
        st = time.time()
        time_run=0
        while win is None and time_run<timeout:
            all_window = gw.getAllTitles()
            for each_window in all_window:
                if title in each_window:
                    win = each_window
            if win !=None:
                return True
            ed = time.time()
            time_run = ed-st
        return False    
        
    def check_image_exits(self,image_to_click,timeout=5):
        """Waits for a specific image to appear on the screen.

        This method repeatedly checks for the presence of an image until it is found
        or the timeout duration is exceeded.

        Args:
            image_to_click (str): The name of the image file (without extension) to wait for.
            timeout (int, optional): The maximum time in seconds to wait for the image.
                                     Defaults to 5.

        Returns:
            bool: True if the image appears on screen within the timeout, False otherwise.
        """
        
        
        
        win = False
        st = time.time()
        time_run=0
        self.context.add_log(f"{image_to_click}")
        while win is False and time_run<timeout:
            win = self._check_item_existing(image_to_click)
            if win !=False:
                time.sleep(1)
                return True
            ed = time.time()
            time_run = ed-st
        return False        
                
    def wait_image_disappear(self,image_to_click,timeout=5):
        """Waits for a specific image to disappear from the screen.

        This method repeatedly checks for an image, returning True as soon as it is
        no longer visible, or False if the timeout is reached.

        Args:
            current_file (str): The name of the image file (without extension) to monitor.
            timeout (int, optional): The maximum time in seconds to wait. Defaults to 5.

        Returns:
            bool: True if the image disappears from the screen within the timeout,
                  False otherwise.
        """
    
        win = True
        st = time.time()
        time_run=0
        self.context.add_log(f"{image_to_click}")
        while win is True and time_run<timeout:
            win = self._check_item_existing(image_to_click)
            if win ==False:
                return True
            ed = time.time()
            time_run = ed-st
        return False 

    def wait_ms(self,value):
        """Pauses the execution for a specified number of milliseconds.

        Args:
            value (int): The number of milliseconds to wait.

        Returns:
            str: A confirmation message indicating the sleep duration.
        """
        time.sleep(value/1000)
        return f'sleep {value}ms'
        
    def wait_second(self,value):
        """Pauses the execution for a specified number of seconds.

        Args:
            value (int or float): The number of seconds to wait.

        Returns:
            str: A confirmation message indicating the sleep duration.
        """
        time.sleep(value)
        return f'sleep {value} second'  

    def screenshot(self,image_link):
        im2 = pyautogui.screenshot(f'screenshot/{image_link}.png')
        pyperclipimg.copy(im2)
        return "done"
        
    def click_if_exit(self,check_image,click_image,timeout=1):
        if self.check_image_exits(check_image,timeout):
            self.left_click(click_image)
            return 'clicked'
        return 'Not found'
        
        
    def check_image_in_cross_2_images(self,image_to_click_x,image_to_click_y,check_image,confidence=0.92,padding=10,timeout=5):
        """ Find and return image location
        """
        location=None
        x_position=None
        file_name= image_to_click_x
        full_path_without_ext = os.path.join(self.click_image_folder_path, image_to_click_x)
        full_path_with_ext = full_path_without_ext + ".txt"

        with open(full_path_with_ext) as json_file:
            
            image_file = json.load(json_file)
        for key, data in image_file.items():
            img_obj = self._base64_pgn(data)
            try:
                location= pyautogui.locateCenterOnScreen(img_obj,grayscale=True, confidence=confidence)
                if location!=None:
                    self.context.add_log(f"Image found: {file_name}")
                    self.context.send_click_status(f"Image found: {file_name}")
                    x_position = location.x
                    break
            except:
                pass
        if x_position is None:
            self.context.send_click_status(f"Image not found: {file_name}")
            self.context.add_log(f"Image not found: {file_name}")
            return False      



        location=None
        y_position=None
        file_name= image_to_click_y
        full_path_without_ext = os.path.join(self.click_image_folder_path, image_to_click_y)
        full_path_with_ext = full_path_without_ext + ".txt"

        with open(full_path_with_ext) as json_file:
            
            image_file = json.load(json_file)
        for key, data in image_file.items():
            img_obj = self._base64_pgn(data)
            try:
                location= pyautogui.locateCenterOnScreen(img_obj,grayscale=True, confidence=confidence)
                if location!=None:
                    self.context.add_log(f"Image found: {file_name}")
                    self.context.send_click_status(f"Image found: {file_name}")
                    y_position = location.y
                    break
            except:
                pass
        if y_position is None:
            self.context.send_click_status(f"Image not found: {file_name}")
            self.context.add_log(f"Image not found: {file_name}")
            return False      

        if x_position is not None and y_position is not None:        
            return self._check_image_in_region_of_current_mouse_position(check_image,x_position ,y_position,confidence,padding,timeout)
        else:
            return False

    def _check_image_in_region_of_current_mouse_position(self, check_image: str,x_position ,y_position, confidence=0.8,padding=10,timeout=5) -> bool:
        """
        Checks if an image exists in a small region around the current mouse cursor position.

        This method is useful for verifying tooltips or context-sensitive icons that appear
        when the mouse hovers over a specific UI element. It first decodes the provided
        base64 image, determines a search region based on the image's size plus a small
        padding, centers this region on the current mouse coordinates, and then searches
        for the image within that specific area.

        Args:
            check_image (str): A string containing the base64 encoded content
                                        of the PNG image to search for.
            confidence (float, optional): The confidence level for the image recognition,
                                          ranging from 0.0 to 1.0. Defaults to 0.8.

        Returns:
            bool: True if the image is found within the calculated region around the mouse
                  cursor, False otherwise.
        """
        loop=0
        mouse_x, mouse_y = x_position ,y_position
        file_name= check_image
        full_path_without_ext = os.path.join(self.click_image_folder_path, file_name)
        full_path_with_ext = full_path_without_ext + ".txt"    
        location=None
        while loop<timeout:
            loop+=1
            
            with open(full_path_with_ext) as json_file:
                image_file_dict = json.load(json_file)
            
            found_in_loop = False
            for key, data in image_file_dict.items():
                img_obj = self._base64_pgn(data)
                
                img_width, img_height = img_obj.size
                region_width = img_width + (padding * 2)
                region_height = img_height + (padding * 2)
                search_left = mouse_x - (region_width // 2)
                search_top = mouse_y - (region_height // 2)                
                search_left = max(0, search_left)
                search_top = max(0, search_top)    
                search_region = (search_left, search_top, region_width, region_height)                
                try:
                    location = pyautogui.locateOnScreen(img_obj, region=search_region, confidence=float(confidence))
                    if location!=None:
                        self.context.add_log(f"Image found in region: {file_name}")
                        return True
                except:
                    pass
            
            time.sleep(0.5)
            
        self.context.send_click_status(f"Image not found after timeout: {file_name}")
        self.context.add_log(f"Image not found after timeout: {file_name}")
        return False
            
class Bot_SAP:
    '''
    A class containing methods specifically designed for automating tasks within
    the SAP GUI environment.
    '''
    def __init__(self, context: Context):
        """Initializes the Bot_sap class.

        Args:
            context (Context): A shared context object for logging and state management.
        """
        
        self.bot_utility = Bot_utility(context)
        
        self.context = context # Store the shared context
        self.log_prefix = f"[{self.__class__.__name__}]"
        self.context.add_log(f"{self.log_prefix} Initialized with context.")
        pass            
        
    def log_in(self,SAP,saplogon_exe):
        """Handles the entire SAP logon process.

        Function:
            This method checks if the target SAP system window is already open.
            If not, it opens the SAP Logon Pad, finds the specified system, enters
            credentials, and handles any initial system message pop-ups.

        Args:
            SAP (str): The name (System ID or SID) of the SAP system to log into.

        Returns:
            None
        """
        win = self.bot_utility.check_win_title_exits(SAP,timeout=1)
        if win==False:
            activate_SAP_log_on = self.bot_utility.check_win_title_exits('SAP Logon',timeout=1)
            if activate_SAP_log_on==False:
                subprocess.Popen(saplogon_exe, creationflags=subprocess.CREATE_NEW_CONSOLE) 
            else:
                self.bot_utility.activate_window('SAP Logon')
            time.sleep(2)
            
            self.bot_utility.check_image_exits('SAP GUI/SAP_logon_button',timeout=10)
            if self.bot_utility.check_image_exits('SAP GUI/SAP_Logon_filter_box',timeout=5) ==False:
                self.bot_utility.left_click('SAP GUI/SAP_Logon_taskbar_incon')

            print ('Logon SAP {} system'.format(SAP))
            self.bot_utility.left_click('SAP GUI/SAP_Logon_filter_box',80,0,0.92)
            time.sleep(1)
            keyboard.press_and_release( 'ctrl+a')
            time.sleep(2)
            pyautogui.write(SAP, interval=0.25)
            time.sleep(1)
            pyautogui.press('enter')
            
            while self.bot_utility.check_win_title_exits(SAP,timeout=1) ==False:
                pass
                time.sleep(1)
        need_avtive = gw.getWindowsWithTitle(SAP)[0]
        if self.bot_utility.check_win_title_exits('System Messages',timeout=2):
            self.bot_utility.activate_window('System Messages')
            time.sleep(1)
            pyautogui.press('enter')
            
        self.bot_utility.activate_window(SAP)
        self.bot_utility.maximize_window(SAP)
        self.enter_tcode('')
        return 
    
    def open_table(self,table_name='/n/RB94/XX_PRCL_MO'):
        """Navigates to transaction SE16 and opens a specified table.

        Args:
            table_name (str): The name of the SAP table to open (e.g., 'MARA').

        Returns:
            None
        """
        self.bot_utility.left_click('SAP GUI/SAP_check_icone',80,0,0.92)
        time.sleep(1)
        keyboard.press_and_release( 'ctrl+a')
        time.sleep(1)
        keyboard.write('/nSE16')
        time.sleep(1)
        keyboard.send('ENTER')
        time.sleep(1)
        if self.bot_utility.check_image_exits('SAP GUI/SAP_Table_name')==True:
            self.bot_utility.left_click('SAP GUI/SAP_Table_name',250,0,0.92)
            time.sleep(1)
            keyboard.press_and_release( 'ctrl+a')
            time.sleep(1)
            keyboard.write(table_name)
            time.sleep(1)
            keyboard.send('ENTER')
        return

    def enter_tcode(self,tcode):
        """Enters a transaction code into the SAP command field and executes it.

        Args:
            tcode (str): The transaction code to execute (e.g., 'VA01'). Can include
                         prefixes like '/n' to open in a new session.

        Returns:
            None
        """
        self.bot_utility.left_click('SAP GUI/SAP_Tcode_box',0,0,0.92)
        time.sleep(1)
        keyboard.press_and_release( 'ctrl+a')
        time.sleep(1)
        keyboard.write('{}'.format(tcode))
        time.sleep(1)
        keyboard.send('ENTER')
        time.sleep(1)
        return
        
    def block_BP(self,partner_code,reason_block):
        """Automates the process of blocking a single Business Partner in SAP.

        Function:
            Navigates to the relevant transaction, enters the partner code, finds the
            partner, applies the block with the given reason, and saves the change.

        Args:
            partner_code (str): The Business Partner ID to be blocked.
            reason_block (str): The text to be entered as the blocking reason.

        Returns:
            str: A status message indicating success ('BP number is blocked'), failure
                 ('BP number is not found'), or an error at a specific step
                 (e.g., 'BOT ERROR step 1').
            Hung add new text
        """
        print (partner_code,reason_block)
        time.sleep(5)
        self.enter_tcode('/n/SAPSLL/SPL_CHSB1LO')
        time.sleep(1)


        if self.bot_utility.check_image_exits('SAP GUI/SAP_Business_partner_box',5):
            self.bot_utility.left_click('SAP GUI/SAP_Business_partner_box',0,0,0.92)
            keyboard.write(str(partner_code))
            keyboard.send('F8')
        else:
            return 'BOT ERROR step 1'
            
        if self.bot_utility.check_image_exits('SAP GUI/SAP_BP_Block_not_found',2):
            time.sleep(1)
            self.bot_utility.left_click('SAP GUI/SAP_BP_block_reason_check_icon',0,0,0.92)

            return 'BP number is not found'

        if self.bot_utility.check_image_exits('SAP GUI/SAP_Nagative_icon',10):
            time.sleep(1)
            self.bot_utility.left_click('SAP GUI/SAP_BP_header',0,20,0.92)
            time.sleep(0.2)
            self.bot_utility.left_click('SAP GUI/SAP_Nagative_icon',0,0,0.92)
        else:
            return 'BOT ERROR step 2'

        if self.bot_utility.check_image_exits('SAP GUI/SAP_BP_block_reason_header',5):
            time.sleep(1)
            self.bot_utility.left_click('SAP GUI/SAP_BP_block_reason_header',0,30,0.92)
            time.sleep(0.2)
            keyboard.write(reason_block)
            time.sleep(1)
            self.bot_utility.left_click('SAP GUI/SAP_BP_block_reason_check_icon',0,0,0.92)
            time.sleep(2)
        else:
            return 'BOT ERROR step 3'
        
        if self.bot_utility.check_image_exits('SAP GUI/SAP_BP_block_save_icon',2):  
            
            count =0
            while self.bot_utility.check_image_exits('SAP GUI/SAP_execute_icon',1) is not True and count < 3:
                count +=1
                self.bot_utility.left_click('SAP GUI/SAP_BP_block_save_icon',0,0,0.92)
                time.sleep(0.1)
                self.bot_utility.left_click('SAP GUI/SAP_BP_block_save_icon',0,0,0.92)
                pyautogui.moveTo(800, 800)
                time.sleep(1)
            
            
        else:
            return 'BOT ERROR step 4'   
        if self.bot_utility.check_image_exits('SAP GUI/SAP_execute_icon',10):
            time.sleep(1)
            return 'BP number is blocked'
        else:
            return 'BOT ERROR step 5'
    
        return
    def block_BP_by_excel(self,file_link,system_name,BP,Reason_blocking,Status_column):
        """Blocks multiple Business Partners listed in an Excel file.

        Function:
            This method reads a list of Business Partners from an Excel sheet, logs into
            the specified SAP system, and iteratively calls the block_BP method for each
            partner, updating the sheet with the result.

        Args:
            file_link (str): The file path for the Excel workbook containing the BP list.
            system_name (str): The SAP system (SID) to log into.
            BP (str): The column letter in the Excel sheet that contains the Business Partner IDs.
            Reason_blocking (str): The column letter containing data to be used in the block reason.
            Status_column (str): The column letter where the bot will write the result status.

        Returns:
            bool: True upon successful completion of the process.
        """
        #excel_link = r"./BP blocking list.xlsx"

        
        workbook = openpyxl.load_workbook(file_link)
        sheet = workbook.active
        
        self.log_in(system_name)
        number_of_line = self.count_non_empty_rows_in_column(sheet,BP)

        current_date = datetime.date.today() 
        formatted_date_ymd = current_date.strftime("%Y-%m-%d")
        
        
        for i in range (2,50000):
            
            if sheet[BP + str(i)].value is not None:
                similarity_value = sheet[Reason_blocking + str(i)].value

                if sheet[Status_column + str(i)].value == None or sheet[Status_column + str(i)].value == "" :
                    BP_code = sheet['A' + str(i)].value
                    Blocking_reason = f"BOT blocked this BP based on similarity check on '{formatted_date_ymd}' value '{similarity_value}%' "  #'sheet['B' + str(i)].value
                    sheet[Status_column + str(i)] = self.block_BP(str(BP_code),Blocking_reason)
                    workbook.save(file_link) # Saving inside loop to preserve progress
                    print ('Finished BP: ',sheet['A' + str(i)].value,Blocking_reason)
            else:
                
                break
        workbook.save(file_link)
        workbook.close()
        return True    
        
            
            
    def un_block_BP(self,bp_number,reason_unblock):
        #bot_sap.enter_tcode('/nbp')
        time.sleep(1)
        if bot_utility.check_image_exits('SAP_unblockBP_open_icon',timeout=10):
            time.sleep(1)
            bot_utility.left_click('SAP_unblockBP_open_icon',0,0,0.92)
        else:
            return 'BP number {} could not unblock step 1'.format(bp_number)        
        if bot_utility.check_image_exits('SAP_unblockBP_BP_box',15):
            time.sleep(0.5)
            bot_utility.left_click('SAP_unblockBP_BP_box',20,0,0.92)
            keyboard.press_and_release( 'ctrl+a')
            time.sleep(1)
            keyboard.write('{}'.format(bp_number))
            time.sleep(1)
            keyboard.send('ENTER')
            time.sleep(1)
        else:
            return 'BP number {} could not unblock step 2'.format(bp_number)
            
        if bot_utility.check_image_exits('SAP_unblockBP_GTS_tab',5):
            bot_utility.left_click('SAP_unblockBP_GTS_tab',0,0,0.92)
            
        else:
            return 'BP number {} could not unblock step 3'.format(bp_number)
            
        if bot_utility.check_image_exits('SAP_unblockBP_change_mode',1) is not True:
            bot_utility.left_click('SAP_unblockBP_to_change_mode',0,0,0.92)

            
        if bot_utility.check_image_exits('SAP_unblockBP_remove_block',2):
            count =0
            while bot_utility.check_image_exits('SAP_unblockBP_remove_block',2) and count < 5:
                bot_utility.left_click('SAP_unblockBP_remove_block',0,4,0.92)
                time.sleep(1)
                bot_utility.left_click('SAP_unblockBP_GTS_tab',0,0,0.92)
                time.sleep(0.1)
                bot_utility.left_click('SAP_unblockBP_GTS_tab',0,0,0.92)
                count +=1

        else:
            return 'BP number {} could not unblock step 4'.format(bp_number)
            
        if bot_utility.check_image_exits('SAP_unblockBP_remove_neg',2):
            count =0
            while bot_utility.check_image_exits('SAP_unblockBP_remove_neg',2) and count < 5:
        
                bot_utility.left_click('SAP_unblockBP_remove_neg',0,1,0.92)
                keyboard.send('DELETE')
                time.sleep(1)
                bot_utility.left_click('SAP_unblockBP_GTS_tab',0,0,0.92)
                time.sleep(0.1)
                bot_utility.left_click('SAP_unblockBP_GTS_tab',0,0,0.92)
                count +=1
        else:
            return 'BP number {} could not unblock step 5'.format(bp_number)
            
        if bot_utility.check_image_exits('SAP_unblockBP_remove_neg_done',2) and bot_utility.check_image_exits('SAP_unblockBP_remove_block_done',2):
            print ('remove done')
            bot_utility.left_click('SAP_unblockBP_accept_entries',0,0,0.92)
        
        if bot_utility.check_image_exits('SAP_unblockBP_reason',5):
            bot_utility.left_click('SAP_unblockBP_reason',0,0,0.92)
            time.sleep(1)
        else:
            return 'BP number {} could not unblock step 6'.format(bp_number)
            
        if bot_utility.check_image_exits('SAP_unblockBP_reason_other',5):
            
            bot_utility.left_click('SAP_unblockBP_reason_other',0,0,0.92)
            time.sleep(1)
        else:
            return 'BP number {} could not unblock step 7'.format(bp_number)
            
        if bot_utility.check_image_exits('SAP_BP_block_reason_header',5):
            time.sleep(1)
            bot_utility.left_click('SAP_BP_block_reason_header',0,80,0.92)
            time.sleep(0.2)
            keyboard.write(reason_unblock)
            time.sleep(1)
        else:
            return 'BP number {} could not unblock step 8'.format(bp_number)
            
        if bot_utility.check_image_exits('SAP_BP_block_reason_check_icon',5):
            bot_utility.left_click('SAP_BP_block_reason_check_icon',0,0,0.92)
        
        else:
            return 'BP number {} could not unblock step 9'.format(bp_number)
            
        if bot_utility.check_image_exits('SAP_unblockBP_save',5):
            count =0
            while bot_utility.check_image_exits('SAP_unblockBP_saved',1) is not True and count < 3:
                count +=1
                bot_utility.left_click('SAP_unblockBP_save',0,0,0.92)
                time.sleep(0.1)
                bot_utility.left_click('SAP_unblockBP_save',0,0,0.92)
                time.sleep(3)
        else:
            return 'BP number {} could not unblock step 10'.format(bp_number)
            
        if bot_utility.check_image_exits('SAP_unblockBP_saved',5):
            return 'BP number {} is unblocked'.format(bp_number)
        else:
            return 'BP number {} could not unblock step 11'.format(bp_number)



