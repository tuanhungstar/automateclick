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
            image_file = self._base64_pgn(data)
            location=None
            try:
                location= pyautogui.locateCenterOnScreen(image_file,grayscale=True, confidence=0.98)
            except:
                pass
            if location!=None:
                return True
        return False
        
    def left_click(self,image_to_click,offset_x=0,offset_y=0,confidence=0.92):
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

        with open(full_path_with_ext) as json_file:
            
            image_file = json.load(json_file)
        for key, data in image_file.items():
            #print (key)
            image_file = self._base64_pgn(data)
            try:
                location= pyautogui.locateCenterOnScreen(image_file,grayscale=True, confidence=confidence)
                if location!=None:
                    pyautogui.click(location.x + offset_x, location.y + offset_y)
                    self.context.add_log(f"{file_name}")
                    self.context.send_click_status(f"Image found: {file_name}")
                    return 'left_click_done'
            except:
                self.context.add_log(f"Image not found: {file_name}")
                self.context.send_click_status(f"Image not found: {file_name}")
                pass
        if location is None:
            self.context.send_click_status(f"Image not found: {file_name}")
            self.context.add_log(f"Image not found: {file_name}")
            #print ('left_click_done fail: {}'.format(image_file),str(key))
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
            #print (key)
            image_file = self._base64_pgn(data)
            try:
                location= pyautogui.locateCenterOnScreen(image_file,grayscale=True, confidence=confidence)
                if location!=None:
                    pyautogui.rightClick(location.x + offset_x, location.y + offset_y)
                    self.context.add_log(f"{file_name}")
                    return 'right_click_done'
            except:
                pass
        if location is None:
            self.context.add_log(f"{file_name}")
            #print ('left_click_done fail: {}'.format(image_file),str(key))
        return 'right_click_done fail: {}'.format(file_name)

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
            
            self.bot_utility.check_image_exits('SAP GUI/SAP GUI/SAP_logon_button',timeout=10)
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
    
    def open_table_se16(self,table_name):
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
            keyboard.write('/n/RB94/XX_PRCL_MO')
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