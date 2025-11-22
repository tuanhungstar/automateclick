from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.expected_conditions import presence_of_element_located
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.remote.webelement import WebElement
from selenium.common.exceptions import NoSuchElementException
import clipboard
import keyboard
import time
import pandas as pd
import numpy as np
import ast

from my_lib.shared_context import ExecutionContext as Context


'''
self.context.add_log(f"{self.log_prefix} Initialized with context.")
'''

class Chrome_selenium():
    def __init__(self, context: Context):
        
        self.context = context # Store the shared context
        self.log_prefix = f"[{self.__class__.__name__}]"
        self.context.add_log(f"{self.log_prefix} Initialized with context.")
        self.strategy_map = {
                                "class_name": By.CLASS_NAME,
                                "id": By.ID,
                                "xpath": By.XPATH,
                                "name": By.NAME,
                                "css_selector": By.CSS_SELECTOR,
                                "link_text": By.LINK_TEXT,
                                "tag_name": By.TAG_NAME
                            }
        return
        
    def connect_chrome(self,chrome_driver_link:str,url):

        #chrome_driver_path = r'chromedriver.exe'
        chrome_driver_path = chrome_driver_link
        service = Service(executable_path=chrome_driver_path)
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-popup-blocking")
        
        #webdriver1 = webdriver.Chrome(executable_path=chrome_driver_path, options=chrome_options) 
        driver = webdriver.Chrome(service=service, options=chrome_options) 
        #url = 'https://portal.eu.micgtm.com/#'
        driver.get(url)
        self.context.add_log(f"{self.log_prefix} Chrome is connected to url{url}")
        return driver
        
    def find_and_send_text(self,driver, 
                       by_strategy: By, 
                       locator_value: str, 
                       text_to_send: str):
        """Finds a single web element and sends text to it.
        By.ID, By.NAME,By.CLASS_NAME,By.TAG_NAME:,By.LINK_TEXT,By.CSS_SELECTOR,By.XPATH: 
        
        
        
        """
        

                
        try:
            element = driver.find_element(by_strategy, locator_value)
            element.clear()
            element.send_keys(text_to_send)

            self.context.add_log(f"{self.log_prefix} send text to {by_strategy} the text: {locator_value}")
            
            return element
        except NoSuchElementException:
            self.context.add_log(f"{self.log_prefix} Error: Could not find element with {by_strategy}: {locator_value}")
            return None
        except Exception as e:

            self.context.add_log(f"{self.log_prefix} An error occurred: {e}")
            return None



            
    def find_element_action(self,driver, 
                       by_strategy, 
                       locator_value: str,
                       Actions = "Click,Clear,Send_Key",
                       text_to_send:str = '' ):
        """
            Finds a single web element and sends text to it.
                                "class_name", "id","xpath","name","css_selector","link_text","tag_name"
        
            elements = webdriver.find_elements('xpath',"//*[contains(text(),'{}')]".format(text))
            
            Action= 'Click' or 'Clear' or 'Send_Key'

            
        
        """
        try:
            find_strategy_object = self.strategy_map[by_strategy]
            element = driver.find_element(find_strategy_object, locator_value)
            if element is not None:
                Actions = Actions.split(",")
                for Action in Actions:
                    if Action.strip() =='Click':

                        element.click()
                    elif Action.strip() =='Clear':

                        element.send_keys(text_to_send)

                    elif Action.strip() =='Send_Key':

                        element.send_keys(text_to_send)

                self.context.add_log(f"{self.log_prefix} clicked {by_strategy}")
            
            return element
        except NoSuchElementException:
            self.context.add_log(f"{self.log_prefix} Error: Could not find element with {by_strategy}")
            return None
        except Exception as e:

            self.context.add_log(f"{self.log_prefix} An error occurred: {e}")
            return None
            
    def get_text_from_element(self,element):
        
        return element.text

    def found_element_action(self,element,
                       Actions = "Click,Clear,Send_Key",
                       text_to_send:str = '' ):
        """
            pass found element to this method and assign action
            
            Action= 'Click' or 'Clear' or 'Send_Key'

            
        
        """
        try:
            if element is not None:
                Actions = Actions.split(",")
                for Action in Actions:
                    if Action.strip() =='Click':

                        element.click()
                    elif Action.strip() =='Clear':

                        element.send_keys(text_to_send)

                    elif Action.strip() =='Send_Key':

                        element.send_keys(text_to_send)

                self.context.add_log(f"{self.log_prefix} clicked ")
            
            return True
        except NoSuchElementException:
            self.context.add_log(f"{self.log_prefix} Error: Could not find element with ")
            return None
        except Exception as e:

            self.context.add_log(f"{self.log_prefix} An error occurred: {e}")
            return None
            
    def get_text_from_element(self,element):
        
        return element.text            


            
    def check_element_exist(self,driver,by_strategy, locator_value,timeout):
        nloop = 0
        found =False
        while nloop<timeout:
            try:
                driver.find_element(by_strategy, locator_value).text
                nloop = timeout
                found =True
                break
            except:
                time.sleep(1)
                nloop +=1    

        return found
        
        
    def find_elememt_by_display_name(self,driver,text,click_request=False):
        
        elements = driver.find_elements(By.XPATH,"//*[contains(text(),'{}')]".format(text))
        found_element =None
        for each_lement in elements:
            if text in each_lement.text:
                found_element = each_lement
                if click_request==True:
                    try:
                        found_element.click()
                    except:
                        pass
                return True
                break
        return False 
        
    def close_driver(self,driver):
            
        driver.quit()
        return 

class Chrome_MIC():
    
    def __init__(self, context: Context):
        
        self.context = context # Store the shared context
        self.log_prefix = f"[{self.__class__.__name__}]"
        self.context.add_log(f"{self.log_prefix} Initialized with context.")
        self.strategy_map = {
                        "class_name": By.CLASS_NAME,
                        "id": By.ID,
                        "xpath": By.XPATH,
                        "name": By.NAME,
                        "css_selector": By.CSS_SELECTOR,
                        "link_text": By.LINK_TEXT,
                        "tag_name": By.TAG_NAME
                    }
        return
        
    def switch_driver_tab(self,driver,tab):
        driver.switch_to.window(driver.window_handles[tab])
        return driver
        
    def get_input_element_by_its_content(self,driver,text):
        elements = driver.find_elements(By.CSS_SELECTOR,'input[type="text"]')
        for element in elements:
           if element.get_attribute('value').strip() == text:
               return  element
        return None
        
    def selenium_left_click_button(self,text):
        loop =0
        try:
            while loop <5:
                webdriver = self.webdriver
                elements = webdriver.find_elements('xpath',"//*[contains(text(),'{}')]".format(text))
                for each_lement in elements:
                    if text in each_lement.text:
                        each_lement.click()
                        return 'left_click_done'
                        break
                time.sleep(1)        
                loop+=1 
        except:
            time.sleep(3)
            loop+=1
        return 'left_click_not_done'  
        
    def enter_input_box(self,element,text):
        element.clear()
        element.send_keys(text)
        
        return
        
    def get_number_records(self,webdriver):
        time.sleep(0.1)
        start_time = time.time()
            
        button_name = 'Query|'
        start_time = time.time()
        while 'Query' in button_name:
            try: 
                button_name = self._get_element_text(webdriver,'Refresh')
            except:
                try:
                    button_name = self._get_element_text(webdriver,'Refresh')
                except:
                    a =''
            total_time = time.time()-start_time
            if total_time>10:
                raise Exception("Query button not response")
        #print ('Done query')
        time.sleep(0.2)        
        #Records_icon = ahk.run_script(AHK_data['AHK_Records_icon'] + ahk_script.script_click_records + ahk_script.ahk_click_f , blocking=True)
        
        start_time = time.time()
        numberofpart=''
        while numberofpart=='':
            try:
                records_status = self._left_click(webdriver,'#')
            except:
                time.sleep(0.1) 

            time.sleep(0.2)

            records = webdriver.find_elements(By.XPATH,"//*[contains(text(),'Records')]")
            for each_lement in records:
                if 'Records: ' in each_lement.text:
                    numberofpart = each_lement.text.replace('Records: ','').strip().replace(',','')
                    break
            total_time = time.time()-start_time
            if total_time>20:
                raise Exception("Records button not response")
        return int(numberofpart)
        
    def _get_element_text(self,webdriver,text):
        elements = None
        try:
            elements = webdriver.find_elements('xpath',"//*[contains(text(),'{}')]".format(text))
            for each_lement in elements:
                if text in each_lement.text:
                    return each_lement.text
                    break
        except:
            return 'Not Found'                
                
        return 'Not Found'     
        
    def _left_click(self,webdriver,text):
        loop =0
        try:
            while loop <5:
                elements = webdriver.find_elements('xpath',"//*[contains(text(),'{}')]".format(text))
                for each_lement in elements:
                    if text in each_lement.text:
                        each_lement.click()
                        return 'left_click_done'
                        break
                time.sleep(1)        
                loop+=1 
        except:
            time.sleep(3)
            loop+=1
        return 'left_click_not_done'  
        

class Chrome_TGDD:
    
    def __init__(self, context: Context):
        
        self.context = context # Store the shared context
        self.log_prefix = f"[{self.__class__.__name__}]"
        self.context.add_log(f"{self.log_prefix} Initialized with context.")

        self.strategy_map = {
                        "class_name": By.CLASS_NAME,
                        "id": By.ID,
                        "xpath": By.XPATH,
                        "name": By.NAME,
                        "css_selector": By.CSS_SELECTOR,
                        "link_text": By.LINK_TEXT,
                        "tag_name": By.TAG_NAME
                    }
                    
                    
        self.dmx_price = [
                          "class_name>><<box-price-present",
                          "xpath>><</html/body/section/div[2]/div[2]/div[4]/div[3]/div[1]/div[1]/span/b",
                          "class_name>><<bs_price"
                          
                        ]
                        
                        
        self.dmx_description = ["class_name>><<product-name"]
                        
        return
        
    def get_price(self,driver,store ='DMX'):
        
        '''
        DMX = Dienmayxanh
        
        '''
        
        if store =='DMX':
            price_elements=self.dmx_price

        for each_price_element in price_elements:
            
            by_strategy = each_price_element.split('>><<')[0]
            locator_value = each_price_element.split('>><<')[1]
            print(by_strategy,locator_value)
            find_strategy_object = self.strategy_map[by_strategy]
            element=None
            try:
                element = driver.find_element(find_strategy_object, locator_value)    

            except:
                pass
            if element is not None:
                price = element.text
                if price !='':
                    return price 
        return None
        
        
        return
        
    def get_description(self,driver,store ='DMX'):
        
        '''
        DMX = Dienmayxanh
        
        '''
        
        if store =='DMX':

            desc_element = self.dmx_description
            
        
        
        for each_price_element in desc_element:
            
            by_strategy = each_price_element.split('>><<')[0]
            locator_value = each_price_element.split('>><<')[1]
            print(by_strategy,locator_value)
            find_strategy_object = self.strategy_map[by_strategy]
            element=None
            try:
                element = driver.find_element(find_strategy_object, locator_value)    

            except:
                pass
            if element is not None:
                desc = element.text
                if desc !='':
                    return desc 
        return None
        
        
        return       