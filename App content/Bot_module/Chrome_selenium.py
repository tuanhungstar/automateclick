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
        
        try:
            #webdriver1 = webdriver.Chrome(executable_path=chrome_driver_path, options=chrome_options) 
            driver = webdriver.Chrome(service=service, options=chrome_options) 
            #url = 'https://portal.eu.micgtm.com/#'
            driver.get(url)
            self.context.add_log(f"{self.log_prefix} Chrome is connected to url{url}")
            return driver
        except Exception as e:
            self.context.add_log(f"{self.log_prefix} Error connecting chrome or navigating to url '{url}': {e}")
            return None
        
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

    def click_multiple_elements(self, driver, by_strategy, locator_values_str: str):
        """
        Finds multiple web elements based on semi-colon separated locator values and clicks them sequentially.
        locator_values_str: A string of locators separated by ';' (e.g., 'id1; id2; id3')
        """
        success_count = 0
        try:
            find_strategy_object = self.strategy_map[by_strategy]
            # Split the string by semi-colon and remove any extra whitespace around each locator
            locators = [loc.strip() for loc in locator_values_str.split(';') if loc.strip()]
            
            for locator in locators:
                try:
                    element = driver.find_element(find_strategy_object, locator)
                    if element is not None:
                        element.click()
                        self.context.add_log(f"{self.log_prefix} successfully clicked {by_strategy}: {locator}")
                        success_count += 1
                except NoSuchElementException:
                    self.context.add_log(f"{self.log_prefix} Error: Could not find element with {by_strategy}: {locator}")
                except Exception as e:
                    self.context.add_log(f"{self.log_prefix} Error clicking element {locator}: {e}")
            
            # Returns True if at least one click was successful
            return success_count > 0
        except Exception as e:
            self.context.add_log(f"{self.log_prefix} A general error occurred in click_multiple_elements: {e}")
            return False
            
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

    def save_website_html(self, driver, file_path: str):
        """Saves the current page HTML source to the specified file path."""
        try:
            import os
            # Ensure directory exists
            directory = os.path.dirname(file_path)
            if directory and not os.path.exists(directory):
                os.makedirs(directory)
                
            page_source = driver.page_source
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(page_source)
            self.context.add_log(f"{self.log_prefix} Successfully saved HTML to {file_path}")
            return True
        except Exception as e:
            self.context.add_log(f"{self.log_prefix} Failed to save HTML: {e}")
            return False

    def clean_html_for_llm(self, driver, assign_to_variable: str = ""):
        """Cleans the HTML of the current page for LLM processing.
        Provide a variable name in 'assign_to_variable' to save it without printing to the log."""
        try:
            from bs4 import BeautifulSoup
        except ImportError:
            self.context.add_log(f"{self.log_prefix} Error: beautifulsoup4 is not installed. Please run 'pip install beautifulsoup4'.")
            return None
            
        try:
            raw_html = driver.page_source
            soup = BeautifulSoup(raw_html, "html.parser")
            
            # Remove all script and style elements
            for script_or_style in soup(["script", "style", "noscript", "svg"]):
                script_or_style.extract()
                
            # Get the text or a simplified HTML
            clean_html = str(soup.body) if soup.body else str(soup)
            self.context.add_log(f"{self.log_prefix} Successfully cleaned HTML.")
            
            if assign_to_variable and assign_to_variable.strip() != "":
                self.context.set_variable(assign_to_variable.strip(), clean_html)
                self.context.add_log(f"{self.log_prefix} HTML saved securely to variable '{assign_to_variable}' (hidden from log).")
                return f"Success - Saved to {assign_to_variable}"
            else:
                return clean_html
        except Exception as e:
            self.context.add_log(f"{self.log_prefix} Failed to clean HTML: {e}")
            return None

    def close_popup(self, driver, popup_selector: str = ".close-popup-class, .modal-close, button[aria-label='Close'], .close", popup_container_id: str = None):
        """Attempts to close a popup using clicking and JavaScript removal."""
        try:
            from selenium.webdriver.common.by import By
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as EC
            
            self.context.add_log(f"{self.log_prefix} Attempting to close popup...")
            success = False
            
            # Option A: Try clicking the close button
            try:
                close_btn = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, popup_selector))
                )
                close_btn.click()
                self.context.add_log(f"{self.log_prefix} Successfully clicked popup close button.")
                success = True
            except Exception:
                self.context.add_log(f"{self.log_prefix} Popup close button not found or click failed.")
                
            # Option B: Try using JavaScript to remove the popup container if ID is provided
            # Even if Option A succeeded, sometimes the overlay persists, so we can try Option B if specified
            if popup_container_id:
                try:
                    popup = driver.find_element(By.ID, popup_container_id) 
                    driver.execute_script("arguments[0].remove();", popup)
                    self.context.add_log(f"{self.log_prefix} Successfully removed popup container via JS.")
                    success = True
                except Exception:
                    self.context.add_log(f"{self.log_prefix} Popup container not found via JS.")
            
            return success
        except Exception as e:
            self.context.add_log(f"{self.log_prefix} Error handling popup: {e}")
            return False

    def check_popup_exists(self, driver, popup_selector: str = ".close-popup-class, .modal-close, button[aria-label='Close'], .close", timeout: int = 3):
        """Checks if a popup exists on the current page."""
        try:
            from selenium.webdriver.common.by import By
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as EC
            
            self.context.add_log(f"{self.log_prefix} Checking if popup exists (timeout: {timeout}s)...")
            
            # Use WebDriverWait to see if element matches the selector
            WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, popup_selector))
            )
            self.context.add_log(f"{self.log_prefix} Popup found.")
            return True
        except Exception:
            self.context.add_log(f"{self.log_prefix} Popup not found within timeout.")
            return False


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
                    return self.gia_dien_may_xanh(price) 
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
    def gia_dien_may_xanh(self,x):
        x=x.replace("Online Giá Rẻ Quá\n","")
        x=x.replace("Xả kho giảm hết - Vui hơn Tết\n","")
        x=x.replace(".","")
        x=x.split('₫')[0]
        return x     
    
    def take_screenshot(self,driver,file_link):
        try:
            driver.save_screenshot(file_link)
            return True
        except:
            return False
        return

class Download_Chromedriver:
    def __init__(self, context: Context):
        self.context = context # Store the shared context
        self.log_prefix = f"[{self.__class__.__name__}]"
        self.context.add_log(f"{self.log_prefix} Initialized with context.")
        
    def download_and_check(self, path_to_save: str):
        from webdriver_manager.chrome import ChromeDriverManager
        import shutil
        import os
        
        try:
            self.context.add_log(f"{self.log_prefix} Checking/Downloading ChromeDriver...")
            # Download or check cache for chromedriver
            driver_path = ChromeDriverManager().install()
            
            # Ensure the target directory exists
            if not os.path.exists(path_to_save):
                os.makedirs(path_to_save)
                
            # Copy to the assigned path (usually driver_path is an exe on Windows)
            filename = os.path.basename(driver_path) # usually chromedriver.exe
            target_path = os.path.join(path_to_save, filename)
            
            # Only copy if source and destination are different
            if os.path.abspath(driver_path) != os.path.abspath(target_path):
                shutil.copy2(driver_path, target_path)
                
            self.context.add_log(f"{self.log_prefix} Successfully provided chromedriver at {target_path}")
            return target_path
            
        except Exception as e:
            self.context.add_log(f"{self.log_prefix} Failed to download/check chromedriver: {e}")
            return None

class HTML_Extractor:
    def __init__(self, context: Context):
        self.context = context
        self.log_prefix = f"[{self.__class__.__name__}]"
        self.context.add_log(f"{self.log_prefix} Initialized with context.")
        
    def get_html_to_variable(self, driver, variable_name: str, clean_html: bool = True):
        """
        Extracts the HTML from the current driver and assigns it to a context variable.
        If clean_html is True, it removes scripts, styles, and SVG data to save tokens for Gemini.
        """
        try:
            html = driver.page_source
            if clean_html:
                try:
                    from bs4 import BeautifulSoup
                    soup = BeautifulSoup(html, "html.parser")
                    # Remove tokens that are usually useless for finding buttons but take many tokens
                    for tag in soup(["script", "style", "noscript", "svg", "path"]):
                        tag.decompose()
                    html = str(soup.body) if soup.body else str(soup)
                except ImportError:
                    self.context.add_log(f"{self.log_prefix} BeautifulSoup4 not found, returning raw HTML.")
            
            self.context.set_variable(variable_name, html)
            self.context.add_log(f"{self.log_prefix} Successfully assigned HTML to variable '{variable_name}'.")
            return True
        except Exception as e:
            self.context.add_log(f"{self.log_prefix} Error extracting HTML: {e}")
            return False
