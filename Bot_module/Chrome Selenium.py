from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
import time
import os
from typing import List, Dict, Any, Optional, Union, Tuple
from datetime import datetime
import base64

class ChromeHandler:
    """
    A comprehensive class for handling Chrome browser automation using Selenium.
    
    This class provides methods to control Chrome browser, navigate web pages,
    interact with elements, handle downloads, manage cookies, and perform
    various web automation tasks.
    
    Attributes:
        driver (webdriver.Chrome): Chrome WebDriver instance
        wait (WebDriverWait): WebDriverWait instance for explicit waits
        actions (ActionChains): ActionChains instance for complex interactions
        operation_log (List[str]): Log of all operations performed
        
    Example:
        >>> chrome = ChromeHandler()
        >>> chrome.start_browser()
        >>> chrome.navigate_to("https://www.google.com")
        >>> chrome.find_element_and_send_keys("q", "Python programming")
        >>> chrome.click_element("input[value='Google Search']")
        >>> chrome.close_browser()
    """
    
    def __init__(self, chrome_driver_path: str = None):
        """
        Initialize the ChromeHandler instance.
        
        Args:
            chrome_driver_path (str, optional): Path to ChromeDriver executable.
                                               If None, assumes chromedriver is in PATH
        """
        self.driver = None
        self.wait = None
        self.actions = None
        self.operation_log = []
        self.chrome_driver_path = chrome_driver_path
        self.default_timeout = 10
    
    def start_browser(self, headless: bool = False, window_size: Tuple[int, int] = (1920, 1080),
                     download_dir: str = None, disable_images: bool = False,
                     disable_gpu: bool = False, incognito: bool = False,
                     user_data_dir: str = None, profile_dir: str = None) -> bool:
        """
        Start Chrome browser with specified options.
        
        Args:
            headless (bool): Run browser in headless mode (default: False)
            window_size (Tuple[int, int]): Browser window size (default: (1920, 1080))
            download_dir (str, optional): Custom download directory
            disable_images (bool): Disable image loading for faster browsing (default: False)
            disable_gpu (bool): Disable GPU acceleration (default: False)
            incognito (bool): Start in incognito mode (default: False)
            user_data_dir (str, optional): Custom user data directory
            profile_dir (str, optional): Specific Chrome profile directory
        
        Returns:
            bool: True if browser started successfully, False otherwise
            
        Example:
            >>> chrome = ChromeHandler()
            >>> success = chrome.start_browser(headless=True, disable_images=True)
            >>> if success:
            ...     print("Browser started successfully")
        """
        try:
            chrome_options = Options()
            
            # Basic options
            if headless:
                chrome_options.add_argument("--headless")
            
            chrome_options.add_argument(f"--window-size={window_size[0]},{window_size[1]}")
            
            if disable_gpu:
                chrome_options.add_argument("--disable-gpu")
            
            if incognito:
                chrome_options.add_argument("--incognito")
            
            # User data and profile directories
            if user_data_dir:
                chrome_options.add_argument(f"--user-data-dir={user_data_dir}")
            
            if profile_dir:
                chrome_options.add_argument(f"--profile-directory={profile_dir}")
            
            # Download directory
            if download_dir:
                if not os.path.exists(download_dir):
                    os.makedirs(download_dir)
                prefs = {"download.default_directory": download_dir}
                chrome_options.add_experimental_option("prefs", prefs)
            
            # Disable images for faster loading
            if disable_images:
                prefs = {"profile.managed_default_content_settings.images": 2}
                chrome_options.add_experimental_option("prefs", prefs)
            
            # Additional performance options
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            
            # Initialize driver
            if self.chrome_driver_path:
                service = Service(self.chrome_driver_path)
                self.driver = webdriver.Chrome(service=service, options=chrome_options)
            else:
                self.driver = webdriver.Chrome(options=chrome_options)
            
            # Remove automation indicators
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            # Initialize wait and actions
            self.wait = WebDriverWait(self.driver, self.default_timeout)
            self.actions = ActionChains(self.driver)
            
            self._log_operation("Chrome browser started successfully")
            return True
            
        except Exception as e:
            print(f"✗ Error starting Chrome browser: {str(e)}")
            return False
    
    def navigate_to(self, url: str, wait_for_load: bool = True) -> bool:
        """
        Navigate to a specific URL.
        
        Args:
            url (str): URL to navigate to
            wait_for_load (bool): Wait for page to load completely (default: True)
        
        Returns:
            bool: True if navigation successful, False otherwise
            
        Example:
            >>> chrome.navigate_to("https://www.example.com")
            True
        """
        try:
            if not self.driver:
                print("✗ Browser not started. Call start_browser() first.")
                return False
            
            self.driver.get(url)
            
            if wait_for_load:
                self.wait.until(lambda driver: driver.execute_script("return document.readyState") == "complete")
            
            self._log_operation(f"Navigated to: {url}")
            return True
            
        except Exception as e:
            print(f"✗ Error navigating to {url}: {str(e)}")
            return False
    
    def find_element(self, selector: str, by: str = "css", timeout: int = None) -> Any:
        """
        Find a single element on the page.
        
        Args:
            selector (str): Element selector (CSS, XPath, ID, etc.)
            by (str): Selection method ('css', 'xpath', 'id', 'name', 'class', 'tag') (default: 'css')
            timeout (int, optional): Custom timeout in seconds
        
        Returns:
            WebElement: Found element, None if not found
            
        Example:
            >>> element = chrome.find_element("input[name='q']", by="css")
            >>> if element:
            ...     print("Search box found")
        """
        try:
            if not self.driver:
                print("✗ Browser not started.")
                return None
            
            by_mapping = {
                'css': By.CSS_SELECTOR,
                'xpath': By.XPATH,
                'id': By.ID,
                'name': By.NAME,
                'class': By.CLASS_NAME,
                'tag': By.TAG_NAME,
                'link_text': By.LINK_TEXT,
                'partial_link': By.PARTIAL_LINK_TEXT
            }
            
            by_method = by_mapping.get(by.lower(), By.CSS_SELECTOR)
            wait_time = timeout or self.default_timeout
            
            wait = WebDriverWait(self.driver, wait_time)
            element = wait.until(EC.presence_of_element_located((by_method, selector)))
            
            return element
            
        except TimeoutException:
            print(f"✗ Element not found: {selector} (timeout: {wait_time}s)")
            return None
        except Exception as e:
            print(f"✗ Error finding element {selector}: {str(e)}")
            return None
    
    def find_elements(self, selector: str, by: str = "css") -> List[Any]:
        """
        Find multiple elements on the page.
        
        Args:
            selector (str): Element selector
            by (str): Selection method (default: 'css')
        
        Returns:
            List[WebElement]: List of found elements
            
        Example:
            >>> elements = chrome.find_elements("div.result", by="css")
            >>> print(f"Found {len(elements)} results")
        """
        try:
            if not self.driver:
                print("✗ Browser not started.")
                return []
            
            by_mapping = {
                'css': By.CSS_SELECTOR,
                'xpath': By.XPATH,
                'id': By.ID,
                'name': By.NAME,
                'class': By.CLASS_NAME,
                'tag': By.TAG_NAME
            }
            
            by_method = by_mapping.get(by.lower(), By.CSS_SELECTOR)
            elements = self.driver.find_elements(by_method, selector)
            
            return elements
            
        except Exception as e:
            print(f"✗ Error finding elements {selector}: {str(e)}")
            return []
    
    def click_element(self, selector: str, by: str = "css", timeout: int = None) -> bool:
        """
        Click an element on the page.
        
        Args:
            selector (str): Element selector
            by (str): Selection method (default: 'css')
            timeout (int, optional): Custom timeout in seconds
        
        Returns:
            bool: True if click successful, False otherwise
            
        Example:
            >>> success = chrome.click_element("button#submit", by="css")
            >>> if success:
            ...     print("Button clicked successfully")
        """
        try:
            element = self.find_element(selector, by, timeout)
            if element:
                # Wait for element to be clickable
                wait_time = timeout or self.default_timeout
                wait = WebDriverWait(self.driver, wait_time)
                clickable_element = wait.until(EC.element_to_be_clickable(element))
                
                # Scroll element into view
                self.driver.execute_script("arguments[0].scrollIntoView(true);", clickable_element)
                time.sleep(0.5)  # Small delay for smooth scrolling
                
                clickable_element.click()
                self._log_operation(f"Clicked element: {selector}")
                return True
            
            return False
            
        except Exception as e:
            print(f"✗ Error clicking element {selector}: {str(e)}")
            return False
    
    def send_keys_to_element(self, selector: str, text: str, by: str = "css", 
                           clear_first: bool = True, timeout: int = None) -> bool:
        """
        Send text to an input element.
        
        Args:
            selector (str): Element selector
            text (str): Text to send
            by (str): Selection method (default: 'css')
            clear_first (bool): Clear existing text before typing (default: True)
            timeout (int, optional): Custom timeout in seconds
        
        Returns:
            bool: True if text sent successfully, False otherwise
            
        Example:
            >>> success = chrome.send_keys_to_element("input[name='username']", "john_doe")
            >>> if success:
            ...     print("Username entered successfully")
        """
        try:
            element = self.find_element(selector, by, timeout)
            if element:
                if clear_first:
                    element.clear()
                
                element.send_keys(text)
                self._log_operation(f"Sent text to element {selector}: {text}")
                return True
            
            return False
            
        except Exception as e:
            print(f"✗ Error sending keys to element {selector}: {str(e)}")
            return False
    
    def get_element_text(self, selector: str, by: str = "css", timeout: int = None) -> str:
        """
        Get text content of an element.
        
        Args:
            selector (str): Element selector
            by (str): Selection method (default: 'css')
            timeout (int, optional): Custom timeout in seconds
        
        Returns:
            str: Element text content, empty string if not found
            
        Example:
            >>> title = chrome.get_element_text("h1.page-title")
            >>> print(f"Page title: {title}")
        """
        try:
            element = self.find_element(selector, by, timeout)
            if element:
                return element.text
            
            return ""
            
        except Exception as e:
            print(f"✗ Error getting text from element {selector}: {str(e)}")
            return ""
    
    def get_element_attribute(self, selector: str, attribute: str, by: str = "css", 
                            timeout: int = None) -> str:
        """
        Get attribute value of an element.
        
        Args:
            selector (str): Element selector
            attribute (str): Attribute name to get
            by (str): Selection method (default: 'css')
            timeout (int, optional): Custom timeout in seconds
        
        Returns:
            str: Attribute value, empty string if not found
            
        Example:
            >>> href = chrome.get_element_attribute("a.download-link", "href")
            >>> print(f"Download link: {href}")
        """
        try:
            element = self.find_element(selector, by, timeout)
            if element:
                return element.get_attribute(attribute) or ""
            
            return ""
            
        except Exception as e:
            print(f"✗ Error getting attribute {attribute} from element {selector}: {str(e)}")
            return ""
    
    def wait_for_element(self, selector: str, by: str = "css", timeout: int = None, 
                        condition: str = "presence") -> bool:
        """
        Wait for an element to meet a specific condition.
        
        Args:
            selector (str): Element selector
            by (str): Selection method (default: 'css')
            timeout (int, optional): Custom timeout in seconds
            condition (str): Wait condition ('presence', 'visible', 'clickable') (default: 'presence')
        
        Returns:
            bool: True if condition met within timeout, False otherwise
            
        Example:
            >>> found = chrome.wait_for_element("div.loading", condition="visible", timeout=30)
            >>> if found:
            ...     print("Loading indicator appeared")
        """
        try:
            if not self.driver:
                print("✗ Browser not started.")
                return False
            
            by_mapping = {
                'css': By.CSS_SELECTOR,
                'xpath': By.XPATH,
                'id': By.ID,
                'name': By.NAME,
                'class': By.CLASS_NAME,
                'tag': By.TAG_NAME
            }
            
            by_method = by_mapping.get(by.lower(), By.CSS_SELECTOR)
            wait_time = timeout or self.default_timeout
            wait = WebDriverWait(self.driver, wait_time)
            
            if condition == "presence":
                wait.until(EC.presence_of_element_located((by_method, selector)))
            elif condition == "visible":
                wait.until(EC.visibility_of_element_located((by_method, selector)))
            elif condition == "clickable":
                wait.until(EC.element_to_be_clickable((by_method, selector)))
            else:
                print(f"✗ Unknown condition: {condition}")
                return False
            
            return True
            
        except TimeoutException:
            print(f"✗ Element {selector} did not meet condition '{condition}' within {wait_time}s")
            return False
        except Exception as e:
            print(f"✗ Error waiting for element {selector}: {str(e)}")
            return False
    
    def scroll_to_element(self, selector: str, by: str = "css") -> bool:
        """
        Scroll to a specific element.
        
        Args:
            selector (str): Element selector
            by (str): Selection method (default: 'css')
        
        Returns:
            bool: True if scroll successful, False otherwise
            
        Example:
            >>> chrome.scroll_to_element("footer")
        """
        try:
            element = self.find_element(selector, by)
            if element:
                self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
                time.sleep(0.5)  # Small delay for smooth scrolling
                self._log_operation(f"Scrolled to element: {selector}")
                return True
            
            return False
            
        except Exception as e:
            print(f"✗ Error scrolling to element {selector}: {str(e)}")
            return False
    
    def scroll_page(self, direction: str = "down", pixels: int = None) -> bool:
        """
        Scroll the page in a specified direction.
        
        Args:
            direction (str): Scroll direction ('up', 'down', 'top', 'bottom') (default: 'down')
            pixels (int, optional): Number of pixels to scroll (ignored for 'top'/'bottom')
        
        Returns:
            bool: True if scroll successful, False otherwise
            
        Example:
            >>> chrome.scroll_page("down", 500)  # Scroll down 500 pixels
            >>> chrome.scroll_page("bottom")     # Scroll to bottom of page
        """
        try:
            if not self.driver:
                print("✗ Browser not started.")
                return False
            
            if direction == "top":
                self.driver.execute_script("window.scrollTo(0, 0);")
            elif direction == "bottom":
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            elif direction == "down":
                scroll_amount = pixels or 300
                self.driver.execute_script(f"window.scrollBy(0, {scroll_amount});")
            elif direction == "up":
                scroll_amount = pixels or 300
                self.driver.execute_script(f"window.scrollBy(0, -{scroll_amount});")
            else:
                print(f"✗ Invalid scroll direction: {direction}")
                return False
            
            self._log_operation(f"Scrolled page: {direction}")
            return True
            
        except Exception as e:
            print(f"✗ Error scrolling page: {str(e)}")
            return False
    
    def take_screenshot(self, filename: str = None, full_page: bool = False) -> str:
        """
        Take a screenshot of the current page.
        
        Args:
            filename (str, optional): Screenshot filename. If None, generates timestamp-based name
            full_page (bool): Take full page screenshot (default: False)
        
        Returns:
            str: Path to saved screenshot file, empty string if failed
            
        Example:
            >>> screenshot_path = chrome.take_screenshot("homepage.png")
            >>> print(f"Screenshot saved: {screenshot_path}")
        """
        try:
            if not self.driver:
                print("✗ Browser not started.")
                return ""
            
            if filename is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"screenshot_{timestamp}.png"
            
            if full_page:
                # Get full page height
                total_height = self.driver.execute_script("return document.body.scrollHeight")
                viewport_height = self.driver.execute_script("return window.innerHeight")
                
                # Take screenshots of each viewport and stitch them together
                # For simplicity, we'll just take a regular screenshot
                # Full implementation would require image stitching
                success = self.driver.save_screenshot(filename)
            else:
                success = self.driver.save_screenshot(filename)
            
            if success:
                self._log_operation(f"Screenshot saved: {filename}")
                return filename
            else:
                print("✗ Failed to save screenshot")
                return ""
            
        except Exception as e:
            print(f"✗ Error taking screenshot: {str(e)}")
            return ""
    
    def handle_alert(self, action: str = "accept") -> str:
        """
        Handle JavaScript alerts, confirms, and prompts.
        
        Args:
            action (str): Action to take ('accept', 'dismiss', 'text') (default: 'accept')
        
        Returns:
            str: Alert text if action is 'text', empty string otherwise
            
        Example:
            >>> alert_text = chrome.handle_alert("text")  # Get alert text
            >>> chrome.handle_alert("accept")             # Accept alert
        """
        try:
            if not self.driver:
                print("✗ Browser not started.")
                return ""
            
            alert = self.driver.switch_to.alert
            alert_text = alert.text
            
            if action == "accept":
                alert.accept()
                self._log_operation("Accepted alert")
            elif action == "dismiss":
                alert.dismiss()
                self._log_operation("Dismissed alert")
            elif action == "text":
                return alert_text
            
            return ""
            
        except Exception as e:
            print(f"✗ Error handling alert: {str(e)}")
            return ""
    
    def switch_to_frame(self, frame_selector: str = None, by: str = "css") -> bool:
        """
        Switch to an iframe or frame.
        
        Args:
            frame_selector (str, optional): Frame selector. If None, switches to default content
            by (str): Selection method (default: 'css')
        
        Returns:
            bool: True if switch successful, False otherwise
            
        Example:
            >>> chrome.switch_to_frame("iframe#content-frame")
            >>> # Do something in the frame
            >>> chrome.switch_to_frame()  # Switch back to main content
        """
        try:
            if not self.driver:
                print("✗ Browser not started.")
                return False
            
            if frame_selector is None:
                self.driver.switch_to.default_content()
                self._log_operation("Switched to default content")
            else:
                frame_element = self.find_element(frame_selector, by)
                if frame_element:
                    self.driver.switch_to.frame(frame_element)
                    self._log_operation(f"Switched to frame: {frame_selector}")
                else:
                    return False
            
            return True
            
        except Exception as e:
            print(f"✗ Error switching frame: {str(e)}")
            return False
    
    def execute_javascript(self, script: str, *args) -> Any:
        """
        Execute JavaScript code in the browser.
        
        Args:
            script (str): JavaScript code to execute
            *args: Arguments to pass to the script
        
        Returns:
            Any: Return value from JavaScript execution
            
        Example:
            >>> title = chrome.execute_javascript("return document.title;")
            >>> print(f"Page title: {title}")
            
            >>> chrome.execute_javascript("arguments[0].style.border = '2px solid red';", element)
        """
        try:
            if not self.driver:
                print("✗ Browser not started.")
                return None
            
            result = self.driver.execute_script(script, *args)
            self._log_operation("Executed JavaScript")
            return result
            
        except Exception as e:
            print(f"✗ Error executing JavaScript: {str(e)}")
            return None
    
    def get_page_source(self) -> str:
        """
        Get the HTML source of the current page.
        
        Returns:
            str: Page HTML source
            
        Example:
            >>> html = chrome.get_page_source()
            >>> if "welcome" in html.lower():
            ...     print("Welcome message found on page")
        """
        try:
            if not self.driver:
                print("✗ Browser not started.")
                return ""
            
            return self.driver.page_source
            
        except Exception as e:
            print(f"✗ Error getting page source: {str(e)}")
            return ""
    
    def get_current_url(self) -> str:
        """
        Get the current page URL.
        
        Returns:
            str: Current URL
            
        Example:
            >>> current_url = chrome.get_current_url()
            >>> print(f"Currently on: {current_url}")
        """
        try:
            if not self.driver:
                print("✗ Browser not started.")
                return ""
            
            return self.driver.current_url
            
        except Exception as e:
            print(f"✗ Error getting current URL: {str(e)}")
            return ""
    
    def refresh_page(self) -> bool:
        """
        Refresh the current page.
        
        Returns:
            bool: True if refresh successful, False otherwise
            
        Example:
            >>> chrome.refresh_page()
        """
        try:
            if not self.driver:
                print("✗ Browser not started.")
                return False
            
            self.driver.refresh()
            self._log_operation("Page refreshed")
            return True
            
        except Exception as e:
            print(f"✗ Error refreshing page: {str(e)}")
            return False
    
    def go_back(self) -> bool:
        """
        Navigate back in browser history.
        
        Returns:
            bool: True if navigation successful, False otherwise
            
        Example:
            >>> chrome.go_back()
        """
        try:
            if not self.driver:
                print("✗ Browser not started.")
                return False
            
            self.driver.back()
            self._log_operation("Navigated back")
            return True
            
        except Exception as e:
            print(f"✗ Error going back: {str(e)}")
            return False
    
    def go_forward(self) -> bool:
        """
        Navigate forward in browser history.
        
        Returns:
            bool: True if navigation successful, False otherwise
            
        Example:
            >>> chrome.go_forward()
        """
        try:
            if not self.driver:
                print("✗ Browser not started.")
                return False
            
            self.driver.forward()
            self._log_operation("Navigated forward")
            return True
            
        except Exception as e:
            print(f"✗ Error going forward: {str(e)}")
            return False
    
    def close_current_tab(self) -> bool:
        """
        Close the current browser tab.
        
        Returns:
            bool: True if close successful, False otherwise
            
        Example:
            >>> chrome.close_current_tab()
        """
        try:
            if not self.driver:
                print("✗ Browser not started.")
                return False
            
            self.driver.close()
            self._log_operation("Closed current tab")
            return True
            
        except Exception as e:
            print(f"✗ Error closing tab: {str(e)}")
            return False
    
    def close_browser(self) -> bool:
        """
        Close the entire browser and quit the WebDriver.
        
        Returns:
            bool: True if close successful, False otherwise
            
        Example:
            >>> chrome.close_browser()
        """
        try:
            if self.driver:
                self.driver.quit()
                self.driver = None
                self.wait = None
                self.actions = None
                self._log_operation("Browser closed")
                return True
            
            return False
            
        except Exception as e:
            print(f"✗ Error closing browser: {str(e)}")
            return False
    
    def get_operation_log(self) -> List[str]:
        """
        Get the complete log of all operations performed.
        
        Returns:
            List[str]: List of all operations with timestamps
            
        Example:
            >>> log = chrome.get_operation_log()
            >>> for entry in log:
            ...     print(entry)
        """
        return self.operation_log.copy()
    
    def clear_log(self):
        """
        Clear the operation log.
        
        Example:
            >>> chrome.clear_log()
        """
        self.operation_log = []
        print("✓ Operation log cleared")
    
    def _log_operation(self, operation: str):
        """
        Internal method to log operations with timestamps.
        
        Args:
            operation (str): Description of the operation performed
        """
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.operation_log.append(f"{timestamp} - {operation}")

# Example usage and demonstrations
if __name__ == "__main__":
    """
    Example demonstrating Chrome browser automation.
    """
    
    # Initialize Chrome handler
    chrome = ChromeHandler()
    
    try:
        # Start browser
        print("Starting Chrome browser...")
        if chrome.start_browser(headless=False, window_size=(1280, 720)):
            print("✓ Browser started successfully")
            
            # Navigate to Google
            print("\nNavigating to Google...")
            if chrome.navigate_to("https://www.google.com"):
                print("✓ Navigation successful")
                
                # Find search box and enter query
                print("\nSearching for 'Python programming'...")
                if chrome.send_keys_to_element("input[name='q']", "Python programming"):
                    print("✓ Search query entered")
                    
                    # Press Enter to search
                    search_box = chrome.find_element("input[name='q']")
                    if search_box:
                        search_box.send_keys(Keys.RETURN)
                        print("✓ Search submitted")
                        
                        # Wait for results and take screenshot
                        if chrome.wait_for_element("div#search", timeout=10):
                            print("✓ Search results loaded")
                            
                            screenshot_path = chrome.take_screenshot("google_search_results.png")
                            if screenshot_path:
                                print(f"✓ Screenshot saved: {screenshot_path}")
                
                # Wait a moment before closing
                time.sleep(2)
        
        # Show operation log
        print("\nOperation Log:")
        for entry in chrome.get_operation_log():
            print(f"  {entry}")
    
    finally:
        # Always close browser
        print("\nClosing browser...")
        chrome.close_browser()
        print("✓ Browser closed")