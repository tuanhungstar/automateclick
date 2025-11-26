import keyboard
import pyautogui
from my_lib.shared_context import ExecutionContext as Context
import Gui_Automate
import Mouse_Key_Automate
import time
class Keyboard:
    def __init__(self,context: Context):
        self.context = Context
        pass
    def send_text(self,text,delay=0):
        keyboard.write(str(text),delay)
        return f"Bot write this text: {text}"

    def send_hotkey(self,hotkey):
        """
        send hotkey link "CTRL+a"
        """
        time.sleep(1)
        keyboard.press_and_release(hotkey)
        return f"Bot send this hotkey: {hotkey}"
        

    def moveMouse(self,x=50,y=50):
        ''' 


        move mouse to deserd position
        (0,0)
          +------------> X-Axis
          |
          |
          |
          |
          |
          |
          |
          v
         Y-Axis
        '''
        pyautogui.moveTo(x,y)
        pass
        return
