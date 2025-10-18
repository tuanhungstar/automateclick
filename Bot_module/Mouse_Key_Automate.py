import keyboard

class Keyboard:
    def __init__(self):

        pass
    def send_text(self,text):
        keyboard.write(str(text))
        return f"Bot write this text: {text}"

    def send_hotkey(self,hotkey):
        """
        send hotkey link "CTRL+a"
        """
        keyboard.press_and_release(hotkey)
        return f"Bot send this hotkey: {hotkey}"
