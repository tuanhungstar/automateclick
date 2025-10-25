import sys
import os
# Import the main application module
# This gives you access to the classes defined inside main_app.py

# --- START OF FIX ---
# Get the absolute path of the directory containing this script.
# This ensures that Python knows where to look for other files like main_app.py.
script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(script_dir)
# --- END OF FIX ---
from PyQt6.QtWidgets import QApplication
import main_app



# The logic below replicates the original entry point from main_app.py's 
# 'if __name__ == "__main__":' block.

def run_gui():
    """Initializes and runs the PyQt6 GUI."""
    
    # Access the QApplication and MainWindow classes from the imported module
    try:
        QApplication = main_app.QApplication 
        MainWindow = main_app.MainWindow
    except AttributeError:
        print("Error: Could not access QApplication or MainWindow from main_app.py.")
        print("Ensure main_app.py is importing PyQt6.QtWidgets.QApplication at the top level.")
        sys.exit(1)

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    run_gui()
