@echo off
REM Set the title of the console window
TITLE Folder Sync App Launcher

REM --- Configuration ---
REM Path to the embedded Python executable
SET PYTHON_EXECUTABLE=.\python-embed\python.exe

REM Name of the Python script to run
SET APP_SCRIPT=folder_sync_app.py

REM --- Execution ---
echo Launching the Folder Sync and Compare Tool...

REM Check if the Python executable exists
if not exist "%PYTHON_EXECUTABLE%" (
    echo ERROR: Python executable not found at "%PYTHON_EXECUTABLE%"
    echo Please ensure the 'python-embed' folder is in the correct location.
    pause
    exit /b 1
)

REM Check if the app script exists
if not exist "%APP_SCRIPT%" (
    echo ERROR: Application script not found: "%APP_SCRIPT%"
    pause
    exit /b 1
)

REM Run the Python application
"%PYTHON_EXECUTABLE%" "%APP_SCRIPT%"

echo Application closed.
@echo on
pause