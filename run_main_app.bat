@echo off
setlocal

:: --- Configuration ---
set "APP_FOLDER=App content"

:: --- 1. Find the application folder relative to this script ---
set "APP_PATH=%~dp0%APP_FOLDER%"

:: --- 2. Check if the application folder exists ---
if not exist "%APP_PATH%" (
    echo [ERROR] The application folder '%APP_FOLDER%' was not found.
    echo Please make sure it is in the same directory as this script.
    pause
    goto :eof
)

:: --- 3. Change the working directory to the application folder ---
:: This is the most important step for Python's imports to work correctly.
cd /d "%APP_PATH%"

:: --- 4. Define the Python executable path (now inside "App content") ---
set "PYTHON_EXE=python-embed\python.exe"

:: --- 5. Check if the local Python installation exists ---
if not exist "%PYTHON_EXE%" (
    echo [ERROR] The local Python installation was not found in '%APP_FOLDER%'.
    echo Please run the setup scripts inside the '%APP_FOLDER%' folder first.
    pause
    goto :eof
)

:: --- 6. Run the application ---
echo [INFO] Starting AutomateTask...
"%PYTHON_EXE%" runner.py

:eof
echo.
echo [INFO] Application has finished.
pause