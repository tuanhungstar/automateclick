@echo off
REM --- Configuration ---
set "PYTHON_EXE=D:\venv_automate\Scripts\python.exe"
set "APP_SCRIPT=AutomateTask.py"

echo Running Python application: %APP_SCRIPT%
echo Using interpreter: %PYTHON_EXE%
echo.

REM --- Execute the application using the specified virtual environment python.exe ---
"%PYTHON_EXE%" "%APP_SCRIPT%"

if %errorlevel% neq 0 (
    echo.
    echo ERROR: The application or the Python interpreter failed to run.
    echo Check the console for Python error messages.
)

echo.
echo Script finished.
pause
