@echo off
REM --- Configuration ---
set "VENV_DIR=D:\venv_automate"
set "PYTHON_EXE=python"

REM --- Windows Virtual Environment Pathing ---
REM The path to the pip executable inside the new venv location
set "PIP_EXE=%VENV_DIR%\Scripts\pip.exe"
REM The path to the Python executable inside the new venv location
set "PYTHON_VENV_EXE=%VENV_DIR%\Scripts\python.exe"

echo Checking for Python...
where %PYTHON_EXE% >nul 2>nul
if %errorlevel% neq 0 (
    echo Error: Python is not found in your system's PATH.
    echo Please ensure Python is installed and accessible via the 'python' command.
    goto :end
)

REM --- Create Virtual Environment in D:\venv_automate ---
if not exist "%VENV_DIR%" (
    echo Creating virtual environment at '%VENV_DIR%'...
    %PYTHON_EXE% -m venv "%VENV_DIR%"
    if exist "%VENV_DIR%" (
        echo Virtual environment created successfully.
    ) else (
        echo Error: Failed to create virtual environment.
        goto :end
    )
) else (
    echo Virtual environment '%VENV_DIR%' already exists.
)

REM --- Install Packages ---
echo Installing required packages (PyQt6, Pillow)...

if not exist "%PIP_EXE%" (
    echo Error: Could not find pip in the new environment. Check the environment creation.
    goto :end
)

REM Install required packages: PyQt6 and Pillow
"%PIP_EXE%" install PyQt6 Pillow

if %errorlevel% neq 0 (
    echo Error: Failed to install one or more packages.
    echo Please check the error messages above.
    goto :end
)

echo.
echo Setup complete. The environment is located at: %VENV_DIR%
echo.

:end
echo Press any key to exit...
pause >nul
