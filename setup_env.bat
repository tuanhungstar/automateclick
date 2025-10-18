@echo off
REM --- Check for Python and Virtual Environment Setup ---

set "VENV_NAME=venv_automate"
set "PYTHON_EXE=python"
set "PIP_EXE=%VENV_NAME%\Scripts\pip"

echo Checking for Python...
where %PYTHON_EXE% >nul 2>nul
if %errorlevel% neq 0 (
    echo Error: Python is not found in your system's PATH.
    echo Please ensure Python is installed and accessible via the 'python' command.
    goto :end
)

REM --- Create Virtual Environment ---
if not exist %VENV_NAME% (
    echo Creating virtual environment '%VENV_NAME%'...
    %PYTHON_EXE% -m venv %VENV_NAME%
    if exist %VENV_NAME% (
        echo Virtual environment created successfully.
    ) else (
        echo Error: Failed to create virtual environment.
        goto :end
    )
) else (
    echo Virtual environment '%VENV_NAME%' already exists.
)

REM --- Activate Environment and Install Packages ---
echo Activating virtual environment and installing packages...

REM Use the pip executable directly from the new environment
if not exist "%PIP_EXE%.exe" (
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
echo Setup complete. The following packages were installed:
echo - PyQt6
echo - Pillow
echo.

:end
echo Press any key to exit...
pause >nul
