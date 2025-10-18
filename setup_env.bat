@echo off
setlocal enableDelayedExpansion

REM --- Configuration ---
REM VENV_DIR is set to a subdirectory 'venv_automate' in the same folder as this BAT file.
set "VENV_NAME=venv_automate"
set "VENV_DIR=%~dp0%VENV_NAME%"
set "PYTHON_EXE=python"
set "PIP_VENV_EXE=%VENV_DIR%\Scripts\pip.exe"
set "PYTHON_VENV_EXE=%VENV_DIR%\Scripts\python.exe"
set "DEP_SCRIPT=discover_and_check_deps.py"
set "MISSING_PACKAGE_LIST=temp_missing_packages.txt"

echo.
echo --- Python Environment Setup and Update ---

REM Check for base Python
where %PYTHON_EXE% >nul 2>nul
if %errorlevel% neq 0 (
    echo Error: Python is not found in your system's PATH.
    echo Please ensure Python is installed and accessible via the 'python' command.
    goto :cleanup
)

REM --- 1. Cleanup and Create Virtual Environment ---

if exist "%VENV_DIR%" (
    echo Existing virtual environment found at '%VENV_DIR%'.
    echo Deleting existing environment before creating a new one...
    REM /S: Removes all directories and files in addition to the directory itself.
    REM /Q: Quiet mode, meaning it will delete without asking for confirmation.
    rmdir /s /q "%VENV_DIR%"
    if exist "%VENV_DIR%" (
        echo Error: Failed to delete existing environment. Please delete manually and try again.
        goto :cleanup
    )
)

echo Creating virtual environment at '%VENV_DIR%'...
%PYTHON_EXE% -m venv "%VENV_DIR%"
if not exist "%VENV_DIR%" (
    echo Error: Failed to create virtual environment.
    goto :cleanup
)

REM Check for VENV's pip
if not exist "%PIP_VENV_EXE%" (
    echo Error: Could not find pip in the new environment at %VENV_DIR%.
    goto :cleanup
)

echo.
echo --- 2. Scanning & Checking for Missing Dependencies ---

REM --- Create Dependency Discovery and Checker Script ---
echo Creating temporary dependency scanner script...
(
    echo import os
    echo import re
    echo import sys
    echo import pkgutil
    echo import subprocess
    
    echo def get_installed_packages(python_executable):
    echo     """Uses pip to get a set of currently installed packages."""
    echo     try:
    echo         # Use the virtual environment's pip to list installed packages
    echo         result = subprocess.run([python_executable, '-m', 'pip', 'freeze'], capture_output=True, text=True, check=True)
    echo         # Extract package names (ignoring versions and local paths)
    echo         installed = set()
    echo         for line in result.stdout.splitlines():
    echo             line = line.strip()
    echo             if line and not line.startswith('-e'): # Skip editable installs
    echo                 if '==' in line:
    echo                     installed.add(line.split('==')[0].strip().lower())
    echo                 elif '@' in line: # Handle packages installed from URLs
    echo                     installed.add(line.split('@')[0].strip().lower())
    echo                 else: # Simple case
    echo                     installed.add(line.lower())
    echo         return installed
    echo     except Exception:
    echo         return set()
    
    echo def discover_and_filter_packages():
    echo     # Standard library modules and built-ins to ignore
    echo     STANDARD_LIB = set(sys.builtin_module_names) | set(name for _, name, _ in pkgutil.iter_modules())
    echo     REQUIRED_PACKAGES = set()
    echo     
    echo     # --- IMPORTANT: Get VENV_DIR from the BAT environment variable ---
    echo     VENV_DIR_PATH = r"%VENV_DIR%" 
    echo     VENV_PYTHON = os.path.join(VENV_DIR_PATH, "Scripts", "python.exe")
    echo     INSTALLED_PACKAGES = get_installed_packages(VENV_PYTHON)
    echo     
    echo     # Scan Python files in the current directory and subdirectories
    echo     # The current directory is where the BAT file is run from.
    echo     for root, _, files in os.walk('.'):
    echo         for file in files:
    echo             if file.endswith('.py'):
    echo                 filepath = os.path.join(root, file)
    echo                 try:
    echo                     with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
    echo                         content = f.read()
    echo                     # Find 'import package' or 'from package import...'
    echo                     matches = re.findall(r'^(?:import|from)\s+([a-zA-Z0-9_.]+)(?:\s+as\s+)?', content, re.MULTILINE)
    echo                     
    echo                     for match in matches:
    echo                         package_name = match.split('.')[0]
    echo                         
    echo                         # Ignore standard library modules and potential local imports
    echo                         if package_name in STANDARD_LIB: continue
    echo                         if os.path.isdir(package_name) or os.path.exists(package_name + '.py'): continue
    echo                         
    echo                         REQUIRED_PACKAGES.add(package_name)
    echo                 except Exception:
    echo                     pass

    echo # Special case mapping for known library names (e.g., PIL imports Pillow)
    echo     if 'PIL' in REQUIRED_PACKAGES:
    echo         REQUIRED_PACKAGES.remove('PIL')
    echo         REQUIRED_PACKAGES.add('Pillow')
    
    echo     # Determine which packages are missing
    echo     MISSING_PACKAGES = set()
    echo     for pkg in REQUIRED_PACKAGES:
    echo         # Check if the required package (converted to lowercase for robust check) is installed
    echo         if pkg.lower() not in INSTALLED_PACKAGES:
    echo             MISSING_PACKAGES.add(pkg)

    echo     # Write the space-separated list of missing packages to a file
    echo     with open(r"%MISSING_PACKAGE_LIST%", 'w') as f:
    echo         f.write(' '.join(MISSING_PACKAGES))

    echo if __name__ == '__main__':
    echo     discover_and_filter_packages()

) > %DEP_SCRIPT%

REM Execute the script using the new VENV's Python
"%PYTHON_VENV_EXE%" %DEP_SCRIPT%

if %errorlevel% neq 0 (
    echo Error: Dependency discovery script failed to execute.
    goto :cleanup
)

REM --- 3. Install Missing Packages ---
echo Reading required packages from %MISSING_PACKAGE_LIST%...
set /p PACKAGES_TO_INSTALL=< %MISSING_PACKAGE_LIST%

if "%PACKAGES_TO_INSTALL%"=="" (
    echo.
    echo All required third-party packages are already installed. Nothing to do.
) else (
    echo Found packages to install: !PACKAGES_TO_INSTALL!
    echo.
    echo Installing packages...
    "%PIP_VENV_EXE%" install !PACKAGES_TO_INSTALL!

    if %errorlevel% neq 0 (
        echo.
        echo ERROR: Failed to install one or more packages.
    ) else (
        echo.
        echo All packages installed successfully!
    )
)

:cleanup
echo.
echo --- 4. Cleanup ---
REM Delete temporary files
if exist "%DEP_SCRIPT%" del %DEP_SCRIPT%
if exist "%MISSING_PACKAGE_LIST%" del %MISSING_PACKAGE_LIST%
echo Cleanup finished.

:final_exit
echo.
echo --- SCRIPT EXECUTION FINISHED ---
echo The environment is located at: %VENV_DIR%
echo Press any key to close this window...
pause >nul
endlocal
