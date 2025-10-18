@echo off
setlocal enableDelayedExpansion

REM --- Configuration ---
set "VENV_DIR=%~dp0venv_automate"
set "PYTHON_EXE=python"
set "PIP_VENV_EXE=%VENV_DIR%\\Scripts\\pip.exe"
set "PYTHON_VENV_EXE=%VENV_DIR%\\Scripts\\python.exe"
set "DEP_SCRIPT=discover_and_check_deps.py"
set "MISSING_PACKAGE_LIST=temp_missing_packages.txt"
set "APP_SCRIPT=main_app.py"

echo.
echo --- Python Environment Setup and Application Runner ---

REM Check for base Python
where %PYTHON_EXE% >nul 2>nul
if %errorlevel% neq 0 (
    echo Error: Python is not found in your system's PATH.
    echo Please ensure Python is installed and accessible via the 'python' command.
    goto :cleanup
)

REM --- 1. Create Virtual Environment (if missing) ---
if not exist "%VENV_DIR%" (
    echo Creating virtual environment at '%VENV_DIR%'...
    %PYTHON_EXE% -m venv "%VENV_DIR%"
    if not exist "%VENV_DIR%" (
        echo Error: Failed to create virtual environment at %VENV_DIR%.
        goto :cleanup
    )
) else (
    echo Virtual environment '%VENV_DIR%' already exists.
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
    echo.
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
    echo     except Exception as e:
    echo         # print(f"Error checking installed packages: {e}", file=sys.stderr)
    echo         return set()
    echo.
    echo def discover_and_filter_packages():
    echo     # Standard library modules and built-ins to ignore
    echo     STANDARD_LIB = set(sys.builtin_module_names) ^| set(name for _, name, _ in pkgutil.iter_modules())
    echo     REQUIRED_PACKAGES = set()
    echo.
    echo     # Get installed packages from the specific VENV
    echo     VENV_PYTHON = os.path.join(r"%VENV_DIR%", "Scripts", "python.exe")
    echo     INSTALLED_PACKAGES = get_installed_packages(VENV_PYTHON)
    echo.
    echo     # Scan Python files in the current directory and subdirectories
    echo     os.chdir(os.path.dirname(os.path.abspath(__file__)))
    echo     for root, _, files in os.walk('.'):
    echo         for file in files:
    echo             if file.endswith('.py'):
    echo                 filepath = os.path.join(root, file)
    echo                 try:
    echo                     with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
    echo                         content = f.read()
    echo                     # Find 'import package' or 'from package import...'
    echo                     matches = re.findall(r'^(?:import|from)\s+([a-zA-Z0-9_.]+)(?:\s+as\s+)?', content, re.MULTILINE)
    echo.
    echo                     for match in matches:
    echo                         package_name = match.split('.')[0]
    echo.
    echo                         # Ignore standard library modules and potential local imports
    echo                         if package_name in STANDARD_LIB: continue
    echo                         if os.path.isdir(package_name) or os.path.exists(package_name + '.py'): continue
    echo.
    echo                         REQUIRED_PACKAGES.add(package_name)
    echo                 except Exception:
    echo                     pass
    echo.
    echo     # Special case mapping for known library names (e.g., PIL imports Pillow)
    echo     if 'PIL' in REQUIRED_PACKAGES:
    echo         REQUIRED_PACKAGES.remove('PIL')
    echo         REQUIRED_PACKAGES.add('Pillow')
    echo.
    echo     # Determine which packages are missing
    echo     MISSING = REQUIRED_PACKAGES - INSTALLED_PACKAGES
    echo     # A more robust check for packages that pip lists differently than their import name
    echo     # For example 'python-docx' is imported as 'docx'
    echo     # This is a simplified check; a more comprehensive solution would use a mapping library.
    echo     if 'cv2' in REQUIRED_PACKAGES and 'opencv-python' not in INSTALLED_PACKAGES:
    echo         MISSING.add('opencv-python')
    echo.
    echo     if MISSING:
    echo         print("Missing packages found:")
    echo         with open(r"%MISSING_PACKAGE_LIST%", 'w') as f:
    echo             for pkg in sorted(list(MISSING)):
    echo                 print(f"- {pkg}")
    echo                 f.write(pkg + '\n')
    echo     else:
    echo         print("All required packages are already installed in the virtual environment.")
    echo.
    echo if __name__ == "__main__":
    echo     discover_and_filter_packages()
) > "%DEP_SCRIPT%"

REM --- Run the Dependency Script ---
echo Running dependency scanner...
"%PYTHON_VENV_EXE%" "%DEP_SCRIPT%"

REM --- 3. Install Missing Dependencies ---
if exist "%MISSING_PACKAGE_LIST%" (
    echo.
    echo --- Installing Missing Dependencies ---
    "%PIP_VENV_EXE%" install -r "%MISSING_PACKAGE_LIST%"
    if %errorlevel% neq 0 (
        echo.
        echo Error: Failed to install one or more packages.
        echo Please check the output above for details.
        goto :cleanup
    )
    echo.
    echo --- All dependencies installed successfully. ---
)

REM --- 4. Run the main application ---
echo.
echo --- Running Python application: %APP_SCRIPT% ---
echo Using interpreter: %PYTHON_VENV_EXE%
echo.

"%PYTHON_VENV_EXE%" "%APP_SCRIPT%"

if %errorlevel% neq 0 (
    echo.
    echo ERROR: The application or the Python interpreter failed to run.
    echo Check the console for Python error messages.
)

:cleanup
echo.
echo --- Script finished. ---
if exist "%DEP_SCRIPT%" del "%DEP_SCRIPT%"
if exist "%MISSING_PACKAGE_LIST%" del "%MISSING_PACKAGE_LIST%"
pause
endlocal
