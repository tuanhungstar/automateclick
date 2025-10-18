@echo off
setlocal enableDelayedExpansion

REM --- Configuration ---
set "VENV_DIR=D:\venv_automate"
set "PYTHON_EXE=python"
set "PIP_VENV_EXE=%VENV_DIR%\Scripts\pip.exe"
set "PYTHON_VENV_EXE=%VENV_DIR%\Scripts\python.exe"
set "DEP_SCRIPT=discover_deps.py"
set "PACKAGE_LIST=temp_packages.txt"

echo.
echo --- Python Environment Setup ---

REM Check for base Python
where %PYTHON_EXE% >nul 2>nul
if %errorlevel% neq 0 (
    echo Error: Python is not found in your system's PATH.
    goto :end
)

REM --- 1. Create Virtual Environment ---
if not exist "%VENV_DIR%" (
    echo Creating virtual environment at '%VENV_DIR%'...
    %PYTHON_EXE% -m venv "%VENV_DIR%"
    if not exist "%VENV_DIR%" (
        echo Error: Failed to create virtual environment.
        goto :end
    )
) else (
    echo Virtual environment '%VENV_DIR%' already exists. Skipping creation.
)

REM Check for VENV's pip
if not exist "%PIP_VENV_EXE%" (
    echo Error: Could not find pip in the new environment at %VENV_DIR%.
    goto :end
)

echo.
echo --- 2. Scanning for Dependencies ---

REM --- Create Dependency Discovery Script (discover_deps.py) ---
echo Creating temporary dependency scanner script...
(
    echo import os
    echo import re
    echo import sys
    echo import pkgutil
    echo STANDARD_LIB = set(sys.builtin_module_names) | set(name for _, name, _ in pkgutil.iter_modules())
    echo THIRD_PARTY_PACKAGES = set()
    echo os.chdir(os.path.dirname(os.path.abspath(__file__)))
    echo for root, _, files in os.walk('.'):
    echo     for file in files:
    echo         if file.endswith('.py'):
    echo             filepath = os.path.join(root, file)
    echo             try:
    echo                 with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
    echo                     content = f.read()
    echo                 
    echo                 # Find 'import package' or 'from package import...'
    echo                 matches = re.findall(r'^(?:import|from)\s+([a-zA-Z0-9_.]+)(?:\s+as\s+)?', content, re.MULTILINE)
    echo                 
    echo                 for match in matches:
    echo                     package_name = match.split('.')[0]
    echo                     
    echo                     # Filter out standard library modules and local modules
    echo                     if package_name in STANDARD_LIB:
    echo                         continue
    echo                     
    echo                     # Simple heuristic: Skip package names that are likely local files/folders (e.g., 'my_lib')
    echo                     # For absolute robustness, a pyproject.toml or setup.py is needed, but this is the best for a simple scanner.
    echo                     if os.path.isdir(package_name) or os.path.exists(package_name + '.py'):
    echo                         continue
    echo                     
    echo                     THIRD_PARTY_PACKAGES.add(package_name)

    echo             except Exception as e:
    echo                 # print(f"Error processing {filepath}: {e}")
    echo                 pass

    echo # Special case mapping for known library names
    echo if 'PIL' in THIRD_PARTY_PACKAGES:
    echo     THIRD_PARTY_PACKAGES.remove('PIL')
    echo     THIRD_PARTY_PACKAGES.add('Pillow')
    
    echo # Write the space-separated list of packages to a file
    echo with open('%PACKAGE_LIST%', 'w') as f:
    echo     f.write(' '.join(THIRD_PARTY_PACKAGES))
) > %DEP_SCRIPT%

REM Execute the script using the new VENV's Python
"%PYTHON_VENV_EXE%" %DEP_SCRIPT%

if %errorlevel% neq 0 (
    echo Error: Dependency discovery script failed to execute.
    goto :cleanup
)

REM --- Read Packages from Output File ---
echo Reading required packages from %PACKAGE_LIST%...
set /p PACKAGES_TO_INSTALL=< %PACKAGE_LIST%

if "%PACKAGES_TO_INSTALL%"=="" (
    echo No third-party packages were found.
    goto :cleanup
)

echo Found packages: !PACKAGES_TO_INSTALL!
echo.
echo --- 3. Installing Packages ---
"%PIP_VENV_EXE%" install !PACKAGES_TO_INSTALL!

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Failed to install all packages. Check your internet connection and package names.
) else (
    echo.
    echo All required packages installed successfully!
)

:cleanup
echo.
echo --- 4. Cleanup ---
del %DEP_SCRIPT%
del %PACKAGE_LIST%
echo Cleanup finished.

:end
echo.
echo Setup script complete. The environment is located at: %VENV_DIR%
pause
endlocal
