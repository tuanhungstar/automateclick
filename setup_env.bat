@echo off
set /p ENV_PATH="Enter the full path for the virtual environment (e.g., C:\MyEnvs\ProjectEnv): "

REM If no path is entered, default to a 'venv' folder in the current directory
if "%ENV_PATH%"=="" set "ENV_PATH=venv"

echo Creating virtual environment folder at: %ENV_PATH%
if not exist "%ENV_PATH%" (
    mkdir "%ENV_PATH%"
)

echo Creating virtual environment...
python -m venv "%ENV_PATH%"

echo Activating the virtual environment and installing requirements...
call "%ENV_PATH%\Scripts\activate.bat" && pip install -r requirements.txt

echo.
echo Setup complete.
pause
