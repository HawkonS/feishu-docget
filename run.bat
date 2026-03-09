@echo off
setlocal

echo [INFO] Checking environment...

:: 1. Check if Python is installed
python --version >nul 2>&1
if %errorlevel% equ 0 goto :FOUND_PYTHON

echo [WARN] Python is not installed or not in PATH.
echo.
echo Do you want to automatically install Python 3.10?
set /p "CHOICE=(Y/N): "
if /i "%CHOICE%" neq "Y" (
    echo.
    echo Please install Python manually from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo.
echo [INFO] Downloading Python 3.10 installer...
set "INSTALLER_URL=https://www.python.org/ftp/python/3.10.11/python-3.10.11-amd64.exe"
set "INSTALLER_PATH=%TEMP%\python-installer.exe"

:: Download using PowerShell
powershell -Command "Invoke-WebRequest -Uri '%INSTALLER_URL%' -OutFile '%INSTALLER_PATH%'"
if not exist "%INSTALLER_PATH%" (
    echo [ERROR] Failed to download Python installer.
    pause
    exit /b 1
)

echo [INFO] Installing Python... (This may take a few minutes)
echo [INFO] Please accept the User Account Control (UAC) prompt if it appears.

:: Install Python silently and add to PATH
"%INSTALLER_PATH%" /passive PrependPath=1 Include_test=0

:: Clean up installer
if exist "%INSTALLER_PATH%" del "%INSTALLER_PATH%"

echo.
echo [INFO] Python installation completed.
echo [IMPORTANT] Please CLOSE this window and RESTART the script to apply changes.
pause
exit /b 0

:FOUND_PYTHON
echo [INFO] Python found.

:: 2. Install dependencies
echo [INFO] Installing dependencies...
set REQUIRED_PACKAGES=Flask requests python-docx lxml Pillow

pip install %REQUIRED_PACKAGES%
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to install dependencies.
    echo Please check your network connection or try running as Administrator.
    pause
    exit /b 1
)

:: 3. Start Application
echo.
echo [INFO] Starting feishu_docget service...

set PYTHONPATH=%CD%

echo [INFO] Loading configuration...
python -c "from src.core.config_loader import ConfigLoader; ConfigLoader.load_config()"

echo [INFO] Service started. Keep this window open.
python src/app.py

if %errorlevel% neq 0 (
    echo [ERROR] Service exited with error.
    pause
)
