@echo off
cd /d "%~dp0"
echo.
echo === Balance check: setup ===
where python >nul 2>nul
if errorlevel 1 (
  echo Python is not installed. Trying to install it...
  winget install -e --id Python.Python.3.11 --accept-source-agreements --accept-package-agreements
  echo.
  echo Close this window and double-click setup.bat again. If Python still is not found, ask IT to install Python 3.11.
  pause
  exit /b 1
)
python -m pip install --quiet -r requirements.txt
if errorlevel 1 (
  echo.
  echo Could not install the needed parts. Send a photo of this window to the person who set this up.
  pause
  exit /b 1
)
python setup_wizard.py
echo.
pause
