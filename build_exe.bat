@echo off
rem Run this ONCE on any Windows PC that has Python. It makes dist\BalanceCheck\ : a folder that needs NO Python install.
cd /d "%~dp0"
python -m pip install --quiet -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --clean --onedir --name BalanceCheck ^
  --add-data "templates;templates" --add-data "config.example.json;." --add-data "sample;sample" ^
  --collect-all pypdfium2 --collect-all openpyxl --hidden-import win32com.client --hidden-import pythoncom ^
  --hidden-import setup_wizard --hidden-import waitress owner_main.py
echo.
echo Done. Give the owner the folder  dist\BalanceCheck  (zip it). He double-clicks BalanceCheck.exe.
pause
