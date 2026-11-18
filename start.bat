@echo off
rem Double-click to start the Balance Confirmation tool, then open http://localhost:8080
cd /d "%~dp0"
if not exist config.json (
  echo config.json is missing. Copy config.example.json to config.json and fill in the paths.
  pause
  exit /b 1
)
start "" http://localhost:8080
python web.py
pause
