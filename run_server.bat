@echo off
rem Runs the tool as a server. Needs config.json next to this file (see SERVER_SETUP.md).
cd /d "%~dp0"
if not exist config.json (
  echo config.json is missing. Copy config.example.json to config.json and fill in the paths.
  pause
  exit /b 1
)
python web.py
pause
