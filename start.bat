@echo off
rem Shoko Review UI launcher (ASCII only to avoid codepage issues)
cd /d "%~dp0"
python server.py --port 3471
if errorlevel 1 (
  echo.
  echo Failed to start. Try: pip install -r requirements.txt
  pause
)
