@echo off
rem Shoko Review UI launcher (ASCII only - do not add non-ASCII characters to this file.
rem Windows cmd.exe parses .bat files using the console codepage, which can be Shift-JIS
rem on Japanese Windows. Mixing that with a UTF-8-saved file corrupts multi-byte text and
rem can break command parsing on lines far below the corrupted one. Keep all comments,
rem echo messages and literals in this file plain ASCII. See CHANGELOG v2026-08-21.2/.3.)
setlocal
chcp 65001 >nul
cd /d "%~dp0"

rem --- Detect Python ---
set "PY_CMD="
where python >nul 2>nul
if not errorlevel 1 set "PY_CMD=python"
if not defined PY_CMD (
  where py >nul 2>nul
  if not errorlevel 1 set "PY_CMD=py -3"
)
if not defined PY_CMD (
  echo Python was not found.
  echo Please install Python 3.11 or later from https://www.python.org/downloads/
  echo During installation, check "Add python.exe to PATH".
  pause
  exit /b 1
)

rem --- Create virtual environment (first run only) ---
set "VENV_PY=.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
  echo First-time setup, please wait...
  %PY_CMD% -m venv .venv
  if errorlevel 1 (
    echo Failed to create the virtual environment.
    pause
    exit /b 1
  )
)

rem --- Install dependencies, but only when requirements.txt has changed ---
set "NEED_INSTALL=0"
if not exist ".venv\.requirements.snapshot" (
  set "NEED_INSTALL=1"
) else (
  fc /b requirements.txt ".venv\.requirements.snapshot" >nul 2>nul
  if errorlevel 1 set "NEED_INSTALL=1"
)

if "%NEED_INSTALL%"=="1" (
  echo Installing/updating dependencies...
  "%VENV_PY%" -m pip install --upgrade pip >nul
  "%VENV_PY%" -m pip install -r requirements.txt
  if errorlevel 1 (
    echo Failed to install dependencies.
    pause
    exit /b 1
  )
  copy /y requirements.txt ".venv\.requirements.snapshot" >nul
)

rem --- Create a desktop shortcut (first run only) ---
rem Kept as a single line (no caret continuations) and the label is built from Unicode
rem code points so this file never contains a raw non-ASCII byte.
if not exist ".venv\.shortcut_created" (
  powershell -NoProfile -Command "$n=[char]0x8A3C+[char]0x6191+[char]0x30EC+[char]0x30D3+[char]0x30E5+[char]0x30FC+'UI'; $ws=New-Object -ComObject WScript.Shell; $d=[Environment]::GetFolderPath('Desktop'); $lnk=$ws.CreateShortcut((Join-Path $d ($n+'.lnk'))); $lnk.TargetPath='%~f0'; $lnk.WorkingDirectory='%~dp0'; $lnk.IconLocation='%SystemRoot%\System32\shell32.dll,13'; $lnk.Save()" >nul 2>nul
  echo. > ".venv\.shortcut_created"
)

"%VENV_PY%" server.py --port 3471
if errorlevel 1 (
  echo.
  echo Failed to start.
  pause
)
