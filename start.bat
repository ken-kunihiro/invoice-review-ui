@echo off
rem Shoko Review UI launcher (ASCII only to avoid codepage issues)
setlocal
cd /d "%~dp0"

rem --- Pythonの検出 ---
set "PY_CMD="
where python >nul 2>nul
if not errorlevel 1 set "PY_CMD=python"
if not defined PY_CMD (
  where py >nul 2>nul
  if not errorlevel 1 set "PY_CMD=py -3"
)
if not defined PY_CMD (
  echo Python が見つかりません。
  echo https://www.python.org/downloads/ から Python 3.11 以上をインストールしてください。
  echo （インストール時に "Add python.exe to PATH" にチェックを入れてください）
  pause
  exit /b 1
)

rem --- 仮想環境の作成（初回のみ） ---
set "VENV_PY=.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
  echo 初回セットアップ中... しばらくお待ちください
  %PY_CMD% -m venv .venv
  if errorlevel 1 (
    echo 仮想環境の作成に失敗しました
    pause
    exit /b 1
  )
)

rem --- 依存パッケージのインストール（requirements.txtが変わった時だけ） ---
set "NEED_INSTALL=0"
if not exist ".venv\.requirements.snapshot" (
  set "NEED_INSTALL=1"
) else (
  fc /b requirements.txt ".venv\.requirements.snapshot" >nul 2>nul
  if errorlevel 1 set "NEED_INSTALL=1"
)

if "%NEED_INSTALL%"=="1" (
  echo 依存パッケージをインストール/更新しています...
  "%VENV_PY%" -m pip install --upgrade pip >nul
  "%VENV_PY%" -m pip install -r requirements.txt
  if errorlevel 1 (
    echo パッケージのインストールに失敗しました
    pause
    exit /b 1
  )
  copy /y requirements.txt ".venv\.requirements.snapshot" >nul
)

rem --- デスクトップショートカットの作成（初回のみ） ---
if not exist ".venv\.shortcut_created" (
  powershell -NoProfile -Command ^
    "$ws = New-Object -ComObject WScript.Shell;" ^
    "$desktop = [Environment]::GetFolderPath('Desktop');" ^
    "$lnk = $ws.CreateShortcut((Join-Path $desktop '証憑レビューUI.lnk'));" ^
    "$lnk.TargetPath = '%~f0';" ^
    "$lnk.WorkingDirectory = '%~dp0';" ^
    "$lnk.IconLocation = '%SystemRoot%\System32\shell32.dll,13';" ^
    "$lnk.Save()" >nul 2>nul
  echo. > ".venv\.shortcut_created"
)

"%VENV_PY%" server.py --port 3471
if errorlevel 1 (
  echo.
  echo 起動に失敗しました
  pause
)
