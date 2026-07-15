@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist "venv\Scripts\python.exe" (
  set "PY=venv\Scripts\python.exe"
) else (
  set "PY=python"
)

set "APP_NAME=亚马逊评论工具"

echo Using Python: %PY%
"%PY%" -m pip install --upgrade pip -q
if errorlevel 1 (
  echo Install pip failed.
  pause
  exit /b 1
)

"%PY%" -m pip install -r requirements.txt -q
if errorlevel 1 (
  echo Install requirements failed.
  pause
  exit /b 1
)

"%PY%" -m pip install "pyinstaller>=6.0" -q
if errorlevel 1 (
  echo Install PyInstaller failed.
  pause
  exit /b 1
)

"%PY%" -m PyInstaller --noconfirm --clean --windowed --name "%APP_NAME%" --hidden-import run --hidden-import main --hidden-import analysis --hidden-import YidekeLogin --hidden-import pandas --hidden-import openpyxl --hidden-import requests --hidden-import websocket --hidden-import DrissionPage --hidden-import pywinauto --hidden-import pywinauto.application --hidden-import comtypes --hidden-import pythoncom --hidden-import pywintypes --hidden-import psutil --collect-all DrissionPage --collect-all pywinauto run.py
if errorlevel 1 (
  echo Build failed.
  pause
  exit /b 1
)

if exist "%APP_NAME%.spec" (
  del /q "%APP_NAME%.spec"
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$dist=Join-Path (Get-Location) 'dist'; $src=Join-Path $dist '%APP_NAME%'; $zip=Join-Path $dist ('%APP_NAME%' + '.zip'); if(Test-Path -LiteralPath $zip){Remove-Item -LiteralPath $zip -Force}; Compress-Archive -LiteralPath $src -DestinationPath $zip -Force"
if errorlevel 1 (
  echo Zip failed.
  pause
  exit /b 1
)

echo Build finished.
pause
