@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist "venv\Scripts\python.exe" (
  set "PY=venv\Scripts\python.exe"
) else (
  set "PY=python"
)

echo 使用 Python: %PY%

echo [1/3] 安装/更新依赖与 PyInstaller ...
"%PY%" -m pip install --upgrade pip -q
"%PY%" -m pip install -r requirements.txt -q
"%PY%" -m pip install "pyinstaller>=6.0" -q

echo [2/3] 打包「亚马逊评论下载」down_run.py（无控制台窗口）...
"%PY%" -m PyInstaller --noconfirm --clean ^
  --windowed ^
  --name "亚马逊评论下载" ^
  --hidden-import main ^
  --hidden-import YidekeLogin ^
  --hidden-import pandas ^
  --hidden-import openpyxl ^
  --hidden-import cv2 ^
  --hidden-import DrissionPage ^
  --hidden-import pywinauto ^
  --hidden-import pywinauto.application ^
  --hidden-import comtypes ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --hidden-import psutil ^
  --hidden-import dateutil ^
  --collect-all DrissionPage ^
  --collect-all pywinauto ^
  down_run.py

if errorlevel 1 (
  echo 打包失败。
  pause
  exit /b 1
)

echo [3/3] 完成。
echo 可执行文件: dist\亚马逊评论下载\亚马逊评论下载.exe
echo 分发时请打包整个 dist\亚马逊评论下载 文件夹。
powershell -NoProfile -Command "Compress-Archive -Path 'dist\亚马逊评论下载' -DestinationPath 'dist\亚马逊评论下载.zip' -Force" 2>nul
pause
