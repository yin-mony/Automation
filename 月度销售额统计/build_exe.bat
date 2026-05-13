@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist "venv\Scripts\python.exe" (
  set "PY=venv\Scripts\python.exe"
) else (
  set "PY=python"
)

echo 使用 Python: %PY%

echo [1/3] 安装/更新 PyInstaller ...
"%PY%" -m pip install --upgrade pip -q
"%PY%" -m pip install "pyinstaller>=6.0" -q

echo [2/3] 打包（无控制台窗口；含 DrissionPage 完整资源）...
"%PY%" -m PyInstaller --noconfirm --clean ^
  --windowed ^
  --name "月度销售额统计" ^
  --hidden-import main ^
  --hidden-import analysis ^
  --hidden-import YidekeLogin ^
  --hidden-import pandas ^
  --hidden-import openpyxl ^
  --hidden-import cv2 ^
  --hidden-import DrissionPage ^
  --hidden-import pywinauto ^
  --hidden-import psutil ^
  --hidden-import dateutil ^
  --collect-all DrissionPage ^
  run.py

if errorlevel 1 (
  echo 打包失败。
  pause
  exit /b 1
)

echo [3/3] 完成。
echo 请运行: dist\月度销售额统计\月度销售额统计.exe （连同 dist\月度销售额统计 文件夹一起分发）
pause
