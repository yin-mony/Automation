@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
  set "PY=.venv\Scripts\python.exe"
) else if exist "venv\Scripts\python.exe" (
  set "PY=venv\Scripts\python.exe"
) else (
  set "PY=python"
)

echo 使用 Python: %PY%

echo [1/4] 安装/更新 PyInstaller ...
"%PY%" -m pip install --upgrade pip -q
"%PY%" -m pip install -r requirements-build.txt -q

echo [2/4] 打包「月度销售额汇总」（运行文件.py）...
"%PY%" -m PyInstaller -w -F --clean --name "月度销售额汇总" --distpath "dist" --workpath "build\summary" --specpath "build\spec" --hidden-import analysis --hidden-import openpyxl.cell._writer --collect-submodules openpyxl "运行文件.py"
if errorlevel 1 goto :fail

echo [3/4] 打包「月度销售额下载」（run.py）...
"%PY%" -m PyInstaller -w -F --clean --name "月度销售额下载" --distpath "dist" --workpath "build\download" --specpath "build\spec" --hidden-import main --hidden-import YidekeLogin --hidden-import cv2 --hidden-import pywinauto --hidden-import DrissionPage --collect-submodules DrissionPage run.py
if errorlevel 1 goto :fail

echo [4/4] 完成。
echo 产物: dist\月度销售额汇总.exe  dist\月度销售额下载.exe
pause
exit /b 0

:fail
echo 打包失败。
pause
exit /b 1
