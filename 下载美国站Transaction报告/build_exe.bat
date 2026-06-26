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

echo [1/2] 安装/更新 PyInstaller ...
"%PY%" -m pip install --upgrade pip -q
"%PY%" -m pip install "pyinstaller>=6.0" -q

echo [2/2] 打包「下载美国站Transaction报告」...
"%PY%" -m PyInstaller -w -F --clean ^
  --name "下载美国站Transaction报告" ^
  --distpath "dist" ^
  --workpath "build" ^
  --hidden-import main ^
  --hidden-import YidekeLogin ^
  --hidden-import pywinauto ^
  --hidden-import DrissionPage ^
  --hidden-import dateutil ^
  --hidden-import dateutil.relativedelta ^
  --collect-submodules DrissionPage ^
  run.py
if errorlevel 1 goto :fail

echo 完成。产物: dist\下载美国站Transaction报告.exe
pause
exit /b 0

:fail
echo 打包失败。
pause
exit /b 1
