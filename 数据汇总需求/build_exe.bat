@echo off
chcp 65001 >nul
cd /d "%~dp0"

python -m pip install -r requirements.txt
if errorlevel 1 exit /b 1

python -m PyInstaller --noconfirm --windowed --onefile ^
  --name "数据汇总处理工具" ^
  --collect-submodules openpyxl ^
  run.py

echo.
echo 生成的 exe 在 dist\数据汇总处理工具.exe
pause
