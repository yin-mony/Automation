@echo off
chcp 65001 >nul
cd /d "%~dp0"

python -m pip install -r requirements.txt
if errorlevel 1 exit /b 1

python -m PyInstaller --noconfirm --windowed --onefile ^
  --name "按门店拆分表格" ^
  --collect-submodules openpyxl ^
  Tabellen_teilen.py

echo.
echo 生成的 exe 在 dist\按门店拆分表格.exe
pause
