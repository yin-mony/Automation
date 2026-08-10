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

echo [1/3] 安装/更新依赖与 PyInstaller ...
"%PY%" -m pip install --upgrade pip -q
"%PY%" -m pip install "pyinstaller>=6.0" DrissionPage pandas openpyxl PyQt5 -q

echo [2/3] 清理旧产物 ...
if exist "dist\1688商品链接采集.exe" del /q "dist\1688商品链接采集.exe"

echo [3/3] 打包「1688商品链接采集」...
"%PY%" -m PyInstaller --clean --noconfirm 1688商品链接采集.spec
if errorlevel 1 goto :fail

echo.
echo 完成。产物: dist\1688商品链接采集.exe
echo 说明: 目标机器需已安装 Chrome/Edge，DrissionPage 才能正常采集。
pause
exit /b 0

:fail
echo 打包失败。
pause
exit /b 1
