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
"%PY%" -m pip install "pyinstaller>=6.0" requests -q

echo [2/3] 清理旧产物 ...
if exist "dist\日报生成器.exe" del /q "dist\日报生成器.exe"

echo [3/3] 打包「日报生成器」...
"%PY%" -m PyInstaller --clean --noconfirm 日报生成器.spec
if errorlevel 1 goto :fail

echo.
echo 完成。产物: dist\日报生成器.exe
pause
exit /b 0

:fail
echo 打包失败。
pause
exit /b 1
