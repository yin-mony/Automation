@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo === 安装/更新打包依赖 ===
python -m pip install -r requirements.txt -q
python -m pip install pyinstaller>=6.0.0 -q

echo === 开始打包单文件 exe ===
python -m PyInstaller --noconfirm --clean boss.spec

if %ERRORLEVEL% NEQ 0 (
    echo 打包失败
    exit /b 1
)

echo.
echo 打包完成: dist\BOSS直聘筛选简历.exe
pause
