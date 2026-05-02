@echo off
chcp 65001 >nul
cd /d "%~dp0"

python -m pip install -r requirements.txt
if errorlevel 1 (
    echo 依赖安装失败，请确认已安装 Python 并已加入 PATH。
    pause
    exit /b 1
)

python run.py --mode gui
if errorlevel 1 pause
