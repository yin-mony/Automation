@echo off
chcp 65001 >nul
cd /d "%~dp0"
pyinstaller --noconfirm --clean build.spec
if %ERRORLEVEL%==0 (
    echo.
    echo 打包完成: dist\赛狐ERP自动化.exe
) else (
    echo 打包失败，错误码 %ERRORLEVEL%
)
pause
