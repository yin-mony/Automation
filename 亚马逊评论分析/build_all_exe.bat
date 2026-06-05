@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ===== 打包 1/2：亚马逊评论下载 (down_run.py) =====
call build_down_exe.bat
if errorlevel 1 exit /b 1

echo.
echo ===== 打包 2/2：亚马逊评论分析 (excel_run.py) =====
call build_excel_exe.bat
if errorlevel 1 exit /b 1

echo.
echo 全部打包完成。
echo   dist\亚马逊评论下载\亚马逊评论下载.exe
echo   dist\亚马逊评论分析\亚马逊评论分析.exe
pause
