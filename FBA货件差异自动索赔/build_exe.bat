@echo off
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
  set "PY=.venv\Scripts\python.exe"
) else if exist "venv\Scripts\python.exe" (
  set "PY=venv\Scripts\python.exe"
) else (
  set "PY=python"
)

echo Using Python: %PY%

echo [1/4] Install project requirements ...
"%PY%" -m pip install --upgrade pip -q
"%PY%" -m pip install -r requirements.txt -q
if errorlevel 1 goto :fail

echo [2/4] Clean old dist files ...
if exist "dist" (
  for %%F in ("dist\*.exe") do del /q "%%~fF" 2>nul
  if exist "dist\run_config.json" del /q "dist\run_config.json"
) else (
  mkdir "dist"
)

echo [3/4] Build exe with PyInstaller ...
"%PY%" -c "import pathlib, subprocess, sys; spec=next(pathlib.Path('.').glob('*.spec')); sys.exit(subprocess.run([sys.executable, '-m', 'PyInstaller', '--clean', '--noconfirm', str(spec)]).returncode)"
if errorlevel 1 goto :fail

echo [4/4] Copy runtime config ...
if exist "run_config.json" (
  copy /y "run_config.json" "dist\run_config.json" >nul
  echo Copied run_config.json to dist.
) else (
  echo run_config.json not found, skip runtime config copy.
)

echo.
echo Build complete. Please distribute the whole dist folder.
pause
exit /b 0

:fail
echo Build failed.
pause
exit /b 1
