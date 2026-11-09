@echo off
setlocal
cd /d "%~dp0"

py -3.15 -c "import sys; raise SystemExit(sys.version_info[:3] != (3, 15, 2))" >nul 2>&1
if errorlevel 1 (
  echo Python 3.15.2 is required but was not found.
  echo Install Python 3.15.2 from https://www.python.org/downloads/windows/
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  py -3.15 -m venv .venv
  if errorlevel 1 goto failed
)

".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto failed
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed

echo Setup complete. Double-click run_app.bat to launch the application.
pause
exit /b 0

:failed
echo Setup failed. Check the Python installation and your internet connection.
pause
exit /b 1
