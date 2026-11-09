@echo off
setlocal
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" app.py
) else (
  py -3.15 app.py
)

if errorlevel 1 (
  echo.
  echo The app could not start. Run setup.bat first.
  pause
)
