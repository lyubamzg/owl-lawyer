@echo off
setlocal
set "PYTHONUTF8=1"
if not exist "%~dp0runtime\Scripts\pythonw.exe" (
  echo SOVA runtime is missing. Please run the installer again.
  pause
  exit /b 1
)
start "" "%~dp0runtime\Scripts\pythonw.exe" "%~dp0app\desktop_app.py"
