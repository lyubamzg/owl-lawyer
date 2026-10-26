@echo off
setlocal
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0packaging\windows\install-from-source.ps1"
if errorlevel 1 (
  echo.
  echo Installation failed. See "%LOCALAPPDATA%\OwlLawyer\Logs\setup.log"
  pause
  exit /b 1
)
pause
