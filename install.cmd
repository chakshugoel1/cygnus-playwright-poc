@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
if errorlevel 1 (
  echo Installer failed.
  exit /b %errorlevel%
)
echo Installer finished.
