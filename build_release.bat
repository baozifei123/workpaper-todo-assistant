@echo off
chcp 65001 >nul
cd /d "%~dp0"
call build_exe.bat
if errorlevel 1 exit /b 1
powershell -NoProfile -ExecutionPolicy Bypass -File "tools\build_release.ps1" -Version 0.3.0
if errorlevel 1 exit /b 1
echo Release complete.
