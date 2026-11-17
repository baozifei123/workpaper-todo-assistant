@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -m PyInstaller --noconfirm --clean --windowed ^
  --name "底稿未完成助手" ^
  --icon "assets\app.ico" ^
  --add-data "assets\app.ico;assets" ^
  --exclude-module numpy ^
  --exclude-module scipy ^
  --exclude-module pandas ^
  --exclude-module matplotlib ^
  --exclude-module PIL ^
  --exclude-module pytest ^
  --exclude-module comtypes.npsupport ^
  run_app.py
if errorlevel 1 (
  echo Build failed.
  exit /b 1
)
copy /Y "docs\使用说明.md" "dist\底稿未完成助手\使用说明.md" >nul
echo.
echo Build complete: dist\底稿未完成助手\底稿未完成助手.exe
