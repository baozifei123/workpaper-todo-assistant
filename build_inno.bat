@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "ISCC="
where iscc >nul 2>nul
if not errorlevel 1 set "ISCC=iscc"
if not defined ISCC if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not defined ISCC (
  echo 未找到 Inno Setup 6 的 iscc.exe。
  echo 请安装 Inno Setup 6 后重试，或使用发布包中的“安装程序.exe”。
  exit /b 2
)
"%ISCC%" installer.iss
if errorlevel 1 exit /b 1
echo Inno Setup 安装包已生成到 release 目录。
