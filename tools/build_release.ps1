param(
    [string]$Version = "0.3.0"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$dist = Join-Path $root "dist\底稿未完成助手"
$exe = Join-Path $dist "底稿未完成助手.exe"
if (-not (Test-Path $exe)) {
    throw "请先运行 build_exe.bat 生成程序：$exe"
}

$releaseRoot = Join-Path $root "release"
$stage = Join-Path $releaseRoot ("底稿未完成助手-" + $Version)
if (Test-Path $stage) {
    Remove-Item -Recurse -Force $stage
}
New-Item -ItemType Directory -Force -Path $stage | Out-Null

$appStage = Join-Path $stage "底稿未完成助手"
Copy-Item -Path $dist -Destination $appStage -Recurse -Force
$installerExe = Join-Path $root "dist\安装程序.exe"
if (Test-Path $installerExe) {
    Copy-Item -Path $installerExe -Destination $stage -Force
}
$innoSetupExe = Join-Path $releaseRoot ("底稿未完成助手-" + $Version + "-安装程序.exe")
if (Test-Path $innoSetupExe) {
    Copy-Item -Path $innoSetupExe -Destination $stage -Force
}
Copy-Item -Path (Join-Path $root "installer.iss") -Destination $stage -Force -ErrorAction SilentlyContinue
Copy-Item -Path (Join-Path $root "build_inno.bat") -Destination $stage -Force -ErrorAction SilentlyContinue
Copy-Item -Path (Join-Path $root "docs\使用说明.md") -Destination $stage -Force
Copy-Item -Path (Join-Path $root "docs\使用说明.md") -Destination $appStage -Force
Copy-Item -Path (Join-Path $PSScriptRoot "install_per_user.ps1") -Destination $stage -Force
Copy-Item -Path (Join-Path $PSScriptRoot "uninstall_per_user.ps1") -Destination $stage -Force
Copy-Item -Path (Join-Path $PSScriptRoot "sign_exe.ps1") -Destination $stage -Force
Copy-Item -Path (Join-Path $PSScriptRoot "test_signing.ps1") -Destination $stage -Force
Copy-Item -Path (Join-Path $root "docs\发布与签名说明.md") -Destination $stage -Force -ErrorAction SilentlyContinue
Copy-Item -Path (Join-Path $root "docs\加载项方案评估.md") -Destination $stage -Force -ErrorAction SilentlyContinue
Copy-Item -Path (Join-Path $root "docs\兼容性验证记录.md") -Destination $stage -Force -ErrorAction SilentlyContinue
Copy-Item -Path (Join-Path $root "CHANGELOG.md") -Destination $stage -Force -ErrorAction SilentlyContinue
Copy-Item -Path (Join-Path $root "README.md") -Destination $stage -Force -ErrorAction SilentlyContinue

$installBat = Join-Path $stage "安装.bat"
$installContent = @"
@echo off
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_per_user.ps1"
pause
"@
Set-Content -Path $installBat -Value $installContent -Encoding UTF8

$uninstallBat = Join-Path $stage "卸载.bat"
$uninstallContent = @"
@echo off
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall_per_user.ps1"
pause
"@
Set-Content -Path $uninstallBat -Value $uninstallContent -Encoding UTF8

$installReadme = Join-Path $stage "安装说明.txt"
$installReadmeContent = @"
安装方法（任选一种）：
1. 双击“安装程序.exe”，按提示安装到当前用户目录，不需要管理员权限。
2. 双击“底稿未完成助手-$Version-安装程序.exe”（Inno Setup 标准安装包），按向导安装。
3. 双击“安装.bat”，使用 PowerShell 脚本安装。

安装后从开始菜单启动“底稿未完成助手”。数据默认保存在 %LOCALAPPDATA%\WorkpaperTodo。
卸载：开始菜单“卸载底稿未完成助手”，或运行“卸载.bat”。卸载默认保留事项数据。
"@
Set-Content -Path $installReadme -Value $installReadmeContent -Encoding UTF8

$zip = $stage + ".zip"
if (Test-Path $zip) {
    Remove-Item -Force $zip
}
Compress-Archive -Path (Join-Path $stage "*") -DestinationPath $zip -Force

Write-Host "发布目录：$stage"
Write-Host "发布压缩包：$zip"
Get-FileHash $exe -Algorithm SHA256 | Format-List Path, Hash
if (Test-Path $innoSetupExe) {
    Get-FileHash $innoSetupExe -Algorithm SHA256 | Format-List Path, Hash
}
Get-FileHash $zip -Algorithm SHA256 | Format-List Path, Hash
