param(
    [string]$Source = "",
    [string]$Target = (Join-Path $env:LOCALAPPDATA "Programs\底稿未完成助手"),
    [switch]$NoShortcut,
    [switch]$NoLaunch
)

$ErrorActionPreference = "Stop"
$exeName = "底稿未完成助手.exe"
if (-not $Source) {
    $candidates = @(
        (Join-Path $PSScriptRoot "底稿未完成助手"),
        (Join-Path $PSScriptRoot "..\dist\底稿未完成助手")
    )
    $Source = $candidates | Where-Object { Test-Path (Join-Path $_ $exeName) } | Select-Object -First 1
}
if (-not $Source) {
    throw "找不到程序文件夹，请确认安装包完整。"
}
$sourceExe = Join-Path $Source $exeName

if (-not (Test-Path $sourceExe)) {
    throw "找不到安装源：$sourceExe"
}
if (Get-Process -Name "底稿未完成助手" -ErrorAction SilentlyContinue) {
    throw "底稿未完成助手正在运行，请先退出程序再安装。"
}

New-Item -ItemType Directory -Force -Path $Target | Out-Null
Copy-Item -Path (Join-Path $Source "*") -Destination $Target -Recurse -Force

if (-not $NoShortcut) {
    $startMenu = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"
    New-Item -ItemType Directory -Force -Path $startMenu | Out-Null
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut((Join-Path $startMenu "底稿未完成助手.lnk"))
    $shortcut.TargetPath = Join-Path $Target $exeName
    $shortcut.WorkingDirectory = $Target
    $shortcut.IconLocation = Join-Path $Target $exeName
    $shortcut.Description = "审计底稿未完成事项管理"
    $shortcut.Save()
}

Write-Host "已安装到：$Target"
Write-Host "数据目录不变：$env:LOCALAPPDATA\WorkpaperTodo"
if (-not $NoLaunch) {
    Start-Process (Join-Path $Target $exeName)
}
