param(
    [string]$Target = (Join-Path $env:LOCALAPPDATA "Programs\底稿未完成助手"),
    [switch]$RemoveData
)

$ErrorActionPreference = "Stop"
if (Get-Process -Name "底稿未完成助手" -ErrorAction SilentlyContinue) {
    throw "底稿未完成助手正在运行，请先退出程序再卸载。"
}

$shortcut = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\底稿未完成助手.lnk"
if (Test-Path $shortcut) {
    Remove-Item -Force $shortcut
}
if (Test-Path $Target) {
    Remove-Item -Recurse -Force $Target
    Write-Host "已删除程序目录：$Target"
} else {
    Write-Host "程序目录不存在：$Target"
}

$dataDir = Join-Path $env:LOCALAPPDATA "WorkpaperTodo"
if ($RemoveData) {
    if (Test-Path $dataDir) {
        Remove-Item -Recurse -Force $dataDir
        Write-Host "已删除数据目录：$dataDir"
    }
} else {
    Write-Host "事项数据保留在：$dataDir"
    Write-Host "如需一并删除，请运行：.\uninstall_per_user.ps1 -RemoveData"
}
