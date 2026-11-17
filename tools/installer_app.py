# -*- coding: utf-8 -*-
"""按用户安装程序（无管理员权限）。

发布包中本程序与“底稿未完成助手”文件夹放在同一目录；双击后复制到
%LOCALAPPDATA%\\Programs\\底稿未完成助手，并创建当前用户开始菜单快捷方式。
"""
from __future__ import annotations

import argparse
import ctypes
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

APP_NAME = "底稿未完成助手"
WINDOW_TITLE = "底稿未完成事项管理"
EXE_NAME = "底稿未完成助手.exe"
CREATE_NO_WINDOW = 0x08000000


def log(message: str):
    path = Path(tempfile.gettempdir()) / "wpt_installer.log"
    try:
        with open(path, "a", encoding="utf-8") as f:
            f.write(message + "\n")
    except OSError:
        pass


def payload_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / APP_NAME
    return Path(__file__).resolve().parents[1] / "dist" / APP_NAME


def default_target() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home())
    return Path(base) / "Programs" / APP_NAME


def start_menu_dir() -> Path:
    base = os.environ.get("APPDATA") or str(Path.home())
    return Path(base) / "Microsoft" / "Windows" / "Start Menu" / "Programs"


def app_is_running() -> bool:
    try:
        hwnd = ctypes.windll.user32.FindWindowW(None, WINDOW_TITLE)
        return bool(hwnd)
    except Exception:
        return False


def create_shortcut(path: Path, target: str, arguments: str = "",
                    working_dir: str = "", icon: str = ""):
    def q(value: str) -> str:
        return value.replace("'", "''")

    script = (
        "$ws = New-Object -ComObject WScript.Shell; "
        "$s = $ws.CreateShortcut('%s'); "
        "$s.TargetPath = '%s'; "
        "$s.Arguments = '%s'; "
        "$s.WorkingDirectory = '%s'; "
        "$s.IconLocation = '%s'; "
        "$s.Save()"
    ) % (q(str(path)), q(target), q(arguments), q(working_dir), q(icon or target))
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-Command", script],
        check=True, creationflags=CREATE_NO_WINDOW)


def write_uninstall_script(target: Path):
    script = r'''param([switch]$RemoveData)
$ErrorActionPreference = "Stop"
if (Get-Process -Name "底稿未完成助手" -ErrorAction SilentlyContinue) {
    throw "底稿未完成助手正在运行，请先退出程序再卸载。"
}
$target = "%s"
$shortcut = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\底稿未完成助手.lnk"
$uninstallShortcut = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\卸载底稿未完成助手.lnk"
foreach ($item in @($shortcut, $uninstallShortcut)) {
    if (Test-Path $item) { Remove-Item -Force $item }
}
if (Test-Path $target) { Remove-Item -Recurse -Force $target }
$dataDir = Join-Path $env:LOCALAPPDATA "WorkpaperTodo"
if ($RemoveData) {
    if (Test-Path $dataDir) { Remove-Item -Recurse -Force $dataDir }
    Write-Host "已删除程序和数据：$target"
} else {
    Write-Host "已删除程序目录：$target"
    Write-Host "事项数据保留在：$dataDir"
}
''' % str(target).replace('"', '""')
    (target / "uninstall_per_user.ps1").write_text(script, encoding="utf-8-sig")


def install(target: Path, no_shortcut: bool, no_launch: bool, quiet: bool):
    source = payload_dir()
    source_exe = source / EXE_NAME
    if not source_exe.is_file():
        raise RuntimeError("找不到安装源：%s" % source_exe)
    if app_is_running():
        raise RuntimeError("程序正在运行，请先从托盘退出后再安装。")
    target.mkdir(parents=True, exist_ok=True)
    shutil.copytree(source, target, dirs_exist_ok=True)
    write_uninstall_script(target)
    uninstall_script = target / "uninstall_per_user.ps1"
    if not no_shortcut:
        menu = start_menu_dir()
        menu.mkdir(parents=True, exist_ok=True)
        exe = str(target / EXE_NAME)
        create_shortcut(menu / (APP_NAME + ".lnk"), exe, working_dir=str(target),
                        icon=exe)
        create_shortcut(
            menu / ("卸载" + APP_NAME + ".lnk"),
            "powershell.exe",
            arguments='-NoProfile -ExecutionPolicy Bypass -File "%s"' % uninstall_script,
            working_dir=str(target))
    log("installed to %s" % target)
    if not quiet:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(
            "安装完成",
            "已安装到：\n%s\n\n数据目录仍为 %%LOCALAPPDATA%%\\WorkpaperTodo。"
            % target)
        root.destroy()
    if not no_launch:
        subprocess.Popen([str(target / EXE_NAME)], cwd=str(target))


def uninstall(target: Path, remove_data: bool, quiet: bool):
    if app_is_running():
        raise RuntimeError("程序正在运行，请先从托盘退出后再卸载。")
    menu = start_menu_dir()
    for name in (APP_NAME + ".lnk", "卸载" + APP_NAME + ".lnk"):
        try:
            (menu / name).unlink()
        except FileNotFoundError:
            pass
    if target.exists():
        shutil.rmtree(target)
    data_dir = Path(os.environ.get("LOCALAPPDATA") or str(Path.home())) / "WorkpaperTodo"
    if remove_data and data_dir.exists():
        shutil.rmtree(data_dir)
    log("uninstalled from %s remove_data=%s" % (target, remove_data))
    if not quiet:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(
            "卸载完成",
            "程序目录和快捷方式已删除。\n事项数据%s。"
            % ("已删除" if remove_data else "保留在：\n%s" % data_dir))
        root.destroy()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--target", default="")
    parser.add_argument("--no-shortcut", action="store_true")
    parser.add_argument("--no-launch", action="store_true")
    parser.add_argument("--quiet", action="store_true")
    parser.add_argument("--uninstall", action="store_true")
    parser.add_argument("--remove-data", action="store_true")
    args = parser.parse_args()
    target = Path(args.target) if args.target else default_target()
    try:
        if args.uninstall:
            uninstall(target, args.remove_data, args.quiet)
        else:
            install(target, args.no_shortcut, args.no_launch, args.quiet)
        return 0
    except Exception as exc:
        log("ERROR: %r" % exc)
        if not args.quiet:
            try:
                import tkinter as tk
                from tkinter import messagebox
                root = tk.Tk()
                root.withdraw()
                messagebox.showerror("安装/卸载失败", str(exc))
                root.destroy()
            except Exception:
                pass
        return 1


if __name__ == "__main__":
    sys.exit(main())
