# -*- coding: utf-8 -*-
"""从清单定位回底稿：打开/切换工作簿，验证内容，选中单元格。

安全约束：本模块只读取工作簿路径和单元格内容，只执行 Activate/Select，
绝不调用 Save/Close/Quit，也不修改任何单元格。
"""
from __future__ import annotations

import glob
import os
import re
import subprocess
import time
from dataclasses import dataclass
from typing import Optional, Tuple

from . import capture, nativeom, win32util
from .models import Item


class JumpError(Exception):
    kind = "general"

    def __init__(self, message: str, detail: str = ""):
        super().__init__(message)
        self.user_message = message
        self.detail = detail


class FileMissing(JumpError):
    kind = "file_missing"


class SheetMissing(JumpError):
    kind = "sheet_missing"


class ContentMismatch(JumpError):
    kind = "content_mismatch"

    def __init__(self, old: str, new: str):
        super().__init__("单元格内容与登记时不一致")
        self.old = old
        self.new = new


class OpenTimeout(JumpError):
    kind = "open_timeout"


@dataclass
class JumpOutcome:
    file_path: str
    sheet_name: str
    cell_address: str
    source_app: str
    window_hwnd: int


def normalize_text(text: str) -> str:
    # Normalise line-ending representation only; other whitespace may be part
    # of an audit procedure and must not silently bypass the mismatch warning.
    return (text or "").replace("\r\n", "\n").replace("\r", "\n")


def same_text(a: str, b: str) -> bool:
    return normalize_text(a) == normalize_text(b)


def find_open_workbook(item: Item) -> Optional[Tuple[nativeom.SheetWindow, object]]:
    """Find an already-open workbook matching item.file_path.

    Returns (sheet_window, workbook) or None. Prefers the recorded source app.
    """
    # Recompute from the actual path rather than trusting a stale cached key.
    target = nativeom.path_key(item.file_path)
    if not target:
        return None

    apps = []
    seen = set()
    for sw in nativeom.iter_sheet_windows(visible_only=True):
        try:
            app_hwnd = int(getattr(sw.app, "Hwnd", 0) or 0)
        except Exception:
            app_hwnd = 0
        key = (sw.source_app, app_hwnd or sw.xlmain)
        if key in seen:
            continue
        seen.add(key)
        apps.append(sw)

    for preferred in (True, False):
        for sw in apps:
            if preferred and item.source_app and sw.source_app != item.source_app:
                continue
            try:
                for wb in nativeom.iter_workbooks(sw.app):
                    if nativeom.is_addin(wb):
                        continue
                    if nativeom.path_key(nativeom.workbook_path(wb)) == target:
                        return sw, wb
            except Exception:
                continue
    return None


def _resolve_exe(app: str) -> str:
    if app == "excel":
        candidates = []
        try:
            import winreg
            for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                try:
                    with winreg.OpenKey(
                            hive,
                            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe") as key:
                        candidates.append(winreg.QueryValue(key, None))
                except OSError:
                    pass
        except Exception:
            pass
        for path in candidates:
            if path and os.path.exists(path):
                return path
        return ""
    if app == "wps":
        candidates = []
        try:
            import winreg
            for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                for sub in (
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\et.exe",
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\wps.exe",
                ):
                    try:
                        with winreg.OpenKey(hive, sub) as key:
                            candidates.append(winreg.QueryValue(key, None))
                    except OSError:
                        pass
        except Exception:
            pass
        patterns = []
        for base in (
            os.environ.get("LOCALAPPDATA", ""),
            os.environ.get("ProgramFiles", ""),
            os.environ.get("ProgramFiles(x86)", ""),
        ):
            if base:
                patterns.append(os.path.join(
                    base, "Kingsoft", "WPS Office", "*", "office6", "et.exe"))
                patterns.append(os.path.join(
                    base, "WPS Office", "*", "office6", "et.exe"))
        for pattern in patterns:
            candidates.extend(glob.glob(pattern))
        for path in candidates:
            if path and os.path.exists(path):
                return path
        return ""
    return ""


def _launch(item: Item) -> None:
    exe = item.source_exe if item.source_exe and os.path.exists(item.source_exe) else ""
    if not exe:
        exe = _resolve_exe(item.source_app)
    try:
        if exe:
            subprocess.Popen(
                [exe, item.file_path],
                cwd=os.path.dirname(item.file_path) or None,
                close_fds=True)
        else:
            os.startfile(item.file_path)  # noqa: S606 - user-requested open
    except Exception as exc:
        raise JumpError("无法启动 Excel/WPS 打开该文件：%s" % exc) from exc


def open_workbook(item: Item, timeout: float = 30.0):
    if not os.path.exists(item.file_path):
        raise FileMissing("底稿文件不存在：%s" % item.file_path)
    _launch(item)
    deadline = time.time() + timeout
    while time.time() < deadline:
        found = find_open_workbook(item)
        if found:
            return found
        time.sleep(0.5)
    raise OpenTimeout(
        "已尝试打开文件，但未检测到对应的表格窗口。"
        "可能被密码、文件占用或程序弹窗阻止，请人工确认文件是否已打开。")


def find_xlmain_for_workbook(workbook) -> int:
    target = nativeom.path_key(nativeom.workbook_path(workbook))
    if not target:
        return 0
    for sw in nativeom.iter_sheet_windows(visible_only=True):
        try:
            if nativeom.path_key(nativeom.workbook_path(sw.workbook)) == target:
                return sw.xlmain
        except Exception:
            continue
    return 0


def current_snapshot(workbook, sheet_name: str, address: str) -> str:
    try:
        return capture.read_cell_snapshot(workbook, sheet_name, address)
    except Exception as exc:
        raise JumpError("无法读取目标单元格内容，无法核对原位置，请稍后重试。",
                        detail=repr(exc)) from exc


def _activate_workbook(wb, app) -> None:
    """Activate a workbook and verify the application really switched to it."""
    errors = []
    actions = (
        lambda: nativeom.retry(wb.Activate, attempts=4, delay=0.2),
        lambda: nativeom.retry(lambda: wb.Windows.Item(1).Activate(),
                               attempts=4, delay=0.2),
    )
    for action in actions:
        try:
            action()
        except Exception as exc:
            errors.append(repr(exc))
            continue
        try:
            active = nativeom.retry(lambda: app.ActiveWorkbook,
                                    attempts=3, delay=0.15)
            active_key = nativeom.path_key(nativeom.workbook_path(active))
            target_key = nativeom.path_key(nativeom.workbook_path(wb))
            if target_key and active_key == target_key:
                return
            errors.append("激活后的工作簿与目标不一致")
        except Exception as exc:
            errors.append(repr(exc))
    raise JumpError(
        "无法切换到目标工作簿，请人工在 Excel/WPS 中切换后再试。",
        detail="; ".join(errors))


def jump_to_item(item: Item, allow_mismatch: bool = False) -> JumpOutcome:
    if not os.path.exists(item.file_path):
        raise FileMissing("底稿文件不存在：%s" % item.file_path)

    found = find_open_workbook(item)
    if not found:
        found = open_workbook(item)
    sw, wb = found
    actual_path = nativeom.workbook_path(wb)
    if not actual_path or nativeom.path_key(actual_path) != nativeom.path_key(item.file_path):
        raise JumpError("无法确认已打开的工作簿与登记文件一致，请人工核对后重试。")
    try:
        app = wb.Application
    except Exception as exc:
        raise JumpError("无法连接该工作簿，请人工打开后再试。", detail=repr(exc)) from exc

    _activate_workbook(wb, app)

    try:
        ws = capture.get_worksheet(wb, item.sheet_name)
    except Exception as exc:
        raise SheetMissing(
            "工作簿中找不到工作表「%s」，文件可能已被修改。" % item.sheet_name) from exc

    try:
        nativeom.retry(ws.Activate, attempts=4, delay=0.2)
    except Exception as exc:
        raise JumpError(
            "无法激活工作表「%s」（可能被隐藏或保护）。" % item.sheet_name,
            detail=repr(exc)) from exc

    latest = current_snapshot(wb, item.sheet_name, item.cell_address)
    if not same_text(item.snapshot, latest) and not allow_mismatch:
        raise ContentMismatch(item.snapshot, latest)

    try:
        nativeom.retry(lambda: ws.Range(item.cell_address).Select(),
                       attempts=4, delay=0.2)
    except Exception as exc:
        raise JumpError(
            "无法选中单元格 %s!%s，请人工确认工作表是否受保护。" %
            (item.sheet_name, item.cell_address), detail=repr(exc)) from exc

    hwnd = find_xlmain_for_workbook(wb) or sw.xlmain
    brought_forward = win32util.bring_to_front(hwnd)
    if (not brought_forward and
            win32util.resolve_sheet_window(win32util.get_foreground_window()) != hwnd):
        raise JumpError("目标单元格已选中，但 Windows 未允许切换到前台；"
                        "请在任务栏中手动切换到对应底稿核对。")
    return JumpOutcome(
        file_path=item.file_path,
        sheet_name=item.sheet_name,
        cell_address=item.cell_address,
        source_app=sw.source_app,
        window_hwnd=hwnd,
    )
