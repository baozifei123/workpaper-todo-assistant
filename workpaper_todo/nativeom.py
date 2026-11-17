# -*- coding: utf-8 -*-
"""通过窗口 NativeOM 访问 Excel / WPS 表格的对象模型。

原理：Excel 和 WPS 表格的工作表视图窗口类名都是 EXCEL7，
调用 oleacc.AccessibleObjectFromWindow(EXCEL7, OBJID_NATIVEOM) 可以拿到
该窗口的 Window 对象，再通过 Window.Parent 得到工作簿、通过
Workbook.Application 得到 Application。这种方式不需要注册 COM ProgID，
也能准确连接到用户当前正在使用的那个窗口/实例。
"""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import os
import time
from dataclasses import dataclass
from typing import Iterator, List, Optional

import comtypes
import comtypes.automation
import comtypes.client.dynamic

from . import win32util

oleacc = ctypes.WinDLL("oleacc", use_last_error=True)

OBJID_NATIVEOM = 0xFFFFFFF0
IDispatchPtr = ctypes.POINTER(comtypes.automation.IDispatch)

oleacc.AccessibleObjectFromWindow.argtypes = [
    wt.HWND,
    wt.DWORD,
    ctypes.POINTER(comtypes.GUID),
    ctypes.POINTER(IDispatchPtr),
]
oleacc.AccessibleObjectFromWindow.restype = ctypes.c_long


class NativeOMError(Exception):
    pass


# Excel/WPS may reject COM calls while busy (editing, calculating, modal dialog).
# These HRESULTs are worth a short retry.
_RETRY_HRESULTS = {
    -2147418111,  # RPC_E_CALL_REJECTED
    -2147417846,  # RPC_E_SERVERCALL_RETRYLATER
    -2147417845,  # RPC_E_SERVERCALL_REJECTED
    -2147417844,
    -2147417843,
    -2147417842,
}


def is_retryable(exc: Exception) -> bool:
    hr = getattr(exc, "hresult", None)
    if hr is None and exc.args and isinstance(exc.args[0], int):
        hr = exc.args[0]
    if hr is not None and (hr & 0xFFFFFFFF) >= 0x80010000:
        if hr in _RETRY_HRESULTS:
            return True
    text = str(exc).lower()
    return "rejected" in text or "拒绝" in text or "retry" in text


def retry(func, attempts: int = 5, delay: float = 0.2):
    """Call func, retrying transient 'call rejected' errors."""
    last = None
    for i in range(attempts):
        try:
            return func()
        except Exception as exc:
            last = exc
            if i == attempts - 1 or not is_retryable(exc):
                raise
            time.sleep(delay)
    if last is not None:
        raise last


@dataclass
class SheetWindow:
    xlmain: int
    excel7: int
    win: object
    workbook: object
    app: object
    source_app: str          # 'excel' / 'wps'
    source_exe: str
    title: str


def native_window(hwnd: int):
    """Return the Window object for an EXCEL7 window handle."""
    iid = comtypes.automation.IDispatch._iid_
    hr = -1
    ptr = None
    for attempt in range(3):
        ptr = IDispatchPtr()
        hr = oleacc.AccessibleObjectFromWindow(
            wt.HWND(hwnd), OBJID_NATIVEOM, ctypes.byref(iid), ctypes.byref(ptr))
        if hr == 0:
            break
        time.sleep(0.12)
    if hr != 0:
        raise NativeOMError("无法读取该表格窗口（HRESULT 0x%08X）" % (hr & 0xFFFFFFFF))
    try:
        return comtypes.client.dynamic.Dispatch(ptr)
    except Exception as exc:
        raise NativeOMError("无法连接表格窗口：%s" % exc) from exc


def _source_app(exe_path: str) -> str:
    base = os.path.basename(exe_path or "").lower()
    if base == "et.exe":
        return "wps"
    if base == "excel.exe":
        return "excel"
    return ""


def get_sheet_object(xlmain: int) -> SheetWindow:
    """From an XLMAIN top-level window, return the workbook/app behind it."""
    children = win32util.find_child_by_class(xlmain, "EXCEL7")
    if not children:
        raise NativeOMError("当前窗口不是可读取的表格视图")
    visible = [h for h in children if win32util.user32.IsWindowVisible(h)]
    if not visible:
        raise NativeOMError("无法确认当前表格视图，请切换到目标底稿后重试。")

    candidates = []
    for excel7 in visible:
        try:
            win = native_window(excel7)
            workbook = retry(lambda: win.Parent, attempts=4, delay=0.2)
            app = retry(lambda: workbook.Application, attempts=4, delay=0.2)
            if workbook is None or app is None:
                raise NativeOMError("无法取得当前工作簿")
            candidates.append((excel7, win, workbook, app))
        except Exception as exc:
            # Never substitute Application.ActiveWorkbook when Parent fails.
            raise NativeOMError("无法确认该窗口对应的工作簿，请重试。") from exc

    if len(candidates) > 1:
        # WPS can expose multiple visible EXCEL7 views under one XLMAIN while
        # only one workbook tab is active. Select exactly one by identity.
        active_candidates = []
        for candidate in candidates:
            _, _, workbook, app = candidate
            try:
                own_path = workbook_path(workbook)
                active = retry(lambda: app.ActiveWorkbook, attempts=3, delay=0.2)
                active_path = workbook_path(active)
                if not own_path or not active_path:
                    raise NativeOMError("无法确认活动工作簿")
                if path_key(own_path) == path_key(active_path):
                    active_candidates.append(candidate)
            except Exception as exc:
                raise NativeOMError("无法确认当前活动工作簿，请重试。") from exc
        if len(active_candidates) != 1:
            raise NativeOMError("存在多个无法区分的表格视图，请切换到目标底稿后重试。")
        candidates = active_candidates

    excel7, win, workbook, app = candidates[0]

    pid = win32util.get_window_pid(xlmain)
    exe = win32util.get_process_path(pid)
    return SheetWindow(
        xlmain=int(xlmain),
        excel7=int(excel7),
        win=win,
        workbook=workbook,
        app=app,
        source_app=_source_app(exe),
        source_exe=exe,
        title=win32util.get_window_text(xlmain),
    )


def iter_sheet_windows(visible_only: bool = True) -> Iterator[SheetWindow]:
    for xlmain in win32util.enum_top_windows("XLMAIN", visible_only=visible_only):
        try:
            sw = get_sheet_object(xlmain)
        except Exception:
            continue
        if sw.source_app not in ("excel", "wps"):
            continue
        yield sw


def iter_workbooks(app) -> Iterator[object]:
    """Iterate Workbooks of an Application object."""
    try:
        count = int(retry(lambda: app.Workbooks.Count, attempts=3, delay=0.15))
    except Exception:
        return
    for i in range(1, count + 1):
        try:
            yield retry(lambda i=i: app.Workbooks.Item(i), attempts=3, delay=0.15)
        except Exception:
            continue


def workbook_path(workbook) -> str:
    try:
        return str(retry(lambda: workbook.FullName, attempts=4, delay=0.15))
    except Exception:
        return ""


def workbook_name(workbook) -> str:
    try:
        return str(retry(lambda: workbook.Name, attempts=4, delay=0.15))
    except Exception:
        return ""


def is_addin(workbook) -> bool:
    try:
        return bool(retry(lambda: workbook.IsAddin, attempts=3, delay=0.1))
    except Exception:
        return False


def path_key(path: str) -> str:
    """Normalized key for comparing file paths."""
    if not path:
        return ""
    try:
        path = os.path.abspath(os.path.expanduser(path))
    except Exception:
        pass
    return os.path.normcase(os.path.normpath(path))
