# -*- coding: utf-8 -*-
"""Windows 窗口/进程/全局热键工具（纯 ctypes，无第三方依赖）。"""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import threading
from dataclasses import dataclass
from typing import Callable, List, Optional, Tuple

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

# --- constants -------------------------------------------------------------
GA_ROOT = 2
GA_ROOTOWNER = 3
GW_OWNER = 4
SW_RESTORE = 9
SW_SHOW = 5
WM_HOTKEY = 0x0312
WM_QUIT = 0x0012
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
MOD_NOREPEAT = 0x4000
ERROR_HOTKEY_ALREADY_REGISTERED = 1409
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
FILE_READ_ATTRIBUTES = 0x0080
FILE_SHARE_READ_WRITE_DELETE = 0x0007
OPEN_EXISTING = 3
FILE_FLAG_BACKUP_SEMANTICS = 0x02000000
INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value

# --- prototypes ------------------------------------------------------------
user32.GetForegroundWindow.restype = wt.HWND
user32.GetForegroundWindow.argtypes = []

user32.GetClassNameW.argtypes = [wt.HWND, wt.LPWSTR, ctypes.c_int]
user32.GetClassNameW.restype = ctypes.c_int

user32.GetWindowTextW.argtypes = [wt.HWND, wt.LPWSTR, ctypes.c_int]
user32.GetWindowTextW.restype = ctypes.c_int

user32.GetWindowThreadProcessId.argtypes = [wt.HWND, ctypes.POINTER(wt.DWORD)]
user32.GetWindowThreadProcessId.restype = wt.DWORD

user32.GetAncestor.argtypes = [wt.HWND, ctypes.c_uint]
user32.GetAncestor.restype = wt.HWND

user32.GetWindow.argtypes = [wt.HWND, ctypes.c_uint]
user32.GetWindow.restype = wt.HWND

user32.IsWindowVisible.argtypes = [wt.HWND]
user32.IsWindowVisible.restype = wt.BOOL

user32.IsIconic.argtypes = [wt.HWND]
user32.IsIconic.restype = wt.BOOL

user32.ShowWindow.argtypes = [wt.HWND, ctypes.c_int]
user32.ShowWindow.restype = wt.BOOL

user32.SetForegroundWindow.argtypes = [wt.HWND]
user32.SetForegroundWindow.restype = wt.BOOL

user32.BringWindowToTop.argtypes = [wt.HWND]
user32.BringWindowToTop.restype = wt.BOOL

user32.AttachThreadInput.argtypes = [wt.DWORD, wt.DWORD, wt.BOOL]
user32.AttachThreadInput.restype = wt.BOOL

kernel32.GetCurrentThreadId.argtypes = []
kernel32.GetCurrentThreadId.restype = wt.DWORD

user32.RegisterHotKey.argtypes = [wt.HWND, ctypes.c_int, ctypes.c_uint, ctypes.c_uint]
user32.RegisterHotKey.restype = wt.BOOL

user32.UnregisterHotKey.argtypes = [wt.HWND, ctypes.c_int]
user32.UnregisterHotKey.restype = wt.BOOL

user32.PostThreadMessageW.argtypes = [wt.DWORD, ctypes.c_uint, wt.WPARAM, wt.LPARAM]
user32.PostThreadMessageW.restype = wt.BOOL

user32.GetMessageW.argtypes = [ctypes.POINTER(wt.MSG), wt.HWND, ctypes.c_uint, ctypes.c_uint]
user32.GetMessageW.restype = ctypes.c_int

user32.PeekMessageW.argtypes = [ctypes.POINTER(wt.MSG), wt.HWND, ctypes.c_uint,
                               ctypes.c_uint, ctypes.c_uint]
user32.PeekMessageW.restype = wt.BOOL

user32.TranslateMessage.argtypes = [ctypes.POINTER(wt.MSG)]
user32.TranslateMessage.restype = wt.BOOL

user32.DispatchMessageW.argtypes = [ctypes.POINTER(wt.MSG)]
user32.DispatchMessageW.restype = wt.LPARAM

kernel32.OpenProcess.argtypes = [wt.DWORD, wt.BOOL, wt.DWORD]
kernel32.OpenProcess.restype = wt.HANDLE

kernel32.QueryFullProcessImageNameW.argtypes = [
    wt.HANDLE, wt.DWORD, wt.LPWSTR, ctypes.POINTER(wt.DWORD)]
kernel32.QueryFullProcessImageNameW.restype = wt.BOOL

kernel32.CloseHandle.argtypes = [wt.HANDLE]
kernel32.CloseHandle.restype = wt.BOOL

kernel32.CreateMutexW.argtypes = [ctypes.c_void_p, wt.BOOL, wt.LPCWSTR]
kernel32.CreateMutexW.restype = wt.HANDLE

user32.FindWindowW.argtypes = [wt.LPCWSTR, wt.LPCWSTR]
user32.FindWindowW.restype = wt.HWND

# --- monitor/work-area helpers --------------------------------------------
MONITOR_DEFAULTTONEAREST = 2
MONITORINFOF_PRIMARY = 1


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", wt.LONG),
        ("top", wt.LONG),
        ("right", wt.LONG),
        ("bottom", wt.LONG),
    ]


class MONITORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wt.DWORD),
        ("rcMonitor", RECT),
        ("rcWork", RECT),
        ("dwFlags", wt.DWORD),
    ]


MONITORENUMPROC = ctypes.WINFUNCTYPE(
    wt.BOOL, wt.HANDLE, wt.HDC, ctypes.POINTER(RECT), wt.LPARAM)

user32.MonitorFromWindow.argtypes = [wt.HWND, wt.DWORD]
user32.MonitorFromWindow.restype = wt.HANDLE
user32.MonitorFromPoint.argtypes = [wt.POINT, wt.DWORD]
user32.MonitorFromPoint.restype = wt.HANDLE
user32.GetMonitorInfoW.argtypes = [wt.HANDLE, ctypes.POINTER(MONITORINFO)]
user32.GetMonitorInfoW.restype = wt.BOOL
user32.EnumDisplayMonitors.argtypes = [
    wt.HDC, ctypes.POINTER(RECT), MONITORENUMPROC, wt.LPARAM]
user32.EnumDisplayMonitors.restype = wt.BOOL



class BY_HANDLE_FILE_INFORMATION(ctypes.Structure):
    _fields_ = [
        ("dwFileAttributes", wt.DWORD),
        ("ftCreationTime", wt.FILETIME),
        ("ftLastAccessTime", wt.FILETIME),
        ("ftLastWriteTime", wt.FILETIME),
        ("dwVolumeSerialNumber", wt.DWORD),
        ("nFileSizeHigh", wt.DWORD),
        ("nFileSizeLow", wt.DWORD),
        ("nNumberOfLinks", wt.DWORD),
        ("nFileIndexHigh", wt.DWORD),
        ("nFileIndexLow", wt.DWORD),
    ]


kernel32.CreateFileW.argtypes = [
    wt.LPCWSTR, wt.DWORD, wt.DWORD, ctypes.c_void_p, wt.DWORD,
    wt.DWORD, wt.HANDLE,
]
kernel32.CreateFileW.restype = wt.HANDLE
kernel32.GetFileInformationByHandle.argtypes = [
    wt.HANDLE, ctypes.POINTER(BY_HANDLE_FILE_INFORMATION)]
kernel32.GetFileInformationByHandle.restype = wt.BOOL

WNDENUMPROC = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

user32.EnumWindows.argtypes = [WNDENUMPROC, wt.LPARAM]
user32.EnumWindows.restype = wt.BOOL
user32.EnumChildWindows.argtypes = [wt.HWND, WNDENUMPROC, wt.LPARAM]
user32.EnumChildWindows.restype = wt.BOOL


# --- basic helpers ---------------------------------------------------------
def get_foreground_window() -> int:
    return int(user32.GetForegroundWindow() or 0)


def get_window_class(hwnd: int) -> str:
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def get_window_text(hwnd: int) -> str:
    buf = ctypes.create_unicode_buffer(512)
    user32.GetWindowTextW(hwnd, buf, 512)
    return buf.value


def get_window_pid(hwnd: int) -> int:
    pid = wt.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return int(pid.value)


def get_window_thread_id(hwnd: int) -> int:
    pid = wt.DWORD()
    return int(user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid)))


def get_process_path(pid: int) -> str:
    """Return full executable path for a PID, or '' if unavailable."""
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return ""
    try:
        size = wt.DWORD(1024)
        buf = ctypes.create_unicode_buffer(size.value)
        if kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size)):
            return buf.value
        return ""
    finally:
        kernel32.CloseHandle(handle)


def file_identity(path: str) -> tuple[str, str]:
    """Return (volume serial, file index) for a file, or ('', '') on failure.

    This is a hint, not a permanent global identifier: moving a file within the
    same volume preserves it, while copying the file creates a new identity.
    """
    if not path:
        return "", ""
    handle = kernel32.CreateFileW(
        str(path), FILE_READ_ATTRIBUTES, FILE_SHARE_READ_WRITE_DELETE,
        None, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, None)
    if not handle or handle == INVALID_HANDLE_VALUE:
        return "", ""
    try:
        info = BY_HANDLE_FILE_INFORMATION()
        if not kernel32.GetFileInformationByHandle(handle, ctypes.byref(info)):
            return "", ""
        return ("%08X" % int(info.dwVolumeSerialNumber),
                "%08X%08X" % (int(info.nFileIndexHigh), int(info.nFileIndexLow)))
    except Exception:
        return "", ""
    finally:
        kernel32.CloseHandle(handle)


def enum_top_windows(class_name: Optional[str] = None,
                     visible_only: bool = False) -> List[int]:
    result: List[int] = []

    def _cb(hwnd, _lparam):
        if class_name is not None and get_window_class(hwnd) != class_name:
            return True
        if visible_only and not user32.IsWindowVisible(hwnd):
            return True
        result.append(int(hwnd))
        return True

    user32.EnumWindows(WNDENUMPROC(_cb), 0)
    return result


def find_child_by_class(parent: int, class_name: str) -> List[int]:
    result: List[int] = []

    def _cb(hwnd, _lparam):
        if get_window_class(hwnd) == class_name:
            result.append(int(hwnd))
        return True

    user32.EnumChildWindows(parent, WNDENUMPROC(_cb), 0)
    return result


def resolve_sheet_window(hwnd: int) -> Optional[int]:
    """Resolve a sheet view or its child, but never a modal owned dialog."""
    if not hwnd:
        return None
    if get_window_class(hwnd) == "XLMAIN":
        return int(hwnd)
    root = user32.GetAncestor(hwnd, GA_ROOT)
    if root and get_window_class(root) == "XLMAIN":
        return int(root)
    return None


def _monitor_work_area(hmon) -> Optional[Tuple[int, int, int, int, int]]:
    if not hmon:
        return None
    info = MONITORINFO()
    info.cbSize = ctypes.sizeof(info)
    if not user32.GetMonitorInfoW(hmon, ctypes.byref(info)):
        return None
    r = info.rcWork
    return (int(r.left), int(r.top), int(r.right), int(r.bottom),
            1 if info.dwFlags & MONITORINFOF_PRIMARY else 0)


def list_monitor_work_areas() -> List[Tuple[int, int, int, int, int]]:
    """Return each monitor's work area as (left, top, right, bottom, primary)."""
    result: List[Tuple[int, int, int, int, int]] = []

    def _cb(hmon, _hdc, _rect, _data):
        area = _monitor_work_area(hmon)
        if area is not None:
            result.append(area)
        return True

    user32.EnumDisplayMonitors(None, None, MONITORENUMPROC(_cb), 0)
    return result


def monitor_work_area_for_window(hwnd: int) -> Optional[Tuple[int, int, int, int]]:
    """Work area of the monitor nearest to a window handle."""
    if not hwnd:
        return None
    root = user32.GetAncestor(int(hwnd), GA_ROOT) or int(hwnd)
    area = _monitor_work_area(
        user32.MonitorFromWindow(root, MONITOR_DEFAULTTONEAREST))
    return area[:4] if area else None


def monitor_work_area_for_point(x: int, y: int) -> Optional[Tuple[int, int, int, int]]:
    """Work area of the monitor nearest to a screen point."""
    point = wt.POINT(int(x), int(y))
    area = _monitor_work_area(
        user32.MonitorFromPoint(point, MONITOR_DEFAULTTONEAREST))
    return area[:4] if area else None


def intersecting_work_area(
        rect: Tuple[int, int, int, int],
        work_areas: List[Tuple[int, int, int, int, int]]
) -> Optional[Tuple[int, int, int, int, int]]:
    """Pick the work area with the largest overlap, else the nearest one."""
    if not work_areas:
        return None
    x, y, w, h = (int(v) for v in rect[:4])
    if w <= 0 or h <= 0:
        return None
    best = None
    best_overlap = 0
    for area in work_areas:
        left, top, right, bottom = area[:4]
        overlap_x = max(0, min(x + w, right) - max(x, left))
        overlap_y = max(0, min(y + h, bottom) - max(y, top))
        overlap = overlap_x * overlap_y
        if overlap > best_overlap:
            best_overlap = overlap
            best = area
    if best is not None:
        return best
    cx, cy = x + w / 2.0, y + h / 2.0
    return min(work_areas, key=lambda area: (
        (cx - (area[0] + area[2]) / 2.0) ** 2
        + (cy - (area[1] + area[3]) / 2.0) ** 2))


def clamp_rect_to_work_area(
        rect: Tuple[int, int, int, int],
        work_area: Tuple[int, int, int, int, int],
        min_w: int = 0, min_h: int = 0, margin: int = 8
) -> Tuple[int, int, int, int]:
    """Keep a window rect fully inside a monitor work area."""
    left, top, right, bottom = work_area[:4]
    max_w = max(1, right - left - 2 * margin)
    max_h = max(1, bottom - top - 2 * margin)
    w = min(max(int(rect[2]), int(min_w)), max_w)
    h = min(max(int(rect[3]), int(min_h)), max_h)
    x = min(max(int(rect[0]), left + margin), right - margin - w)
    y = min(max(int(rect[1]), top + margin), bottom - margin - h)
    return (x, y, w, h)


def bring_to_front(hwnd: int) -> bool:
    """Restore and activate a window, working around foreground restrictions."""
    if not hwnd:
        return False
    try:
        if user32.IsIconic(hwnd):
            user32.ShowWindow(hwnd, SW_RESTORE)
        else:
            user32.ShowWindow(hwnd, SW_SHOW)
        fg = get_foreground_window()
        if fg == hwnd:
            return True
        fg_tid = get_window_thread_id(fg) if fg else 0
        cur_tid = int(kernel32.GetCurrentThreadId())
        attached = False
        if fg_tid and fg_tid != cur_tid:
            attached = bool(user32.AttachThreadInput(cur_tid, fg_tid, True))
        try:
            user32.BringWindowToTop(hwnd)
            return bool(user32.SetForegroundWindow(hwnd))
        finally:
            if attached:
                user32.AttachThreadInput(cur_tid, fg_tid, False)
    except Exception:
        return False


# --- global hotkey ---------------------------------------------------------
_VK_NAMES = {chr(c).lower(): c for c in range(ord("A"), ord("Z") + 1)}
_VK_NAMES.update({str(d): ord(str(d)) for d in range(10)})
_VK_NAMES.update({"f%d" % i: 0x6F + i for i in range(1, 13)})


@dataclass(frozen=True)
class HotkeySpec:
    mods: int
    vk: int
    text: str


def parse_hotkey(text: str) -> HotkeySpec:
    """Parse 'Ctrl+Alt+M' into a HotkeySpec. Raises ValueError if invalid."""
    raw = (text or "").replace(" ", "")
    parts = [p for p in raw.split("+") if p]
    if not parts:
        raise ValueError("热键为空")
    mods = 0
    key = None
    for part in parts:
        low = part.lower()
        if low in ("ctrl", "control"):
            mods |= MOD_CONTROL
        elif low == "alt":
            mods |= MOD_ALT
        elif low == "shift":
            mods |= MOD_SHIFT
        elif low in ("win", "windows"):
            mods |= MOD_WIN
        else:
            if key is not None:
                raise ValueError("热键只能包含一个主按键")
            key = low
    if not key or key not in _VK_NAMES:
        raise ValueError("不支持的主按键：%s" % (key or "无"))
    if not mods:
        raise ValueError("热键必须包含 Ctrl/Alt/Shift/Win 中的至少一个")
    names = []
    if mods & MOD_CONTROL:
        names.append("Ctrl")
    if mods & MOD_ALT:
        names.append("Alt")
    if mods & MOD_SHIFT:
        names.append("Shift")
    if mods & MOD_WIN:
        names.append("Win")
    names.append(key.upper() if len(key) == 1 else key.upper())
    return HotkeySpec(mods | MOD_NOREPEAT, _VK_NAMES[key], "+".join(names))


HOTKEY_PRESETS = [
    "Ctrl+Alt+M",
    "Ctrl+Alt+K",
    "Ctrl+Alt+J",
    "Ctrl+Shift+M",
    "Alt+Shift+M",
    "Ctrl+Alt+F9",
]


class HotkeyError(Exception):
    def __init__(self, message: str, code: int = 0):
        super().__init__(message)
        self.code = code


class HotkeyThread(threading.Thread):
    """Dedicated thread that owns a global hotkey and its message queue."""

    def __init__(self, hotkey: str, callback: Callable[[], None]):
        super().__init__(daemon=True)
        self.spec = parse_hotkey(hotkey)
        self.callback = callback
        self.thread_id = 0
        self.error: Optional[HotkeyError] = None
        self._ready = threading.Event()
        self._stop_requested = False

    def wait_ready(self, timeout: float = 3.0) -> bool:
        return self._ready.wait(timeout)

    def run(self):
        self.thread_id = int(kernel32.GetCurrentThreadId())
        ok = user32.RegisterHotKey(None, 1, self.spec.mods, self.spec.vk)
        if not ok:
            code = ctypes.get_last_error()
            if code == ERROR_HOTKEY_ALREADY_REGISTERED:
                msg = "热键 %s 已被其他程序占用，请换一个组合键" % self.spec.text
            else:
                msg = "注册热键 %s 失败（错误码 %d）" % (self.spec.text, code)
            self.error = HotkeyError(msg, code)
            self._ready.set()
            return
        msg = wt.MSG()
        # Create the thread message queue before signalling readiness, so a
        # concurrent stop/change can always post WM_QUIT to it.
        user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 0)
        self._ready.set()
        try:
            while not self._stop_requested:
                ret = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
                if ret == 0 or ret == -1:
                    break
                if msg.message == WM_HOTKEY and int(msg.wParam) == 1:
                    try:
                        self.callback()
                    except Exception:
                        pass
        finally:
            user32.UnregisterHotKey(None, 1)

    def stop(self):
        self._stop_requested = True
        if self.thread_id:
            user32.PostThreadMessageW(self.thread_id, WM_QUIT, 0, 0)
