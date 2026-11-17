# -*- coding: utf-8 -*-
"""Windows 托盘图标（纯 ctypes，无第三方依赖）。"""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import logging
import threading
from typing import Callable, Optional

log = logging.getLogger(__name__)

user32 = ctypes.WinDLL("user32", use_last_error=True)
shell32 = ctypes.WinDLL("shell32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

WM_DESTROY = 0x0002
WM_CLOSE = 0x0010
WM_COMMAND = 0x0111
WM_NULL = 0x0000
WM_CONTEXTMENU = 0x007B
WM_LBUTTONUP = 0x0202
WM_LBUTTONDBLCLK = 0x0203
WM_RBUTTONUP = 0x0205
WM_APP = 0x8000
WM_TRAY = WM_APP + 20

NIM_ADD = 0
NIM_MODIFY = 1
NIM_DELETE = 2
NIF_MESSAGE = 0x0001
NIF_ICON = 0x0002
NIF_TIP = 0x0004

MF_STRING = 0x0000
MF_SEPARATOR = 0x0800
TPM_RETURNCMD = 0x0100
TPM_RIGHTBUTTON = 0x0002
IMAGE_ICON = 1
LR_LOADFROMFILE = 0x0010
LR_DEFAULTSIZE = 0x0040
IDI_APPLICATION = 32512
CS_DBLCLKS = 0x0008

ID_SHOW = 1001
ID_CAPTURE = 1002
ID_OPEN_DATA = 1003
ID_EXIT = 1004

LRESULT = ctypes.c_ssize_t
WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM)


class WNDCLASSW(ctypes.Structure):
    _fields_ = [
        ("style", wt.UINT),
        ("lpfnWndProc", WNDPROC),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", wt.HINSTANCE),
        ("hIcon", wt.HICON),
        ("hCursor", wt.HANDLE),
        ("hbrBackground", wt.HBRUSH),
        ("lpszMenuName", wt.LPCWSTR),
        ("lpszClassName", wt.LPCWSTR),
    ]


class NOTIFYICONDATAW(ctypes.Structure):
    _fields_ = [
        ("cbSize", wt.DWORD),
        ("hWnd", wt.HWND),
        ("uID", wt.UINT),
        ("uFlags", wt.UINT),
        ("uCallbackMessage", wt.UINT),
        ("hIcon", wt.HICON),
        ("szTip", wt.WCHAR * 128),
        ("dwState", wt.DWORD),
        ("dwStateMask", wt.DWORD),
        ("szInfo", wt.WCHAR * 256),
        ("uVersion", wt.UINT),
        ("szInfoTitle", wt.WCHAR * 64),
        ("dwInfoFlags", wt.DWORD),
        ("guidItem", ctypes.c_byte * 16),
        ("hBalloonIcon", wt.HICON),
    ]


user32.DefWindowProcW.argtypes = [wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM]
user32.DefWindowProcW.restype = LRESULT
user32.RegisterClassW.argtypes = [ctypes.POINTER(WNDCLASSW)]
user32.RegisterClassW.restype = wt.ATOM
user32.UnregisterClassW.argtypes = [wt.LPCWSTR, wt.HINSTANCE]
user32.UnregisterClassW.restype = wt.BOOL
user32.CreateWindowExW.argtypes = [
    wt.DWORD, wt.LPCWSTR, wt.LPCWSTR, wt.DWORD, ctypes.c_int,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, wt.HWND, wt.HMENU,
    wt.HINSTANCE, ctypes.c_void_p,
]
user32.CreateWindowExW.restype = wt.HWND
user32.DestroyWindow.argtypes = [wt.HWND]
user32.DestroyWindow.restype = wt.BOOL
user32.PostMessageW.argtypes = [wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM]
user32.PostMessageW.restype = wt.BOOL
user32.RegisterWindowMessageW.argtypes = [wt.LPCWSTR]
user32.RegisterWindowMessageW.restype = wt.UINT
user32.GetCursorPos.argtypes = [ctypes.POINTER(wt.POINT)]
user32.GetCursorPos.restype = wt.BOOL
user32.SetForegroundWindow.argtypes = [wt.HWND]
user32.SetForegroundWindow.restype = wt.BOOL
user32.CreatePopupMenu.restype = wt.HMENU
user32.AppendMenuW.argtypes = [wt.HMENU, wt.UINT, ctypes.c_size_t, wt.LPCWSTR]
user32.AppendMenuW.restype = wt.BOOL
user32.TrackPopupMenu.argtypes = [
    wt.HMENU, wt.UINT, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    wt.HWND, ctypes.c_void_p,
]
user32.TrackPopupMenu.restype = wt.UINT
user32.DestroyMenu.argtypes = [wt.HMENU]
user32.DestroyMenu.restype = wt.BOOL
user32.LoadImageW.argtypes = [
    wt.HINSTANCE, wt.LPCWSTR, wt.UINT, ctypes.c_int, ctypes.c_int,
    wt.UINT,
]
user32.LoadImageW.restype = wt.HANDLE
user32.LoadIconW.argtypes = [wt.HINSTANCE, ctypes.c_void_p]
user32.LoadIconW.restype = wt.HICON
user32.GetMessageW.argtypes = [
    ctypes.POINTER(wt.MSG), wt.HWND, wt.UINT, wt.UINT]
user32.GetMessageW.restype = ctypes.c_int
user32.TranslateMessage.argtypes = [ctypes.POINTER(wt.MSG)]
user32.TranslateMessage.restype = wt.BOOL
user32.DispatchMessageW.argtypes = [ctypes.POINTER(wt.MSG)]
user32.DispatchMessageW.restype = LRESULT
user32.PostQuitMessage.argtypes = [ctypes.c_int]
kernel32.GetModuleHandleW.argtypes = [wt.LPCWSTR]
kernel32.GetModuleHandleW.restype = wt.HMODULE
shell32.Shell_NotifyIconW.argtypes = [wt.DWORD, ctypes.POINTER(NOTIFYICONDATAW)]
shell32.Shell_NotifyIconW.restype = wt.BOOL


class TrayIcon:
    """A small notification-area icon with a native popup menu."""

    def __init__(self, icon_path: str, tooltip: str,
                 on_show: Callable[[], None],
                 on_capture: Callable[[], None],
                 on_open_data: Callable[[], None],
                 on_exit: Callable[[], None]):
        self.icon_path = icon_path
        self.tooltip = tooltip[:127]
        self.on_show = on_show
        self.on_capture = on_capture
        self.on_open_data = on_open_data
        self.on_exit = on_exit
        self.hwnd = 0
        self.hicon = 0
        self.error: Optional[str] = None
        self._ready = threading.Event()
        self._stop_requested = False
        self._thread = threading.Thread(target=self._run, daemon=True,
                                        name="tray-icon")
        self._wndproc_ref = WNDPROC(self._wndproc)
        self._nid: Optional[NOTIFYICONDATAW] = None
        self._taskbar_created = user32.RegisterWindowMessageW("TaskbarCreated")

    def start(self, timeout: float = 3.0) -> bool:
        self._thread.start()
        self._ready.wait(timeout)
        if self.error:
            log.warning("tray icon unavailable: %s", self.error)
            return False
        return True

    def stop(self):
        self._stop_requested = True
        if self.hwnd:
            user32.PostMessageW(self.hwnd, WM_CLOSE, 0, 0)
        if self._thread.is_alive():
            self._thread.join(timeout=2.0)

    def set_tooltip(self, text: str):
        self.tooltip = (text or "")[:127]
        if self.hwnd and self._nid is not None:
            self._nid.szTip = self.tooltip
            shell32.Shell_NotifyIconW(NIM_MODIFY, ctypes.byref(self._nid))

    # --- tray thread -------------------------------------------------------
    def _run(self):
        class_name = "WorkpaperTodoTrayWindow"
        hinst = kernel32.GetModuleHandleW(None)
        wc = WNDCLASSW()
        wc.style = CS_DBLCLKS
        wc.lpfnWndProc = self._wndproc_ref
        wc.hInstance = hinst
        wc.lpszClassName = class_name
        atom = user32.RegisterClassW(ctypes.byref(wc))
        if not atom:
            self.error = "RegisterClassW failed"
            self._ready.set()
            return
        try:
            self.hwnd = int(user32.CreateWindowExW(
                0, class_name, "WorkpaperTodoTray", 0, 0, 0, 0, 0,
                None, None, hinst, None) or 0)
            if not self.hwnd:
                self.error = "CreateWindowExW failed"
                return
            self.hicon = self._load_icon()
            if not self._add_icon():
                self.error = "Shell_NotifyIcon(NIM_ADD) failed"
                return
            self._ready.set()
            msg = wt.MSG()
            while not self._stop_requested:
                ret = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
                if ret == 0 or ret == -1:
                    break
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
        except Exception as exc:
            self.error = repr(exc)
            log.exception("tray thread failed")
        finally:
            if self._nid is not None:
                shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(self._nid))
                self._nid = None
            if self.hwnd:
                user32.DestroyWindow(self.hwnd)
                self.hwnd = 0
            user32.UnregisterClassW(class_name, hinst)
            self._ready.set()

    def _load_icon(self):
        icon = 0
        if self.icon_path:
            icon = int(user32.LoadImageW(
                None, self.icon_path, IMAGE_ICON, 0, 0,
                LR_LOADFROMFILE | LR_DEFAULTSIZE) or 0)
        if not icon:
            icon = int(user32.LoadIconW(
                None, ctypes.c_void_p(IDI_APPLICATION)) or 0)
        return icon

    def _add_icon(self) -> bool:
        nid = NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        nid.hWnd = self.hwnd
        nid.uID = 1
        nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP
        nid.uCallbackMessage = WM_TRAY
        nid.hIcon = self.hicon
        nid.szTip = self.tooltip
        self._nid = nid
        return bool(shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(nid)))

    def _wndproc(self, hwnd, msg, wparam, lparam):
        try:
            if msg == WM_TRAY:
                event = int(lparam) & 0xFFFF
                if event in (WM_LBUTTONUP, WM_LBUTTONDBLCLK):
                    self._invoke(self.on_show)
                elif event in (WM_RBUTTONUP, WM_CONTEXTMENU):
                    self._show_menu()
            elif msg == self._taskbar_created:
                self._add_icon()
            elif msg == WM_CLOSE:
                user32.DestroyWindow(hwnd)
            elif msg == WM_DESTROY:
                user32.PostQuitMessage(0)
        except Exception:
            log.exception("tray message handling failed")
        return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    def _show_menu(self):
        menu = user32.CreatePopupMenu()
        if not menu:
            return
        try:
            user32.AppendMenuW(menu, MF_STRING, ID_SHOW, "显示主窗口")
            user32.AppendMenuW(menu, MF_STRING, ID_CAPTURE, "登记最近单元格")
            user32.AppendMenuW(menu, MF_SEPARATOR, 0, None)
            user32.AppendMenuW(menu, MF_STRING, ID_OPEN_DATA, "打开数据文件夹")
            user32.AppendMenuW(menu, MF_STRING, ID_EXIT, "退出")
            point = wt.POINT()
            user32.GetCursorPos(ctypes.byref(point))
            user32.SetForegroundWindow(self.hwnd)
            command = int(user32.TrackPopupMenu(
                menu, TPM_RETURNCMD | TPM_RIGHTBUTTON,
                point.x, point.y, 0, self.hwnd, None))
            user32.PostMessageW(self.hwnd, WM_NULL, 0, 0)
        finally:
            user32.DestroyMenu(menu)
        if command == ID_SHOW:
            self._invoke(self.on_show)
        elif command == ID_CAPTURE:
            self._invoke(self.on_capture)
        elif command == ID_OPEN_DATA:
            self._invoke(self.on_open_data)
        elif command == ID_EXIT:
            self._invoke(self.on_exit)

    @staticmethod
    def _invoke(callback):
        try:
            callback()
        except Exception:
            log.exception("tray callback failed")
