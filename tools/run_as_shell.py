# -*- coding: utf-8 -*-
"""用当前登录用户 shell(explorer.exe)的令牌启动子进程，用于模拟普通权限运行。

用法: python run_as_shell.py "<命令行>" [工作目录]
仅用于测试：让本工具能在管理员终端中验证普通用户场景。
"""
import ctypes
import ctypes.wintypes as wt
import os
import sys

import win32api
import win32con
import win32process
import win32security

advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

LOGON_WITH_PROFILE = 0x00000001
CREATE_NO_WINDOW = 0x08000000
INFINITE = 0xFFFFFFFF


class STARTUPINFOW(ctypes.Structure):
    _fields_ = [
        ("cb", wt.DWORD),
        ("lpReserved", wt.LPWSTR),
        ("lpDesktop", wt.LPWSTR),
        ("lpTitle", wt.LPWSTR),
        ("dwX", wt.DWORD),
        ("dwY", wt.DWORD),
        ("dwXSize", wt.DWORD),
        ("dwYSize", wt.DWORD),
        ("dwXCountChars", wt.DWORD),
        ("dwYCountChars", wt.DWORD),
        ("dwFillAttribute", wt.DWORD),
        ("dwFlags", wt.DWORD),
        ("wShowWindow", wt.WORD),
        ("cbReserved2", wt.WORD),
        ("lpReserved2", ctypes.POINTER(ctypes.c_byte)),
        ("hStdInput", wt.HANDLE),
        ("hStdOutput", wt.HANDLE),
        ("hStdError", wt.HANDLE),
    ]


class PROCESS_INFORMATION(ctypes.Structure):
    _fields_ = [
        ("hProcess", wt.HANDLE),
        ("hThread", wt.HANDLE),
        ("dwProcessId", wt.DWORD),
        ("dwThreadId", wt.DWORD),
    ]


def find_explorer_pid():
    for pid in win32process.EnumProcesses():
        try:
            h = win32api.OpenProcess(
                win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ,
                False, pid)
            try:
                exe = win32process.GetModuleFileNameEx(h, 0)
            finally:
                win32api.CloseHandle(h)
            if os.path.basename(exe).lower() == "explorer.exe":
                return pid
        except Exception:
            continue
    return None


def main():
    if len(sys.argv) < 2:
        print("usage: run_as_shell.py <command line> [cwd]")
        return 2
    cmdline = sys.argv[1]
    cwd = sys.argv[2] if len(sys.argv) > 2 else os.getcwd()

    pid = find_explorer_pid()
    if not pid:
        print("explorer.exe not found")
        return 3

    ph = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION, False, pid)
    try:
        token = win32security.OpenProcessToken(
            ph, win32con.TOKEN_DUPLICATE | win32con.TOKEN_ASSIGN_PRIMARY | win32con.TOKEN_QUERY)
        try:
            dup = win32security.DuplicateTokenEx(
                token,
                win32security.SecurityImpersonation,
                win32con.TOKEN_ALL_ACCESS,
                win32security.TokenPrimary,
                None)
            si = STARTUPINFOW()
            si.cb = ctypes.sizeof(STARTUPINFOW)
            si.dwFlags = 0x00000001  # STARTF_USESHOWWINDOW
            si.wShowWindow = win32con.SW_HIDE
            pi = PROCESS_INFORMATION()
            buf = ctypes.create_unicode_buffer(cmdline)
            ok = advapi32.CreateProcessWithTokenW(
                wt.HANDLE(int(dup)), LOGON_WITH_PROFILE, None, buf,
                CREATE_NO_WINDOW, None, cwd,
                ctypes.byref(si), ctypes.byref(pi))
            if not ok:
                err = ctypes.get_last_error()
                print("CreateProcessWithTokenW failed", err,
                      ctypes.FormatError(err))
                return 4
            print("started pid", pi.dwProcessId)
            kernel32.WaitForSingleObject(pi.hProcess, INFINITE)
            code = wt.DWORD()
            kernel32.GetExitCodeProcess(pi.hProcess, ctypes.byref(code))
            print("exit", code.value)
            win32api.CloseHandle(pi.hProcess)
            win32api.CloseHandle(pi.hThread)
            return 0
        finally:
            win32api.CloseHandle(token)
    finally:
        win32api.CloseHandle(ph)


if __name__ == "__main__":
    sys.exit(main())
