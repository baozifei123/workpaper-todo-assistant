# -*- coding: utf-8 -*-
"""探测 WPS/Excel 已运行实例的可连接方式（只读，不修改任何文档）。"""
import ctypes
import ctypes.wintypes as wt
import subprocess
import sys
import os

USER32 = ctypes.windll.user32
OLEACC = ctypes.windll.oleacc

OBJID_NATIVEOM = -16  # 0xFFFFFFF0


def _cls(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    USER32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def _title(hwnd):
    buf = ctypes.create_unicode_buffer(512)
    USER32.GetWindowTextW(hwnd, buf, 512)
    return buf.value


def _pid(hwnd):
    p = wt.DWORD()
    USER32.GetWindowThreadProcessId(hwnd, ctypes.byref(p))
    return p.value


def list_windows(pid):
    """Return list of (hwnd, class, title, is_child) for a process."""
    found = []
    CB = ctypes.WINFUNCTYPE(ctypes.c_bool, wt.HWND, wt.LPARAM)

    def cb(hwnd, lparam):
        if _pid(hwnd) == pid:
            found.append((hwnd, _cls(hwnd), _title(hwnd)))
        return True

    USER32.EnumWindows(CB(cb), 0)
    children = []
    for hwnd, _, _ in list(found):
        def ccb(ch, lp):
            if _pid(ch) == pid:
                children.append((ch, _cls(ch), _title(ch)))
            return True
        USER32.EnumChildWindows(hwnd, CB(ccb), 0)
    return found, children


def native_om(hwnd):
    """Try AccessibleObjectFromWindow(OBJID_NATIVEOM) and describe result."""
    import pythoncom
    from win32com.client import Dispatch
    ptr = ctypes.POINTER(pythoncom.IUnknown)()
    iid = pythoncom.IID_IDispatch
    hr = OLEACC.AccessibleObjectFromWindow(
        wt.HWND(hwnd), ctypes.c_long(OBJID_NATIVEOM),
        ctypes.byref(iid), ctypes.byref(ptr))
    if hr != 0:
        return f"hr=0x{hr & 0xFFFFFFFF:08X}"
    try:
        obj = Dispatch(ptr)
        info = []
        for attr in ("Name", "Version", "Path", "Hwnd"):
            try:
                info.append(f"{attr}={getattr(obj, attr)!r}")
            except Exception:
                pass
        try:
            app = obj.Application
            info.append(f"App.Name={app.Name!r}")
            info.append(f"App.Path={app.Path!r}")
            try:
                wb = app.ActiveWorkbook
                info.append(f"ActiveWorkbook={wb.FullName!r}")
                info.append(f"ActiveSheet={app.ActiveSheet.Name!r}")
                sel = app.Selection
                info.append(f"Selection={sel.Address!r} Text={sel.Text!r}")
            except Exception as e:
                info.append(f"workbook_err={e!r}")
        except Exception as e:
            info.append(f"app_err={e!r}")
        return "OK " + " | ".join(info)
    except Exception as e:
        return f"Dispatch failed: {e!r}"


def child_approach(name):
    import pythoncom
    pythoncom.CoInitialize()
    try:
        if name == "getactive_ket":
            import win32com.client as w
            app = w.GetActiveObject("Ket.Application")
            return f"OK Name={app.Name!r} Version={app.Version!r} Path={app.Path!r}"
        if name == "dispatch_ket":
            import win32com.client as w
            app = w.Dispatch("Ket.Application")
            return f"OK Name={app.Name!r} Path={app.Path!r}"
        if name == "getactive_clsid":
            import win32com.client as w
            app = w.GetActiveObject("{45540001-5750-5300-4B49-4E47534F4655}")
            return f"OK Name={app.Name!r} Path={app.Path!r}"
        if name == "excel_getactive":
            import win32com.client as w
            app = w.GetActiveObject("Excel.Application")
            return f"OK Name={app.Name!r} Path={app.Path!r}"
    except Exception as e:
        return f"FAIL {e!r}"
    finally:
        pythoncom.CoUninitialize()
    return "UNKNOWN"


def run_child(name, timeout=12):
    code = (
        "import sys; sys.path.insert(0, r'%s'); "
        "import probe_wps_com; print(probe_wps_com.child_approach('%s'))"
    ) % (os.path.dirname(os.path.abspath(__file__)), name)
    try:
        r = subprocess.run([sys.executable, "-X", "utf8", "-c", code],
                           capture_output=True, timeout=timeout)
        out = r.stdout.decode("utf-8", "replace").strip()
        err = r.stderr.decode("utf-8", "replace").strip()
        return (out or "<no stdout>") + ((" | stderr=" + err) if err else "")
    except subprocess.TimeoutExpired:
        return "TIMEOUT"
    except Exception as e:
        return f"RUN_FAIL {e!r}"


def main():
    if len(sys.argv) > 1:
        print(child_approach(sys.argv[1]))
        return
    import win32process
    print("== COM approaches ==")
    for name in ("getactive_ket", "dispatch_ket", "getactive_clsid", "excel_getactive"):
        print(f"{name}: {run_child(name)}")

    print("\n== window classes ==")
    for proc_name, pid in _find_pids():
        tops, children = list_windows(pid)
        print(f"-- {proc_name} pid={pid}")
        for h, c, t in tops[:5]:
            print(f"  TOP {hex(h)} {c!r} {t[:50]!r} nativeom={native_om(h) if c else ''}")
        seen = set()
        for h, c, t in children[:80]:
            if c in seen:
                continue
            seen.add(c)
            print(f"  CHILD {hex(h)} {c!r} {t[:50]!r} nativeom={native_om(h) if c else ''}")


def _find_pids():
    import win32process
    import win32api
    out = []
    for pname in ("et.exe", "EXCEL.EXE"):
        for proc in _iter_procs(pname):
            out.append((pname, proc))
    return out


def _iter_procs(name):
    import win32process
    import win32api
    import win32con
    pids = []
    for pid in win32process.EnumProcesses():
        try:
            h = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
            try:
                exe = win32process.GetModuleFileNameEx(h, 0)
            finally:
                win32api.CloseHandle(h)
            if os.path.basename(exe).lower() == name.lower():
                pids.append(pid)
        except Exception:
            pass
    return pids


if __name__ == "__main__":
    main()
