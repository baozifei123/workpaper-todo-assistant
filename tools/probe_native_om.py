# -*- coding: utf-8 -*-
"""从窗口句柄获取 WPS/Excel 原生对象模型（OBJID_NATIVEOM），只读探测。"""
import ctypes
import ctypes.wintypes as wt
import sys

oleacc = ctypes.windll.oleacc
user32 = ctypes.windll.user32
OBJID_NATIVEOM = -16


def cls(h):
    b = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(h, b, 256)
    return b.value


def title(h):
    b = ctypes.create_unicode_buffer(512)
    user32.GetWindowTextW(h, b, 512)
    return b.value


def pid(h):
    p = wt.DWORD()
    user32.GetWindowThreadProcessId(h, ctypes.byref(p))
    return p.value


CB = ctypes.WINFUNCTYPE(ctypes.c_bool, wt.HWND, wt.LPARAM)


def find_main(pid_value, class_name):
    found = []

    def cb(h, l):
        if pid(h) == pid_value and cls(h) == class_name:
            found.append(h)
        return True
    user32.EnumWindows(CB(cb), 0)
    return found


def find_child(parent, class_name):
    found = []

    def cb(h, l):
        if cls(h) == class_name:
            found.append(h)
        return True
    user32.EnumChildWindows(parent, CB(cb), 0)
    return found


def native_om(hwnd):
    import comtypes
    import comtypes.automation
    import comtypes.client.dynamic
    p = ctypes.POINTER(comtypes.automation.IDispatch)()
    hr = oleacc.AccessibleObjectFromWindow(
        wt.HWND(hwnd), wt.DWORD(OBJID_NATIVEOM & 0xFFFFFFFF),
        ctypes.byref(comtypes.automation.IDispatch._iid_),
        ctypes.byref(p))
    if hr != 0:
        return f"AccessibleObjectFromWindow hr=0x{hr & 0xFFFFFFFF:08X}"
    out = []
    try:
        disp = comtypes.client.dynamic.Dispatch(p)
        for attr in ("Name", "Version", "Path", "Hwnd"):
            try:
                out.append(f"{attr}={getattr(disp, attr)!r}")
            except Exception as e:
                out.append(f"{attr}=ERR({e!r})")
        app = None
        for attr in ("Application", "Parent"):
            try:
                candidate = getattr(disp, attr)
                out.append(f"via {attr}: Name={candidate.Name!r} Path={getattr(candidate, 'Path', 'n/a')!r}")
                app = candidate
                break
            except Exception as e:
                out.append(f"{attr}=ERR({e!r})")
        if app is None:
            app = disp
        try:
            wb = app.ActiveWorkbook
            out.append(f"ActiveWorkbook={wb.FullName!r}")
            out.append(f"ActiveSheet={app.ActiveSheet.Name!r}")
            sel = app.Selection
            out.append(f"Selection.Address={sel.Address!r}")
            out.append(f"Selection.Text={sel.Text!r}")
            try:
                out.append(f"MergeCells={sel.MergeCells!r}")
            except Exception as e:
                out.append(f"MergeCells=ERR({e!r})")
        except Exception as e:
            out.append(f"workbook=ERR({e!r})")
    except Exception as e:
        return f"dynamic.Dispatch failed: {e!r}"
    return "OK " + " | ".join(out)


def main():
    targets = []
    for pid_value, label in [(int(x), y) for x, y in
                             (a.split(":", 1) for a in sys.argv[1:])]:
        targets.append((pid_value, label))
    if not targets:
        print("usage: probe_native_om.py <pid:label> ...")
        return
    for pid_value, label in targets:
        for main_h in find_main(pid_value, "XLMAIN"):
            print(f"== {label} pid={pid_value} XLMAIN={hex(main_h)} title={title(main_h)!r}")
            x7 = find_child(main_h, "EXCEL7")
            print("  EXCEL7:", [hex(x) for x in x7])
            for h in x7:
                print("  result:", native_om(h))


if __name__ == "__main__":
    main()
