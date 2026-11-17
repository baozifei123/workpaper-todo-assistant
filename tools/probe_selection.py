# -*- coding: utf-8 -*-
"""通过 NativeOM 连接窗口，测试选区读取细节（只改选区，不改内容）。"""
import ctypes
import ctypes.wintypes as wt
import sys

import comtypes
import comtypes.automation
import comtypes.client.dynamic

oleacc = ctypes.windll.oleacc
user32 = ctypes.windll.user32
OBJID_NATIVEOM = -16

CB = ctypes.WINFUNCTYPE(ctypes.c_bool, wt.HWND, wt.LPARAM)


def cls(h):
    b = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(h, b, 256)
    return b.value


def pid(h):
    p = wt.DWORD()
    user32.GetWindowThreadProcessId(h, ctypes.byref(p))
    return p.value


def find_xlmain(pid_value):
    found = []

    def cb(h, l):
        if pid(h) == pid_value and cls(h) == "XLMAIN":
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


def get_app(hwnd):
    p = ctypes.POINTER(comtypes.automation.IDispatch)()
    hr = oleacc.AccessibleObjectFromWindow(
        wt.HWND(hwnd), wt.DWORD(OBJID_NATIVEOM & 0xFFFFFFFF),
        ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(p))
    if hr != 0:
        raise RuntimeError(hex(hr))
    win = comtypes.client.dynamic.Dispatch(p)
    return win.Application


def describe(sel):
    out = {}
    for name in ("Address", "Row", "Column", "Text", "Value", "Formula"):
        try:
            v = getattr(sel, name)
            out[name] = v
        except Exception as e:
            out[name] = f"ERR {e!r}"
    for name in ("Cells", "Rows", "Columns"):
        try:
            obj = getattr(sel, name)
            out[name + ".Count"] = obj.Count
        except Exception as e:
            out[name + ".Count"] = f"ERR {e!r}"
    try:
        out["MergeCells"] = sel.MergeCells
    except Exception as e:
        out["MergeCells"] = f"ERR {e!r}"
    try:
        ma = sel.MergeArea
        out["MergeArea.Row"] = ma.Row
        out["MergeArea.Column"] = ma.Column
        out["MergeArea.Rows.Count"] = ma.Rows.Count
        out["MergeArea.Columns.Count"] = ma.Columns.Count
    except Exception as e:
        out["MergeArea"] = f"ERR {e!r}"
    for k, v in out.items():
        print(f"    {k} = {v!r} (type={type(v).__name__})")
    addr = out.get("Address")
    if type(addr).__name__ == "NamedProperty":
        for attempt in ("call", "index"):
            try:
                print(f"    Address({attempt}) = {addr() if attempt == 'call' else addr(1, 1)!r}")
            except Exception as e:
                print(f"    Address({attempt}) ERR {e!r}")


def main():
    pid_value = int(sys.argv[1]) if len(sys.argv) > 1 else None
    mains = find_xlmain(pid_value)
    if not mains:
        print("no XLMAIN for pid", pid_value)
        return 1
    target = None
    app = None
    for m in mains:
        excel7 = find_child(m, "EXCEL7")
        if not excel7:
            continue
        for x7 in excel7:
            try:
                app = get_app(x7)
                print("app:", app.Name, getattr(app, "Path", "n/a"))
                wb = app.ActiveWorkbook
                print("connected:", wb.FullName, "| sheet:", app.ActiveSheet.Name)
                target = x7
                break
            except Exception as e:
                import traceback
                print("skip", hex(x7), repr(e))
                traceback.print_exc()
        if target is not None:
            break
    if target is None:
        return 2
    ws = app.ActiveSheet
    original = app.Selection
    print("original selection:", original.Text)
    try:
        for addr in ("A5", "A9", "A5:B6"):
            print(f"== select {addr}")
            app.Range(addr).Select()
            describe(app.Selection)
    finally:
        try:
            original.Select()
        except Exception as e:
            print("restore failed", e)
    return 0


if __name__ == "__main__":
    sys.exit(main())
