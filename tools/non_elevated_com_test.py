# -*- coding: utf-8 -*-
"""在普通（非管理员）权限下测试 WPS/Excel COM 连接。结果写入文件。"""
import os
import sys
import traceback

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "non_elevated_result.txt")


def log(msg=""):
    print(msg)
    with open(OUT, "a", encoding="utf-8") as f:
        f.write(str(msg) + "\n")


def main():
    open(OUT, "w", encoding="utf-8").close()
    try:
        import ctypes
        log("IsUserAdmin=%s" % bool(ctypes.windll.shell32.IsUserAnAdmin()))
        log("Python=" + sys.executable)
        import pythoncom
        pythoncom.CoInitialize()
        import win32com.client as w

        def describe(app, label):
            try:
                parts = [label, "Name=%r" % app.Name, "Version=%r" % app.Version,
                         "Path=%r" % app.Path]
                try:
                    parts.append("Hwnd=%r" % app.Hwnd)
                except Exception:
                    pass
                try:
                    wb = app.ActiveWorkbook
                    parts.append("ActiveWorkbook=%r" % wb.FullName)
                    parts.append("ActiveSheet=%r" % app.ActiveSheet.Name)
                    sel = app.Selection
                    parts.append("Selection=%r" % sel.Address)
                    parts.append("Text=%r" % sel.Text)
                except Exception as e:
                    parts.append("workbook_err=%r" % e)
                log(" | ".join(parts))
            except Exception as e:
                log(label + " describe_err=%r" % e)

        for label, prog in [
            ("GetActiveObject Ket", "Ket.Application"),
            ("Dispatch Ket", "Ket.Application"),
            ("GetActiveObject CLSID", "{45540001-5750-5300-4B49-4E47534F4655}"),
            ("GetActiveObject Excel", "Excel.Application"),
        ]:
            try:
                if label.startswith("GetActiveObject"):
                    app = w.GetActiveObject(prog)
                else:
                    app = w.Dispatch(prog)
                describe(app, label + " OK")
            except Exception as e:
                log(label + " FAIL %r" % e)
        pythoncom.CoUninitialize()
    except Exception:
        log("FATAL\n" + traceback.format_exc())


if __name__ == "__main__":
    main()
