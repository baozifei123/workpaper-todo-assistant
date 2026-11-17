# -*- coding: utf-8 -*-
"""Real Excel/WPS + Tk capture-button integration (explicit opt-in).

The workbook and database are generated in pytest's temporary directory.
Foreground-window tracking is simulated to avoid sending keys to another app;
the actual Tk button, worker thread, NativeOM, and SQLite storage are exercised.
"""
import os
import subprocess
import time
import tkinter as tk

import comtypes
from openpyxl import Workbook
import pytest

from workpaper_todo import app, config, jump, nativeom, store, win32util
from workpaper_todo.models import Item

pytestmark = pytest.mark.skipif(
    os.environ.get("WPT_RUN_INTEGRATION") != "1",
    reason="设置 WPT_RUN_INTEGRATION=1 才运行真实软件联动测试")


@pytest.fixture(scope="module", autouse=True)
def com_apartment():
    """Keep COM initialized across both Excel and WPS button cases.

    Initializing/uninitializing COM around each Tk test can leave COM proxies
    alive for a moment after CoUninitialize, which raises a fatal RPC error in
    the next case.  One apartment per module is both simpler and safer.
    """
    comtypes.CoInitialize()
    yield
    try:
        comtypes.CoUninitialize()
    except Exception:
        pass


def create_tk_root(attempts=3):
    """Create a Tk root, tolerating a transient Tcl library read failure.

    This pytest process creates one Tk root per Excel/WPS case.  On Windows,
    the second root can occasionally fail while Tcl loads its library files
    after a spreadsheet COM test; the packaged app creates only one root and
    is verified separately.  Retry only the root creation, never the capture.
    """
    last_error = None
    for attempt in range(attempts):
        try:
            return tk.Tk()
        except tk.TclError as exc:
            last_error = exc
            if attempt == attempts - 1:
                raise
            time.sleep(1.0)
    raise last_error


@pytest.mark.parametrize("source", ["excel", "wps"])
def test_capture_button_with_real_workbook(source, tmp_path, monkeypatch):
    path = str(tmp_path / ("button_%s.xlsx" % source))
    book = Workbook()
    sheet = book.active
    sheet.title = "测试工作表"
    sheet["A1"] = "测试按钮登记"
    book.save(path)

    exe = jump._resolve_exe(source)
    if not exe:
        pytest.skip("未安装 %s" % source)
    subprocess.Popen([exe, path], cwd=str(tmp_path))
    item = Item(file_path=path, source_app=source)
    found = None
    deadline = time.monotonic() + 35
    while time.monotonic() < deadline:
        found = jump.find_open_workbook(item)
        if found and found[0].source_app == source:
            break
        time.sleep(0.25)
    assert found and found[0].source_app == source, "未能打开合成测试工作簿"
    _sw, wb = found
    root = None
    gui = None
    try:
        wb.Activate()
        wb.Worksheets.Item("测试工作表").Activate()
        wb.Worksheets.Item("测试工作表").Range("A1").Select()
        target = jump.find_xlmain_for_workbook(wb)
        assert target

        # The global shortcut/tray are not needed for this button test and
        # should not interfere with a user's or another test's resources.
        monkeypatch.setattr(app.App, "_start_hotkey", lambda *a, **kw: None)
        monkeypatch.setattr(app.App, "_start_tray", lambda *a, **kw: None)
        warnings = []
        monkeypatch.setattr(app.messagebox, "showwarning",
                            lambda title, message, **kw: warnings.append(message))
        monkeypatch.setattr(app.messagebox, "showerror",
                            lambda title, message, **kw: warnings.append(message))
        monkeypatch.setattr(win32util, "get_foreground_window", lambda: target)

        data_dir = str(tmp_path / "data")
        cfg = config.Config(data_dir)
        cfg.set("onboarding_completed", True)
        db = store.Store(os.path.join(data_dir, "data.db"))
        root = create_tk_root()
        gui = app.App(root, db, cfg, data_dir)
        gui._track_foreground()
        assert gui.last_sheet_hwnd == target
        gui.capture_btn.invoke()

        deadline = time.monotonic() + 20
        while time.monotonic() < deadline and not db.all_items() and not warnings:
            root.update()
            time.sleep(0.05)
        assert not warnings, "登记按钮提示错误：%s" % warnings
        rows = db.all_items()
        assert len(rows) == 1
        assert (rows[0].source_app, rows[0].sheet_name, rows[0].cell_address,
                rows[0].snapshot, rows[0].status) == (
                    source, "测试工作表", "A1", "测试按钮登记", "未完成")
        assert nativeom.path_key(rows[0].file_path) == nativeom.path_key(path)
    finally:
        if gui is not None:
            gui.on_close(real_exit=True)
        elif root is not None:
            root.destroy()
        if nativeom.path_key(nativeom.workbook_path(wb)) == nativeom.path_key(path):
            wb.Close(False)
