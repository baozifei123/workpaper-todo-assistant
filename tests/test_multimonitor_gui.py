# -*- coding: utf-8 -*-
"""Real multi-monitor window-placement test (skipped on single-monitor hosts)."""
import os
import time
import tkinter as tk
from types import SimpleNamespace

import pytest

from workpaper_todo import app, config, store, win32util
from workpaper_todo.dialogs import BaseDialog


def create_root(attempts=3):
    """Tolerate the known transient Tcl library read failure on repeated roots."""
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


def test_windows_stay_on_secondary_monitor(tmp_path, monkeypatch):
    areas = win32util.list_monitor_work_areas()
    if len(areas) < 2:
        pytest.skip("当前主机只有一个显示器")
    secondary = next(area for area in areas if not area[4])
    rect = (secondary[0] + 50, secondary[1] + 50, 800, 600)

    monkeypatch.setattr(app.App, "_start_hotkey", lambda *a, **k: None)
    monkeypatch.setattr(app.App, "_start_tray", lambda *a, **k: None)
    monkeypatch.setattr(app.App, "_track_foreground", lambda *a, **k: None)
    monkeypatch.setattr(app.App, "_auto_backup", lambda *a, **k: None)

    data_dir = str(tmp_path / "data")
    cfg = config.Config(data_dir)
    cfg.set("onboarding_completed", True)
    cfg.set("close_to_tray", False)
    cfg.set("window_rect", list(rect))
    db = store.Store(os.path.join(data_dir, "data.db"))
    root = create_root()
    gui = None
    dialog = None
    try:
        gui = app.App(root, db, cfg, data_dir)
        root.update()
        area = win32util.monitor_work_area_for_window(root.winfo_id())
        assert area == secondary[:4]
        x, y, w, h = (root.winfo_x(), root.winfo_y(),
                      root.winfo_width(), root.winfo_height())
        assert area[0] <= x and area[1] <= y
        assert x + w <= area[2] and y + h <= area[3]

        dialog = BaseDialog(SimpleNamespace(root=root), "测试对话框", 600, 400)
        dialog.update()
        dialog_area = win32util.monitor_work_area_for_window(dialog.winfo_id())
        assert dialog_area == secondary[:4]
        dx, dy, dw, dh = (dialog.winfo_x(), dialog.winfo_y(),
                          dialog.winfo_width(), dialog.winfo_height())
        assert dialog_area[0] <= dx and dialog_area[1] <= dy
        assert dx + dw <= dialog_area[2] and dy + dh <= dialog_area[3]
    finally:
        if dialog is not None:
            dialog.destroy()
        if gui is not None:
            gui.on_close(real_exit=True)
        else:
            root.destroy()
