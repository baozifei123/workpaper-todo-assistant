# -*- coding: utf-8 -*-
"""Small real-Tk checks for onboarding and relink wizard interactions."""
import os
import tkinter as tk
import time
from types import SimpleNamespace
from unittest.mock import Mock

from workpaper_todo import app as app_module, config, dialogs, store, win32util
from workpaper_todo.models import Item


def create_root(attempts=3):
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


def test_first_run_guide_exposes_steps_and_completes(tmp_path, monkeypatch):
    root = create_root()
    app = SimpleNamespace(
        root=root, scale=1.0, data_dir=str(tmp_path), hotkey_text="Ctrl+Alt+M",
        hotkey_status_text=lambda: "全局热键 Ctrl+Alt+M 已成功注册。",
        open_practice_workbook=lambda: str(tmp_path / "practice.xlsx"),
        dismiss_onboarding=Mock())
    dialog = None
    try:
        dialog = dialogs.FirstRunDialog(app)
        root.update()
        assert "1 / 3" in dialog.page_label.cget("text")
        content_texts = []
        for widget in dialog.content.winfo_children():
            if isinstance(widget, tk.Text):
                content_texts.append(widget.get("1.0", "end"))
            elif isinstance(widget, tk.Label):
                content_texts.append(widget.cget("text"))
        assert app.data_dir in " ".join(content_texts)

        dialog._next()
        root.update()
        assert "2 / 3" in dialog.page_label.cget("text")
        dialog._open_practice()
        assert "practice.xlsx" in dialog.path_var.get()
        dialog._next()
        assert "3 / 3" in dialog.page_label.cget("text")
        dialog._complete()
        app.dismiss_onboarding.assert_called_once_with()
    finally:
        if dialog is not None:
            try:
                dialog.destroy()
            except tk.TclError:
                pass
        root.destroy()


def test_counts_label_has_its_own_row_inside_window(
        tmp_path, monkeypatch):
    monkeypatch.setattr(app_module.App, "_start_hotkey", lambda *a, **k: None)
    monkeypatch.setattr(app_module.App, "_start_tray", lambda *a, **k: None)
    monkeypatch.setattr(app_module.App, "_track_foreground", lambda *a, **k: None)
    monkeypatch.setattr(app_module.App, "_auto_backup", lambda *a, **k: None)
    data_dir = str(tmp_path / "data")
    cfg = config.Config(data_dir)
    cfg.set("onboarding_completed", True)
    cfg.set("close_to_tray", False)
    db = store.Store(os.path.join(data_dir, "data.db"))
    root = create_root()
    gui = None
    try:
        gui = app_module.App(root, db, cfg, data_dir)
        root.update()
        assert gui.counts_label.winfo_ismapped()
        right = gui.counts_label.winfo_rootx() + gui.counts_label.winfo_width()
        assert right <= root.winfo_rootx() + root.winfo_width()
        # The counts must not share the filter row: at 125%/150% DPI the
        # controls use the full width and clip anything packed to the right.
        assert gui.counts_label.master is not gui.quick_view_box.master
        assert "未完成" in gui.counts_label.cget("text")
    finally:
        if gui is not None:
            gui.on_close(real_exit=True)
        else:
            root.destroy()


def test_relink_wizard_requires_explicit_candidate_confirmation(
        tmp_path, monkeypatch):
    target = tmp_path / "synthetic_candidate.xlsx"
    target.write_text("synthetic test file", encoding="utf-8")
    monkeypatch.setattr(dialogs.filedialog, "askopenfilename",
                        lambda **_kwargs: str(target))
    monkeypatch.setattr(win32util, "file_identity",
                        lambda _path: ("vol-new", "index-new"))
    root = create_root()
    app = SimpleNamespace(root=root, scale=1.0)
    item = Item(
        id=1, file_path=r"E:\\missing\\old.xlsx", file_name="old.xlsx",
        sheet_name="练习底稿", cell_address="B5", snapshot="【练习】原登记快照",
        file_volume="vol-old", file_index="index-old")
    dialog = None
    try:
        dialog = dialogs.RelinkDialog(app, item)
        root.update()
        assert str(dialog.continue_button.cget("state")) == tk.DISABLED
        dialog._choose_file()
        assert dialog.target_path_var.get() == str(target)
        assert "指纹不同" in dialog.identity_var.get()
        assert str(dialog.continue_button.cget("state")) == tk.NORMAL
        dialog._continue()
        assert dialog.result == str(target)
    finally:
        if dialog is not None:
            try:
                dialog.destroy()
            except tk.TclError:
                pass
        root.destroy()
