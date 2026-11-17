# -*- coding: utf-8 -*-
"""底稿未完成助手主程序。"""
from __future__ import annotations

import copy
import ctypes
import datetime as _dt
import glob
import logging
import logging.handlers
import os
import queue
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Callable, List, Optional, Tuple

import comtypes

from . import (
    __version__, __window_title__, capture, config, jump, nativeom, practice,
    secure, store, win32util,
)
from .dialogs import (
    DetailDialog, DuplicateDialog, FirstRunDialog, HistoryDialog, MismatchDialog,
    ProjectArchiveDialog, QuickFollowUpDialog, RelinkDialog, SettingsDialog,
    FONT, FONT_BOLD, FONT_SMALL,
)
from .models import (
    Item, SelectionInfo, STATUS_DONE, STATUS_DOING, STATUS_PRIORITY, STATUS_TODO,
    STATUS_WAITING, STATUSES, STATUS_TODO_FILTER, display_address, parse_date,
)
from .toast import Toast
from .tray import TrayIcon
from .win32util import (
    HotkeyError, HotkeyThread, bring_to_front, parse_hotkey,
)

log = logging.getLogger(__name__)

BG = "#F5F7FA"
CARD = "#FFFFFF"
HEADER = "#0F4C81"
TEXT = "#101828"
MUTED = "#667085"
BORDER = "#EAECF0"

STATUS_COLORS = {
    "未完成": ("#B42318", "#FEF3F2"),
    "等待资料": ("#B54708", "#FFFAEB"),
    "进行中": ("#175CD3", "#EFF8FF"),
    "已完成": ("#475467", "#F9FAFB"),
}

AGING_CHOICES = ["全部", "等待>7天", "等待>30天"]
_AGING_DAYS = {"等待>7天": 7, "等待>30天": 30}
QUICK_VIEW_CHOICES = ["待办", "待跟进", "无人负责", "等待>7天", "自定义"]
_SKIP_PROJECT_FOLDERS = {
    "temp", "tmp", "desktop", "downloads", "documents", "桌面", "下载",
    "文档", "c$", "d$", "e$",
}
_GENERIC_WORKPAPER_FOLDERS = {
    "底稿", "工作底稿", "审计底稿", "资料", "workpaper", "workpapers",
    "working papers", "working_paper",
}


def suggest_project(file_path: str, explicit: str, auto: bool = True) -> str:
    """Return an explicit project, a folder-based suggestion, or 未分类."""
    explicit = (explicit or "").strip()
    if explicit and explicit != "未分类":
        return explicit
    if auto and file_path:
        path = os.path.abspath(file_path)
        for _ in range(2):
            folder = os.path.basename(os.path.dirname(path))
            low = folder.lower()
            if not folder:
                break
            if low in _SKIP_PROJECT_FOLDERS:
                break
            if low not in _GENERIC_WORKPAPER_FOLDERS:
                return folder
            path = os.path.dirname(path)
    return "未分类"


def _redact_paths(text: str) -> str:
    """Remove absolute Windows/UNC paths from local log output."""
    text = re.sub(r"[A-Za-z]:\\[^\s\"'<>|]+", "<path>", text)
    text = re.sub(r"\\\\[^\s\"'<>|]+", "<path>", text)
    return text


class _RedactingFormatter(logging.Formatter):
    def format(self, record):
        return _redact_paths(super().format(record))


class Worker:
    """Single daemon worker: serializes COM reads and keeps the UI responsive."""

    def __init__(self, ui_queue: queue.Queue):
        self.queue = queue.Queue()
        self.ui_queue = ui_queue
        self.thread = threading.Thread(target=self._loop, daemon=True)
        self.thread.start()

    def submit(self, func, label, on_success, on_error):
        self.queue.put((func, label, on_success, on_error))

    def stop(self):
        self.queue.put(None)
        self.thread.join(timeout=2.0)

    def _loop(self):
        initialized_raw = False
        initialized_comtypes = False
        try:
            hr = ctypes.windll.ole32.CoInitializeEx(None, 0x2)  # STA
            initialized_raw = hr >= 0
            log.info("worker thread %s CoInitializeEx(STA) hr=0x%08X frozen=%s comtypes=%s",
                     threading.get_ident(), hr & 0xFFFFFFFF,
                     getattr(sys, "frozen", False), getattr(comtypes, "__version__", "?"))
        except Exception:
            log.exception("worker CoInitializeEx failed")
        try:
            comtypes.CoInitialize()
            initialized_comtypes = True
            log.info("worker comtypes.CoInitialize ok")
        except Exception:
            log.exception("worker comtypes.CoInitialize failed")
        try:
            while True:
                item = self.queue.get()
                if item is None:
                    break
                func, label, on_success, on_error = item
                try:
                    result = func()
                except Exception as exc:  # noqa: BLE001 - reported to UI
                    self.ui_queue.put(("task", label, None, exc, on_success, on_error))
                else:
                    self.ui_queue.put(("task", label, result, None, on_success, on_error))
        finally:
            if initialized_comtypes:
                try:
                    comtypes.CoUninitialize()
                except Exception:
                    pass
            if initialized_raw:
                try:
                    ctypes.windll.ole32.CoUninitialize()
                except Exception:
                    pass


class App:
    def __init__(self, root: tk.Tk, db: store.Store, cfg: config.Config,
                 data_dir: str):
        self.root = root
        self.store = db
        self.config = cfg
        self.data_dir = data_dir
        self.scale = max(1.0, root.winfo_fpixels("1i") / 96.0)
        self.toast_obj = Toast(root, self.scale)
        self.ui_queue: queue.Queue = queue.Queue()
        self.worker = Worker(self.ui_queue)
        self.busy = False
        self.items_by_id = {}
        self.hotkey_thread: Optional[HotkeyThread] = None
        self.hotkey_text = cfg.get("hotkey", config.DEFAULT_HOTKEY)
        self._search_after = None
        self._closing = False
        self.last_sheet_hwnd = 0
        self.tray: Optional[TrayIcon] = None
        self.close_to_tray = bool(cfg.get("close_to_tray", True))
        self._tray_hint_shown = bool(cfg.get("tray_hint_shown", False))

        self._setup_style()
        self._build_vars()
        self._build_ui()
        self._restore_window()
        self._start_hotkey(initial=True)
        self._start_tray()
        self.refresh()
        self.root.after(100, self._poll_queue)
        self.root.after(300, self._track_foreground)
        self.root.after(1200, self._auto_backup)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.report_callback_exception = self._report_exception

        self.root.bind_all("<Control-l>", self._accel_locate)
        self.root.bind_all("<Control-e>", self._accel_edit)
        self.root.bind_all("<Control-Shift-D>", self._accel_done)
        self.root.bind_all("<F5>", self._accel_refresh)
        self.root.bind_all("<Control-a>", self._accel_select_all)

        log.info("app started, version=%s, data_dir=%s, hotkey=%s",
                 __version__, data_dir, self.hotkey_text)
        if not self.config.get("onboarding_completed", False):
            self.root.after(700, self.open_onboarding)

    def _accel_locate(self, _event=None):
        if not self._dialog_open():
            self.locate_selected()
        return "break"

    def _accel_edit(self, _event=None):
        if not self._dialog_open():
            self.edit_selected()
        return "break"

    def _accel_done(self, _event=None):
        if not self._dialog_open():
            self.set_selected_status(STATUS_DONE)
        return "break"

    def _accel_refresh(self, _event=None):
        if not self._dialog_open():
            self.refresh()
        return "break"

    def _accel_select_all(self, _event=None):
        if self._dialog_open():
            return "break"
        children = self.tree.get_children()
        if children:
            self.tree.selection_set(*children)
            self.tree.focus(children[0])
            self._update_detail_panel()
        return "break"

    # ------------------------------------------------------------------ UI
    def _setup_style(self):
        style = ttk.Style(self.root)
        for theme in ("vista", "winnative", "clam"):
            try:
                style.theme_use(theme)
                break
            except Exception:
                continue
        row_h = max(26, int(30 * self.scale))
        style.configure("Treeview", font=("Microsoft YaHei UI", 10),
                        rowheight=row_h)
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"),
                        padding=(6, 6))
        style.configure("TButton", font=("Microsoft YaHei UI", 10), padding=(10, 5))
        style.configure("TCombobox", font=("Microsoft YaHei UI", 10))
        style.configure("TEntry", font=("Microsoft YaHei UI", 10))
        style.configure("TLabelframe.Label", font=("Microsoft YaHei UI", 10, "bold"))

    def _build_vars(self):
        self.project_var = tk.StringVar(value=self.config.get("current_project", ""))
        self.status_filter_var = tk.StringVar(
            value=self.config.get("filter_status", STATUS_TODO_FILTER))
        self.quick_view_var = tk.StringVar(
            value=self.config.get("quick_view", "自定义"))
        self.project_filter_var = tk.StringVar(
            value=self.config.get("filter_project", "全部"))
        self.aging_filter_var = tk.StringVar(
            value=self.config.get("filter_aging", "全部"))
        preset = self.quick_view_var.get()
        preset_filters = {
            "待办": (STATUS_TODO_FILTER, "全部"),
            "待跟进": (STATUS_WAITING, "全部"),
            "无人负责": (STATUS_TODO_FILTER, "全部"),
            "等待>7天": (STATUS_WAITING, "等待>7天"),
        }
        if preset in preset_filters:
            status, aging = preset_filters[preset]
            self.status_filter_var.set(status)
            self.aging_filter_var.set(aging)
        self.waiting_sort = self.config.get("sort_waiting", "")
        if self.waiting_sort not in ("asc", "desc"):
            self.waiting_sort = ""
        self.include_archived_var = tk.BooleanVar(
            value=bool(self.config.get("include_archived", False)))
        self.search_var = tk.StringVar(value=self.config.get("search_text", ""))

    def _build_ui(self):
        s = self.scale
        self.root.title(__window_title__)
        self.root.configure(bg=BG)

        header = tk.Frame(self.root, bg=HEADER)
        header.pack(fill="x")
        hinner = tk.Frame(header, bg=HEADER, padx=int(18 * s), pady=int(12 * s))
        hinner.pack(fill="x")
        tk.Label(hinner, text="底稿未完成事项", bg=HEADER, fg="white",
                 font=("Microsoft YaHei UI", 17, "bold")).pack(side="left")
        tk.Label(hinner, text="选中审计程序单元格 → 热键登记 → 清单里一键回到原位置",
                 bg=HEADER, fg="#D1E9FF", font=FONT_SMALL).pack(
                     side="left", padx=(16 * s, 0))
        right = tk.Frame(hinner, bg=HEADER)
        right.pack(side="right")
        self.hotkey_label = tk.Label(right, text="", bg=HEADER, fg="#D1E9FF",
                                     font=FONT_SMALL)
        self.hotkey_label.pack(side="left", padx=(0, 12 * s))
        tk.Button(right, text="设置", command=self.open_settings, bg=HEADER,
                  fg="white", activebackground="#0B3A63", activeforeground="white",
                  relief="solid", bd=1, font=FONT_SMALL, padx=10, pady=2,
                  cursor="hand2").pack(side="left")

        if not self.config.get("onboarding_completed", False):
            self.onboarding_banner = tk.Frame(
                self.root, bg="#EFF8FF", padx=int(18 * s), pady=int(7 * s))
            self.onboarding_banner.pack(fill="x")
            tk.Label(
                self.onboarding_banner,
                text="首次使用：用练习底稿试登记、定位和状态更新；不会触碰真实底稿。",
                bg="#EFF8FF", fg="#1849A9", font=FONT_SMALL,
                anchor="w").pack(side="left", fill="x", expand=True)
            tk.Button(
                self.onboarding_banner, text="开始引导",
                command=self.open_onboarding, bg="#EFF8FF", fg="#1570EF",
                relief="flat", bd=0, font=FONT_BOLD, cursor="hand2").pack(
                    side="right", padx=(8, 0))
            tk.Button(
                self.onboarding_banner, text="隐藏提示",
                command=self.dismiss_onboarding, bg="#EFF8FF", fg="#475467",
                relief="flat", bd=0, font=FONT_SMALL, cursor="hand2").pack(
                    side="right")
        else:
            self.onboarding_banner = None

        toolbar = tk.Frame(self.root, bg=CARD, padx=int(18 * s), pady=int(10 * s))
        toolbar.pack(fill="x")
        tk.Label(toolbar, text="登记到项目", bg=CARD, fg="#344054",
                 font=FONT_BOLD).pack(side="left")
        self.project_box = ttk.Combobox(
            toolbar, textvariable=self.project_var, width=22, font=FONT)
        self.project_box.pack(side="left", padx=(8, 16 * s))
        self.project_box.bind("<<ComboboxSelected>>", self._on_project_change)
        self.project_box.bind("<FocusOut>", self._on_project_change)
        more_btn = tk.Menubutton(
            toolbar, text="更多 ▾", bg=CARD, fg="#344054", relief="solid",
            bd=1, font=FONT_SMALL, padx=8, pady=5, cursor="hand2")
        more_menu = tk.Menu(more_btn, tearoff=0, font=FONT)
        more_menu.add_command(label="项目归档管理…", command=self.open_project_archive)
        more_menu.add_command(label="新手引导…", command=self.open_onboarding)
        more_menu.add_command(label="打开数据文件夹", command=self.open_data_folder)
        more_btn.configure(menu=more_menu)
        more_btn.pack(side="left", padx=(0, 12 * s))

        self.capture_btn = tk.Button(
            toolbar, text="登记当前单元格", command=self.do_capture,
            bg="#12B76A", fg="white", activebackground="#0E9F5B",
            activeforeground="white", relief="flat", bd=0, font=FONT_BOLD,
            padx=int(16 * s), pady=int(7 * s), cursor="hand2")
        self.capture_btn.pack(side="left")
        tk.Button(toolbar, text="定位底稿", command=self.locate_selected,
                  bg="#1570EF", fg="white", activebackground="#0B5BC4",
                  activeforeground="white", relief="flat", bd=0, font=FONT_BOLD,
                  padx=int(14 * s), pady=int(6 * s), cursor="hand2").pack(
                      side="left", padx=(12 * s, 8))
        tk.Button(toolbar, text="编辑详情", command=self.edit_selected,
                  bg=CARD, fg="#344054", relief="solid", bd=1, font=FONT,
                  padx=12, pady=5, cursor="hand2").pack(side="left", padx=(0, 8))

        tk.Button(toolbar, text="刷新", command=self.refresh, bg=CARD, fg="#344054",
                  relief="solid", bd=1, font=FONT, padx=12, pady=5,
                  cursor="hand2").pack(side="left", padx=(0, 8))
        export_btn = tk.Menubutton(
            toolbar, text="导出 ▾", bg=CARD, fg="#344054", relief="solid",
            bd=1, font=FONT, padx=12, pady=5, cursor="hand2")
        export_menu = tk.Menu(export_btn, tearoff=0, font=FONT)
        export_menu.add_command(label="导出 CSV", command=self.export_csv)
        export_menu.add_command(label="导出 Excel 复核清单",
                                command=self.export_xlsx)
        export_menu.add_command(label="导出 HTML 复核清单",
                                command=self.export_html)
        export_btn.configure(menu=export_menu)
        export_btn.pack(side="left", padx=(0, 8))
        tk.Label(toolbar, text="数据只保存在本机", bg=CARD, fg=MUTED,
                 font=FONT_SMALL).pack(side="right")

        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x")

        filters = tk.Frame(self.root, bg=BG, padx=int(18 * s), pady=int(8 * s))
        filters.pack(fill="x")
        tk.Label(filters, text="快捷视图", bg=BG, fg="#344054",
                 font=FONT).pack(side="left")
        self.quick_view_box = ttk.Combobox(
            filters, textvariable=self.quick_view_var, state="readonly",
            values=QUICK_VIEW_CHOICES, width=11, font=FONT)
        self.quick_view_box.pack(side="left", padx=(6, 12 * s))
        self.quick_view_box.bind("<<ComboboxSelected>>", self.apply_quick_view)

        tk.Label(filters, text="状态", bg=BG, fg="#344054", font=FONT).pack(side="left")
        self.status_box = ttk.Combobox(
            filters, textvariable=self.status_filter_var, state="readonly",
            values=[STATUS_TODO_FILTER, "全部"] + STATUSES, width=9, font=FONT)
        self.status_box.pack(side="left", padx=(6, 16 * s))
        self.status_box.bind("<<ComboboxSelected>>", self._on_filter_change)

        tk.Label(filters, text="项目", bg=BG, fg="#344054", font=FONT).pack(side="left")
        self.project_filter_box = ttk.Combobox(
            filters, textvariable=self.project_filter_var, state="readonly",
            values=["全部"], width=18, font=FONT)
        self.project_filter_box.pack(side="left", padx=(6, 16 * s))
        self.project_filter_box.bind("<<ComboboxSelected>>", self._on_filter_change)

        tk.Label(filters, text="账龄", bg=BG, fg="#344054", font=FONT).pack(side="left")
        self.aging_box = ttk.Combobox(
            filters, textvariable=self.aging_filter_var, state="readonly",
            values=AGING_CHOICES, width=10, font=FONT)
        self.aging_box.pack(side="left", padx=(6, 10 * s))
        self.aging_box.bind("<<ComboboxSelected>>", self._on_filter_change)
        tk.Checkbutton(
            filters, text="含已归档", variable=self.include_archived_var,
            command=self._on_archive_toggle, bg=BG, fg="#344054",
            activebackground=BG, font=FONT_SMALL).pack(side="left",
                                                       padx=(0, 14 * s))

        tk.Label(filters, text="搜索", bg=BG, fg="#344054", font=FONT).pack(side="left")
        search_entry = ttk.Entry(filters, textvariable=self.search_var, width=24,
                                 font=FONT)
        search_entry.pack(side="left", padx=(6, 6))
        search_entry.bind("<KeyRelease>", self._on_search_key)
        search_entry.bind("<Return>", lambda e: self.refresh())
        tk.Button(filters, text="清除筛选", command=self.clear_filters, bg=BG,
                  fg="#344054", relief="flat", bd=0, font=FONT_SMALL,
                  cursor="hand2").pack(side="left", padx=(0, 6))

        self.paned = ttk.PanedWindow(self.root, orient="vertical")
        self.paned.pack(fill="both", expand=True)
        table_wrap = tk.Frame(self.paned, bg=CARD, padx=int(18 * s),
                              pady=int(8 * s))
        self.paned.add(table_wrap, weight=3)
        # Counts get their own row: at 125%/150% DPI the filter controls use
        # the full width, and a right-packed label in the same row is clipped.
        list_header = tk.Frame(table_wrap, bg=CARD)
        list_header.grid(row=0, column=0, columnspan=2, sticky="ew",
                         pady=(0, int(4 * s)))
        self.counts_label = tk.Label(list_header, text="", bg=CARD, fg=MUTED,
                                     font=FONT_SMALL, anchor="e")
        self.counts_label.pack(side="right")
        columns = ("status", "program", "project", "file", "location",
                   "missing", "waiting", "created")
        self.tree = ttk.Treeview(table_wrap, columns=columns, show="headings",
                                 selectmode="extended")
        headings = {
            "status": "状态", "program": "程序内容", "project": "项目",
            "file": "底稿文件", "location": "位置", "missing": "缺失资料",
            "waiting": "等待", "created": "登记时间",
        }
        widths = {
            "status": 72, "program": 340, "project": 90, "file": 150,
            "location": 120, "missing": 130, "waiting": 55, "created": 100,
        }
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=int(widths[col] * s),
                             minwidth=int(50 * s), anchor="w", stretch=col == "program")
        self.tree.heading(
            "waiting",
            text=("等待 ↓" if self.waiting_sort == "desc" else
                  "等待 ↑" if self.waiting_sort == "asc" else "等待"),
            command=self.toggle_waiting_sort)
        for status, (fg, bg) in STATUS_COLORS.items():
            self.tree.tag_configure(status, foreground=fg, background=bg)
        vsb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="ew")
        table_wrap.rowconfigure(1, weight=1)
        table_wrap.columnconfigure(0, weight=1)

        detail_wrap = tk.Frame(self.paned, bg=CARD, padx=int(18 * s),
                               pady=int(6 * s))
        self.paned.add(detail_wrap, weight=1)
        tk.Label(detail_wrap, text="事项详情", bg=CARD, fg="#344054",
                 font=FONT_BOLD, anchor="w").pack(fill="x")
        detail_body = tk.Frame(detail_wrap, bg=CARD)
        detail_body.pack(fill="both", expand=True)
        self.detail_text = tk.Text(
            detail_body, height=6, wrap="word", font=FONT_SMALL, bg="#F9FAFB",
            fg="#475467", relief="flat", padx=8, pady=6)
        detail_vsb = ttk.Scrollbar(detail_body, orient="vertical",
                                   command=self.detail_text.yview)
        self.detail_text.configure(yscrollcommand=detail_vsb.set,
                                   state="disabled")
        self.detail_text.pack(side="left", fill="both", expand=True)
        detail_vsb.pack(side="right", fill="y")
        detail_actions = tk.Frame(detail_wrap, bg=CARD)
        detail_actions.pack(fill="x", pady=(5, 0))
        self.quick_followup_btn = tk.Button(
            detail_actions, text="快速跟进…", command=self.quick_follow_up_selected,
            bg=CARD, fg="#344054", relief="solid", bd=1, font=FONT_SMALL,
            padx=8, pady=3, cursor="hand2")
        self.quick_followup_btn.pack(side="left")
        self.received_btn = tk.Button(
            detail_actions, text="资料已收到…",
            command=lambda: self.quick_follow_up_selected(received_today=True),
            bg=CARD, fg="#344054", relief="solid", bd=1, font=FONT_SMALL,
            padx=8, pady=3, cursor="hand2")
        self.received_btn.pack(side="left", padx=(6, 0))

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Double-1>", lambda e: self.locate_selected())
        self.tree.bind("<Return>", lambda e: self.locate_selected())
        self.tree.bind("<Button-3>", self._show_context_menu)
        self.tree.bind("<Control-c>", lambda e: self.copy_program())

        self.menu = tk.Menu(self.root, tearoff=0, font=FONT)
        self.menu.add_command(label="定位底稿", command=self.locate_selected)
        self.menu.add_command(label="编辑详情", command=self.edit_selected)
        self.menu.add_command(label="快速跟进…", command=self.quick_follow_up_selected)
        self.menu.add_command(
            label="资料已收到…",
            command=lambda: self.quick_follow_up_selected(received_today=True))
        self.menu.add_command(label="变更历史…", command=self.show_history_selected)
        self.menu.add_command(label="修改项目…", command=self.change_project_selected)
        status_menu = tk.Menu(self.menu, tearoff=0, font=FONT)
        for status in STATUSES:
            status_menu.add_command(
                label=status, command=lambda s=status: self.set_selected_status(s))
        self.menu.add_cascade(label="标记状态", menu=status_menu)
        self.menu.add_separator()
        self.menu.add_command(label="复制程序内容", command=self.copy_program)
        self.menu.add_command(label="复制文件路径", command=self.copy_path)
        self.menu.add_command(label="打开所在文件夹", command=self.open_selected_folder)
        self.menu.add_command(label="项目归档管理…", command=self.open_project_archive)
        self.menu.add_separator()
        self.menu.add_command(label="删除", command=self.delete_selected)

        statusbar = tk.Frame(self.root, bg=CARD, padx=int(18 * s),
                             pady=int(5 * s))
        statusbar.pack(fill="x", side="bottom")
        self.status_label = tk.Label(statusbar, text="就绪", bg=CARD, fg=MUTED,
                                     font=FONT_SMALL, anchor="w")
        self.status_label.pack(side="left")
        tk.Button(statusbar, text="打开数据文件夹", command=self.open_data_folder,
                  bg=CARD, fg=MUTED, relief="flat", bd=0, font=FONT_SMALL,
                  cursor="hand2").pack(side="right")
        self.data_label = tk.Label(
            statusbar, text="数据：%s" % self.data_dir, bg=CARD, fg=MUTED,
            font=FONT_SMALL, anchor="e")
        self.data_label.pack(side="right", padx=(0, 12 * s))

    def _geometry_to_rect(self, geometry: str) -> Optional[Tuple[int, int, int, int]]:
        """Convert a Tk geometry string to an absolute (x, y, w, h) rect.

        Tk treats a leading minus as an offset from the right/bottom edge;
        the primary screen is the best available reference for old configs.
        New configs store explicit absolute coordinates in ``window_rect``.
        """
        match = re.match(
            r"^(\d+)x(\d+)([+-])(\d+)([+-])(-?\d+)$", str(geometry).strip())
        if not match:
            return None
        width, height = int(match.group(1)), int(match.group(2))
        x_sign, x_value = match.group(3), int(match.group(4))
        y_sign, y_value = match.group(5), int(match.group(6))
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        x = x_value if x_sign == "+" else screen_w - width - x_value
        y = y_value if y_sign == "+" else screen_h - height - y_value
        return (x, y, width, height)

    def _restore_window(self):
        work_areas = win32util.list_monitor_work_areas()
        primary = next(
            (area for area in work_areas if area[4]),
            work_areas[0] if work_areas else None)
        if primary is None:
            primary = (0, 0, self.root.winfo_screenwidth(),
                       self.root.winfo_screenheight(), 1)
        primary_w = max(1, primary[2] - primary[0])
        primary_h = max(1, primary[3] - primary[1])
        min_w = min(int(920 * self.scale), max(760, primary_w - 40))
        min_h = min(int(520 * self.scale), max(460, primary_h - 80))
        self.root.minsize(min_w, min_h)

        rect = None
        saved_rect = self.config.get("window_rect")
        if isinstance(saved_rect, (list, tuple)) and len(saved_rect) == 4:
            try:
                rect = tuple(int(value) for value in saved_rect)
            except (TypeError, ValueError):
                rect = None
        if rect is None:
            rect = self._geometry_to_rect(self.config.get("window_geometry", ""))

        if rect is not None:
            work_area = win32util.intersecting_work_area(rect, work_areas) or primary
            work_w = max(1, work_area[2] - work_area[0])
            work_h = max(1, work_area[3] - work_area[1])
            min_w = min(min_w, max(480, work_w - 20))
            min_h = min(min_h, max(360, work_h - 20))
            self.root.minsize(min_w, min_h)
            x, y, w, h = win32util.clamp_rect_to_work_area(
                rect, work_area, min_w, min_h)
            self.root.geometry("%dx%d+%d+%d" % (w, h, x, y))
            log.info("window restored to %dx%d+%d+%d (work area=%s)",
                     w, h, x, y, work_area)
            return

        w = min(int(1240 * self.scale), max(1, primary_w - 40))
        h = min(int(720 * self.scale), max(1, primary_h - 80))
        w, h = max(w, min_w), max(h, min_h)
        x = primary[0] + max(0, (primary_w - w) // 2)
        y = primary[1] + max(0, (primary_h - h) // 2 - int(20 * self.scale))
        self.root.geometry("%dx%d+%d+%d" % (w, h, x, y))

    # ----------------------------------------------------------- data views
    def project_values(self) -> List[str]:
        values = self.store.distinct_projects(include_archived=False)
        current = self.project_var.get().strip()
        if current and current not in values:
            values.append(current)
        return values

    def _aging_days(self) -> Optional[int]:
        return _AGING_DAYS.get(self.aging_filter_var.get())

    def _effective_list_filters(self):
        view = self.quick_view_var.get()
        if view == "待办":
            return {"status_filter": STATUS_TODO_FILTER,
                    "aging_days": None, "unassigned_only": False}
        if view == "待跟进":
            return {"status_filter": STATUS_WAITING,
                    "aging_days": None, "unassigned_only": False}
        if view == "无人负责":
            return {"status_filter": STATUS_TODO_FILTER,
                    "aging_days": None, "unassigned_only": True}
        if view == "等待>7天":
            return {"status_filter": STATUS_WAITING,
                    "aging_days": _AGING_DAYS["等待>7天"],
                    "unassigned_only": False}
        return {"status_filter": self.status_filter_var.get(),
                "aging_days": self._aging_days(), "unassigned_only": False}

    def apply_quick_view(self, _event=None):
        presets = {
            "待办": (STATUS_TODO_FILTER, "全部"),
            "待跟进": (STATUS_WAITING, "全部"),
            "无人负责": (STATUS_TODO_FILTER, "全部"),
            "等待>7天": (STATUS_WAITING, "等待>7天"),
        }
        preset = presets.get(self.quick_view_var.get())
        if preset:
            self.status_filter_var.set(preset[0])
            self.aging_filter_var.set(preset[1])
        self.config.data.update({
            "quick_view": self.quick_view_var.get(),
            "filter_status": self.status_filter_var.get(),
            "filter_aging": self.aging_filter_var.get(),
        })
        self.config.save()
        self.refresh()

    def clear_filters(self):
        self.quick_view_var.set("待办")
        self.status_filter_var.set(STATUS_TODO_FILTER)
        self.project_filter_var.set("全部")
        self.aging_filter_var.set("全部")
        self.include_archived_var.set(False)
        self.search_var.set("")
        self.config.data.update({
            "quick_view": "待办", "filter_status": STATUS_TODO_FILTER,
            "filter_project": "全部", "filter_aging": "全部",
            "include_archived": False, "search_text": "",
        })
        self.config.save()
        self.refresh()

    def toggle_waiting_sort(self):
        self.waiting_sort = "asc" if self.waiting_sort == "desc" else "desc"
        self.config.set("sort_waiting", self.waiting_sort)
        self.tree.heading(
            "waiting", text="等待 ↑" if self.waiting_sort == "asc" else "等待 ↓",
            command=self.toggle_waiting_sort)
        self.refresh()

    def _short(self, text: str, limit: int) -> str:
        text = (text or "").replace("\r", " ").replace("\n", " ").strip()
        return text if len(text) <= limit else text[: limit - 1] + "…"

    def refresh(self, select_id: Optional[int] = None):
        include_archived = bool(self.include_archived_var.get())
        projects = self.store.distinct_projects(include_archived=include_archived)
        self.project_box.configure(values=self.project_values())
        filter_values = ["全部"] + projects
        self.project_filter_box.configure(values=filter_values)
        if self.project_filter_var.get() not in filter_values:
            self.project_filter_var.set("全部")
            self.config.set("filter_project", "全部")

        current = self.tree.selection()
        keep_id = None
        if current:
            try:
                keep_id = int(current[0])
            except Exception:
                keep_id = None

        items = self.store.list_items(
            project=self.project_filter_var.get(),
            keyword=self.search_var.get().strip() or None,
            include_archived=include_archived,
            **self._effective_list_filters())
        if self.waiting_sort in ("asc", "desc"):
            reverse_age = self.waiting_sort == "desc"
            items.sort(key=lambda it: (
                STATUS_PRIORITY.get(it.status, 99),
                it.waiting_days is None,
                -(it.waiting_days or 0) if reverse_age else (it.waiting_days or 0)))
        self.tree.delete(*self.tree.get_children())
        self.items_by_id = {it.id: it for it in items}
        for it in items:
            wait_text = ("%d 天" % it.waiting_days
                         if it.waiting_days is not None else "")
            self.tree.insert(
                "", "end", iid=str(it.id), tags=(it.status,),
                values=(
                    it.status,
                    self._short(it.program_text, 80),
                    it.project or "未分类",
                    it.file_name,
                    it.location_text,
                    self._short(it.missing_info, 26),
                    wait_text,
                    it.created_at[5:16],
                ))
        target = select_id if select_id is not None else keep_id
        if target is not None and str(target) in self.tree.get_children():
            self.tree.selection_set(str(target))
            self.tree.focus(str(target))
            self.tree.see(str(target))
        elif target is None and items and not self.tree.selection():
            first = str(items[0].id)
            self.tree.selection_set(first)
            self.tree.focus(first)
        self._update_detail_panel()

        counts = self.store.counts(include_archived=include_archived)
        self.counts_label.configure(
            text="共 %d 条 · 未完成 %d · 等待资料 %d · 进行中 %d · 已完成 %d" % (
                counts.get("合计", 0), counts.get("未完成", 0),
                counts.get("等待资料", 0), counts.get("进行中", 0),
                counts.get("已完成", 0)))
        if self.tray is not None:
            self.tray.set_tooltip(
                "底稿未完成助手 · 待办 %d · 等待资料 %d" % (
                    counts.get("合计", 0) - counts.get("已完成", 0),
                    counts.get("等待资料", 0)))

    def _ensure_item_visible(self, item_id: int):
        item = self.store.get_item(item_id)
        if item and self.store.project_archived(item.project):
            self.include_archived_var.set(True)
            self.config.set("include_archived", True)
        ids = {it.id for it in self.store.list_items(
            project=self.project_filter_var.get(),
            keyword=self.search_var.get().strip() or None,
            include_archived=self.include_archived_var.get(),
            **self._effective_list_filters())}
        if item_id in ids:
            return
        self.quick_view_var.set("待办")
        self.status_filter_var.set(STATUS_TODO_FILTER)
        self.project_filter_var.set("全部")
        self.search_var.set("")
        self.aging_filter_var.set("全部")
        self.config.data.update({
            "quick_view": "待办", "filter_status": STATUS_TODO_FILTER,
            "filter_project": "全部", "filter_aging": "全部",
            "search_text": "",
        })
        self.config.save()

    def _on_tree_select(self, _event=None):
        self._update_detail_panel()

    def _update_detail_panel(self):
        selection = self.tree.selection()
        items = []
        for iid in selection:
            try:
                item = self.items_by_id.get(int(iid))
            except Exception:
                item = None
            if item is not None:
                items.append(item)
        if not items:
            text = "选择一条事项，在这里查看完整程序、缺失资料和备注。"
        else:
            it = items[0]
            lines = []
            if len(items) > 1:
                lines.append("已选择 %d 条；下方显示第一条。" % len(items))
                lines.append("")
            lines.append("%s · %s · %s" % (
                it.status, it.project or "未分类", it.source_name))
            lines.append("")
            lines.append("程序内容：")
            lines.append(it.program_text or "（空）")
            lines.append("")
            lines.append("位置：%s!%s%s" % (
                it.sheet_name, display_address(it.cell_address),
                ("（合并 %s）" % it.merge_range) if it.merge_range else ""))
            lines.append("文件：%s" % it.file_path)
            if it.status == STATUS_WAITING:
                days = it.waiting_days
                lines.append("等待：%s" % (
                    "%d 天" % days if days is not None else "未记录起始日"))
            if (it.missing_info or it.missing_owner or it.missing_requested_at
                    or it.missing_received_at):
                lines.append("缺失资料：%s" % (it.missing_info or "—"))
                lines.append("负责人：%s    索取：%s    收到：%s" % (
                    it.missing_owner or "—", it.missing_requested_at or "—",
                    it.missing_received_at or "—"))
            if it.note:
                lines.append("备注：%s" % it.note)
            lines.append("登记：%s    更新：%s    完成：%s" % (
                it.created_at, it.updated_at, it.completed_at or "—"))
            text = "\n".join(lines)
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.insert("1.0", text)
        self.detail_text.configure(state="disabled")
        self.detail_text.yview_moveto(0)
        state = tk.NORMAL if len(items) == 1 else tk.DISABLED
        self.quick_followup_btn.configure(state=state)
        self.received_btn.configure(state=state)

    def _on_archive_toggle(self, _event=None):
        self.config.set("include_archived", self.include_archived_var.get())
        self.refresh()

    # ------------------------------------------------------------- handlers
    def _on_project_change(self, _event=None):
        self.config.set("current_project", self.project_var.get().strip())
        self.project_box.configure(values=self.project_values())

    def _on_filter_change(self, _event=None):
        if (_event is not None and
                getattr(_event, "widget", None) in
                (getattr(self, "status_box", None), getattr(self, "aging_box", None))):
            self.quick_view_var.set("自定义")
        self.config.data.update({
            "quick_view": self.quick_view_var.get(),
            "filter_status": self.status_filter_var.get(),
            "filter_project": self.project_filter_var.get(),
            "filter_aging": self.aging_filter_var.get(),
        })
        self.config.save()
        self.refresh()

    def _on_search_key(self, event=None):
        if event is not None and event.keysym in ("Return",):
            return
        if self._search_after is not None:
            try:
                self.root.after_cancel(self._search_after)
            except Exception:
                pass
        self._search_after = self.root.after(300, self._do_search)

    def _do_search(self):
        self._search_after = None
        self.config.set("search_text", self.search_var.get().strip())
        self.refresh()

    def clear_search(self):
        self.search_var.set("")
        self.config.set("search_text", "")
        self.refresh()

    def do_capture(self, from_hotkey: bool = False,
                   foreground_hwnd: int = 0):
        if self.busy:
            self.toast("请稍候", "上一步操作尚未完成", "warning")
            return
        if self._dialog_open():
            self.toast("请先处理当前窗口", "有对话框尚未关闭", "warning", 2200)
            return
        project = self.project_var.get().strip()

        target_hwnd = foreground_hwnd if from_hotkey else self.last_sheet_hwnd
        if not from_hotkey and not target_hwnd:
            self._ensure_visible()
            messagebox.showwarning(
                "未检测到表格窗口",
                "请先切换到 Excel 或 WPS 表格，选中一个单元格，"
                "再返回本窗口点击“登记当前单元格”。",
                parent=self.root)
            return

        def ok(info: SelectionInfo):
            self._handle_capture(info, project)

        def err(exc):
            self._handle_capture_error(exc)

        self._submit("正在读取当前单元格",
                     lambda: capture.read_selection(target_hwnd,
                                                    require_foreground=from_hotkey),
                     ok, err)

    def _handle_capture(self, info: SelectionInfo, project: str):
        project = suggest_project(
            info.file_path, project,
            bool(self.config.get("auto_project_from_folder", True)))
        if not self.project_var.get().strip() or self.project_var.get().strip() == "未分类":
            self.project_var.set(project)
            self.config.set("current_project", project)
        dup = self.store.find_active_duplicate(
            info.file_key, info.sheet_name, info.cell_address,
            info.file_volume, info.file_index)
        if dup:
            self._ensure_visible()
            choice = DuplicateDialog(self, dup).show()
            if choice == "existing":
                self.refresh(select_id=dup.id)
                self.show_detail(dup.id)
                return
            if choice == "update":
                if nativeom.path_key(dup.file_path) != info.file_key:
                    self.store.update_path(dup.id, info.file_path,
                                           info.file_volume, info.file_index)
                self.store.update_snapshot(dup.id, info.snapshot)
                self.refresh(select_id=dup.id)
                self.toast("已更新现有事项", "内容快照已更新，可直接补充资料。",
                           "success")
                self.show_detail(dup.id)
                return
            if choice != "add":
                return
        item_id = self.store.add_item(info, project)
        self._ensure_item_visible(item_id)
        self.refresh(select_id=item_id)
        self.set_status("已登记：%s" % info.file_name)
        self.toast("已登记", "%s · %s!%s · 项目：%s" % (
            info.file_name, info.sheet_name, display_address(info.cell_address),
            project), "success", 6000, action_text="撤销",
            action=lambda: self.undo_capture(item_id),
            secondary_action_text="查看事项",
            secondary_action=lambda: self.show_detail(item_id))

    def undo_capture(self, item_id: int):
        item = self.store.get_item(item_id)
        if item is None:
            self.refresh()
            return
        if item.status != STATUS_TODO or item.updated_at != item.created_at:
            self.toast("未撤销", "该事项已被修改，请到详情中确认后删除。", "warning")
            return
        self.store.delete_item(item_id)
        self.refresh()
        self.toast("已撤销", "刚登记的事项已删除。", "info")

    def _handle_capture_error(self, exc):
        self.set_status("登记失败")
        if isinstance(exc, capture.CaptureError):
            self._ensure_visible()
            messagebox.showwarning("无法登记", exc.user_message, parent=self.root)
        else:
            log.error("capture error: %r", exc)
            self._ensure_visible()
            messagebox.showerror(
                "登记失败",
                "读取当前单元格时发生错误。\n请确认 Excel/WPS 未处于弹窗或编辑状态，"
                "然后重试。\n\n详情已写入日志文件。",
                parent=self.root)

    def locate_selected(self):
        items = self._selected_items()
        if not items:
            messagebox.showinfo("未选择事项", "请先在清单中选择一条事项。",
                                parent=self.root)
            return
        if len(items) > 1:
            messagebox.showinfo("请选择一条", "定位底稿一次只能选择一条事项。",
                                parent=self.root)
            return
        self.locate_item(items[0])

    def locate_item(self, item: Item, allow_mismatch: bool = False,
                    after_success: Optional[Callable] = None,
                    snapshot_to_save: Optional[str] = None,
                    relink_from: Optional[str] = None):
        if self.busy:
            self.toast("请稍候", "上一步操作尚未完成", "warning")
            return
        if self._dialog_open():
            self.toast("请先处理当前窗口", "有对话框尚未关闭", "warning", 2200)
            return

        def ok(outcome):
            if snapshot_to_save is not None:
                self.store.update_snapshot(item.id, snapshot_to_save)
                self.refresh(select_id=item.id)
            if after_success:
                try:
                    after_success(outcome)
                except Exception:
                    log.exception("after_success failed")
                    self._ensure_visible()
                    messagebox.showerror(
                        "未能保存修改",
                        "已定位底稿，但未能保存相关路径或事项信息；请检查数据文件后重试。",
                        parent=self.root)
                    return
            self.set_status("已定位：%s" % item.file_name)
            self.toast("已定位", "%s · %s!%s" % (
                os.path.basename(outcome.file_path), outcome.sheet_name,
                display_address(outcome.cell_address)), "success")

        def err(exc):
            self._handle_jump_error(
                item, exc, allow_mismatch, after_success, relink_from=relink_from)

        self._submit("正在定位底稿",
                     lambda: jump.jump_to_item(item, allow_mismatch), ok, err)

    def _handle_jump_error(self, item: Item, exc, allow_mismatch: bool,
                           after_success, relink_from: Optional[str] = None):
        if isinstance(exc, jump.ContentMismatch):
            self._ensure_visible()
            dialog_kwargs = {"relink_from": relink_from} if relink_from else {}
            choice = MismatchDialog(
                self, item, exc.old, exc.new, **dialog_kwargs).show()
            if choice == "jump":
                self.locate_item(item, allow_mismatch=True,
                                 after_success=after_success,
                                 relink_from=relink_from)
            elif choice == "update":
                # Persist only after the same newly confirmed content has been
                # rechecked and the target was successfully selected.
                candidate = copy.copy(item)
                candidate.snapshot = exc.new
                self.locate_item(candidate, allow_mismatch=False,
                                 after_success=after_success,
                                 snapshot_to_save=exc.new,
                                 relink_from=relink_from)
            return

        if isinstance(exc, jump.FileMissing):
            self._ensure_visible()
            path = RelinkDialog(self, item).show()
            if not path:
                return
            new_volume, new_index = win32util.file_identity(path)
            new_item = copy.copy(item)
            new_item.file_path = path
            new_item.file_key = nativeom.path_key(path)
            new_item.file_name = os.path.basename(path)
            new_item.file_volume = new_volume
            new_item.file_index = new_index

            def relocated(_outcome):
                if new_volume and new_index:
                    self.store.update_path(item.id, path, new_volume, new_index)
                else:
                    self.store.update_path(item.id, path)
                self.store.log_event(item.id, "relocate", item.file_path, path)
                self.refresh(select_id=item.id)
                self.toast("已更新路径", os.path.basename(path), "success")

            self.locate_item(new_item, allow_mismatch=False,
                             after_success=relocated, relink_from=item.file_path)
            return

        if isinstance(exc, jump.SheetMissing):
            self._ensure_visible()
            messagebox.showerror("找不到工作表", exc.user_message, parent=self.root)
            return

        if isinstance(exc, jump.OpenTimeout):
            self._ensure_visible()
            messagebox.showwarning("未能自动打开", exc.user_message, parent=self.root)
            return

        if isinstance(exc, jump.JumpError):
            self._ensure_visible()
            messagebox.showerror("定位失败", exc.user_message, parent=self.root)
            return

        log.error("jump error: %r", exc)
        self._ensure_visible()
        messagebox.showerror(
            "定位失败", "定位底稿时发生错误，详情已写入日志文件。", parent=self.root)

    def edit_selected(self):
        items = self._selected_items()
        if not items:
            return
        if len(items) > 1:
            messagebox.showinfo("请选择一条", "编辑详情一次只能选择一条事项。",
                                parent=self.root)
            return
        self.show_detail(items[0].id)

    def show_detail(self, item_id: int):
        item = self.store.get_item(item_id)
        if not item:
            self.refresh()
            return
        result = DetailDialog(self, item).show()
        if result == "locate":
            fresh = self.store.get_item(item_id)
            if fresh:
                self.locate_item(fresh)
        elif result in ("saved", "deleted"):
            self.refresh(select_id=item_id if result == "saved" else None)

    def update_item(self, item_id: int, values: dict):
        self.store.update_item(item_id, **values)
        self.set_status("已保存")

    def delete_item(self, item_id: int, refresh: bool = True):
        self.store.delete_item(item_id)
        if refresh:
            self.refresh()
        self.set_status("已删除")
        self.toast("已删除", "事项已从清单移除", "info")

    def delete_selected(self):
        items = self._selected_items()
        if not items:
            return
        if len(items) == 1:
            message = "删除后不可恢复，确定删除这条事项吗？\n\n%s" % self._short(
                items[0].program_text, 80)
        else:
            message = "删除后不可恢复，确定删除选中的 %d 条事项吗？" % len(items)
        if messagebox.askyesno("删除事项", message, parent=self.root):
            for item in items:
                self.store.delete_item(item.id)
            self.refresh()
            self.set_status("已删除 %d 条" % len(items))
            self.toast("已删除", "已从清单移除 %d 条事项。" % len(items), "info")

    def quick_follow_up_selected(self, received_today: bool = False):
        items = self._selected_items()
        if not items:
            messagebox.showinfo("未选择事项", "请先选择一条事项。",
                                parent=self.root)
            return
        if len(items) > 1:
            messagebox.showinfo("请选择一条", "快速跟进一次只能编辑一条事项。",
                                parent=self.root)
            return
        item = items[0]
        result = QuickFollowUpDialog(
            self, item, received_today=received_today).show()
        if result == "saved":
            self.refresh(select_id=item.id)
            self.toast("跟进信息已保存", item.file_name, "success")

    def set_selected_status(self, status: str):
        items = self._selected_items()
        if not items:
            return
        for item in items:
            self.store.set_status(item.id, status)
        self.refresh(select_id=items[0].id)
        self.set_status("已标记为%s（%d 条）" % (status, len(items)))
        if len(items) == 1:
            message = "%s → %s" % (self._short(items[0].program_text, 30), status)
        else:
            message = "%d 条事项 → %s" % (len(items), status)
        self.toast("状态已更新", message, "success")

    def change_project_selected(self):
        items = self._selected_items()
        if not items:
            return
        initial = items[0].project if len(items) == 1 else "未分类"
        value = simpledialog.askstring(
            "修改项目", "输入项目名称（留空表示未分类）：",
            initialvalue=initial, parent=self.root)
        if value is None:
            return
        project = value.strip() or "未分类"
        for item in items:
            self.store.update_item(item.id, project=project)
        self.refresh(select_id=items[0].id)
        self.toast("项目已更新", "%d 条事项 → %s" % (len(items), project),
                   "success")

    def show_history_selected(self):
        items = self._selected_items()
        if not items:
            return
        if len(items) > 1:
            messagebox.showinfo("请选择一条", "变更历史一次只能查看一条事项。",
                                parent=self.root)
            return
        HistoryDialog(self, items[0]).show()

    def copy_program(self):
        items = self._selected_items()
        if not items:
            return
        text = "\n\n---\n\n".join(it.program_text for it in items)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.toast("已复制", "程序内容已复制到剪贴板", "info", 2200)

    def copy_path(self):
        items = self._selected_items()
        if not items:
            return
        text = "\n".join(it.file_path for it in items)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.toast("已复制", "文件路径已复制到剪贴板", "info", 2200)

    def open_selected_folder(self):
        items = self._selected_items()
        if not items:
            return
        folder = os.path.dirname(items[0].file_path)
        if folder and os.path.isdir(folder):
            subprocess.Popen(["explorer", folder])
        else:
            messagebox.showinfo("文件夹不存在", "该文件夹当前不存在。",
                                parent=self.root)

    def export_csv(self):
        default_name = "底稿未完成事项_%s.csv" % _dt.datetime.now().strftime("%Y%m%d")
        initial_dir = self.config.get("last_export_dir", "") or self.data_dir
        path = filedialog.asksaveasfilename(
            parent=self.root, title="导出 CSV", initialdir=initial_dir,
            initialfile=default_name, defaultextension=".csv",
            filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            count = self.store.export_csv(path)
            self.config.set("last_export_dir", os.path.dirname(path))
            messagebox.showinfo("导出完成", "已导出 %d 条事项到：\n%s" % (count, path),
                                parent=self.root)
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc), parent=self.root)

    def export_xlsx(self):
        default_name = "底稿未完成事项复核清单_%s.xlsx" % _dt.datetime.now().strftime("%Y%m%d")
        initial_dir = self.config.get("last_export_dir", "") or self.data_dir
        path = filedialog.asksaveasfilename(
            parent=self.root, title="导出 Excel 复核清单",
            initialdir=initial_dir, initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            count = self.store.export_xlsx(path)
            self.config.set("last_export_dir", os.path.dirname(path))
            messagebox.showinfo(
                "导出完成",
                "已导出 %d 条事项到：\n%s\n\n"
                "“打开文件”列含底稿链接；精确位置请使用“定位底稿”按钮。"
                % (count, path), parent=self.root)
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc), parent=self.root)

    def export_html(self):
        default_name = "底稿未完成事项复核清单_%s.html" % _dt.datetime.now().strftime("%Y%m%d")
        initial_dir = self.config.get("last_export_dir", "") or self.data_dir
        path = filedialog.asksaveasfilename(
            parent=self.root, title="导出复核清单", initialdir=initial_dir,
            initialfile=default_name, defaultextension=".html",
            filetypes=[("HTML 文件", "*.html"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            count = self.store.export_html(path)
            self.config.set("last_export_dir", os.path.dirname(path))
            messagebox.showinfo(
                "导出完成",
                "已导出 %d 条事项到：\n%s\n\n"
                "链接只打开底稿文件；精确位置请使用“定位底稿”按钮。" % (count, path),
                parent=self.root)
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc), parent=self.root)

    def open_data_folder(self):
        try:
            os.startfile(self.data_dir)  # noqa: S606 - local folder
        except Exception as exc:
            messagebox.showerror("无法打开文件夹", str(exc), parent=self.root)

    def backup_data(self, dest: str):
        self.store.backup(dest)

    def _start_tray(self):
        icon_path = _resource_path(os.path.join("assets", "app.ico"))
        self.tray = TrayIcon(
            icon_path, "底稿未完成助手",
            lambda: self.ui_queue.put(("tray", "show")),
            lambda: self.ui_queue.put(("tray", "capture")),
            lambda: self.ui_queue.put(("tray", "data")),
            lambda: self.ui_queue.put(("tray", "exit")))
        if not self.tray.start():
            self.tray = None
            self.close_to_tray = False
            log.warning("tray unavailable; close button will exit")
        else:
            log.info("tray icon started")

    def set_close_to_tray(self, enabled: bool):
        self.close_to_tray = bool(enabled and self.tray is not None)
        self.config.set("close_to_tray", self.close_to_tray)
        if enabled and self.tray is None:
            messagebox.showwarning(
                "托盘不可用", "当前系统未能创建托盘图标，关闭窗口将直接退出程序。",
                parent=self.root)

    def open_project_archive(self):
        ProjectArchiveDialog(self).show()

    def show_project_items(self, project: str):
        self.include_archived_var.set(True)
        self.config.set("include_archived", True)
        self.project_filter_var.set(project)
        self.status_filter_var.set("全部")
        self.aging_filter_var.set("全部")
        self.search_var.set("")
        self.config.set("filter_project", project)
        self.config.set("filter_status", "全部")
        self.config.set("filter_aging", "全部")
        self.config.set("search_text", "")
        self.refresh()

    def _auto_backup(self):
        if not bool(self.config.get("auto_backup_enabled", True)):
            return
        today = _dt.date.today().isoformat()
        if self.config.get("last_auto_backup", "") == today:
            return
        try:
            backup_dir = os.path.join(self.data_dir, "backups")
            os.makedirs(backup_dir, exist_ok=True)
            dest = os.path.join(
                backup_dir, "auto-%s.db" % _dt.datetime.now().strftime("%Y%m%d-%H%M%S"))
            self.store.backup(dest)
            self.config.set("last_auto_backup", today)
            keep = max(1, int(self.config.get("auto_backup_keep", 7)))
            files = sorted(glob.glob(os.path.join(backup_dir, "auto-*.db")))
            for old in files[:-keep]:
                try:
                    os.remove(old)
                except OSError:
                    pass
            log.info("auto backup completed")
        except Exception:
            log.exception("auto backup failed")

    def encrypted_backup(self, parent=None):
        parent = parent or self.root
        default_name = "底稿未完成事项加密备份_%s.db.enc" % _dt.datetime.now().strftime(
            "%Y%m%d_%H%M%S")
        dest = filedialog.asksaveasfilename(
            parent=parent, title="加密备份",
            initialdir=os.path.join(self.data_dir, "backups"),
            initialfile=default_name, defaultextension=".enc",
            filetypes=[("加密备份", "*.enc"), ("所有文件", "*.*")])
        if not dest:
            return
        handle, temp_path = tempfile.mkstemp(prefix="wpt_backup_", suffix=".db")
        os.close(handle)
        try:
            self.store.backup(temp_path)
            secure.encrypt_file(temp_path, dest)
            messagebox.showinfo(
                "加密备份完成",
                "已加密备份到：\n%s\n\n"
                "该文件只能由当前 Windows 用户在同一台电脑上解密。"
                "请同时保留普通备份。" % dest,
                parent=parent)
        except Exception as exc:
            messagebox.showerror("加密备份失败", str(exc), parent=parent)
        finally:
            try:
                os.remove(temp_path)
            except OSError:
                pass

    def decrypt_backup(self, parent=None):
        parent = parent or self.root
        src = filedialog.askopenfilename(
            parent=parent, title="解密备份",
            initialdir=os.path.join(self.data_dir, "backups"),
            filetypes=[("加密备份", "*.enc"), ("所有文件", "*.*")])
        if not src:
            return
        dest = filedialog.asksaveasfilename(
            parent=parent, title="保存解密后的数据库",
            initialdir=os.path.join(self.data_dir, "backups"),
            initialfile="解密后的备份.db", defaultextension=".db",
            filetypes=[("SQLite 数据库", "*.db"), ("所有文件", "*.*")])
        if not dest:
            return
        try:
            secure.decrypt_file(src, dest)
            conn = sqlite3.connect(dest)
            try:
                result = conn.execute("PRAGMA integrity_check").fetchone()
                if not result or result[0] != "ok":
                    raise ValueError("解密后的数据库完整性检查未通过")
            finally:
                conn.close()
            messagebox.showinfo(
                "解密完成",
                "已解密到：\n%s\n\n"
                "如需恢复，请退出本程序，用该文件替换数据目录中的 data.db。" % dest,
                parent=parent)
        except Exception as exc:
            try:
                if os.path.exists(dest):
                    os.remove(dest)
            except OSError:
                pass
            messagebox.showerror("解密失败", str(exc), parent=parent)

    def change_data_dir(self, parent=None):
        parent = parent or self.root
        target = filedialog.askdirectory(
            parent=parent, title="选择新的数据目录",
            initialdir=os.path.dirname(self.data_dir) or self.data_dir)
        if not target:
            return
        target = os.path.abspath(target)
        current = os.path.abspath(self.data_dir)
        if os.path.normcase(target) == os.path.normcase(current):
            messagebox.showinfo("数据目录", "所选目录就是当前数据目录。", parent=parent)
            return
        if os.path.normcase(target).startswith(os.path.normcase(current) + os.sep):
            messagebox.showwarning(
                "目录无效", "不能选择当前数据目录的子目录。", parent=parent)
            return
        target_db = os.path.join(target, "data.db")
        if os.path.exists(target_db):
            if not messagebox.askyesno(
                    "目标目录已有数据",
                    "目标目录已有 data.db。\n\n"
                    "是否切换到该目录的现有数据？\n"
                    "本操作不会覆盖该目录中的文件。",
                    parent=parent):
                return
            config.set_data_dir(target)
            messagebox.showinfo(
                "已切换数据目录",
                "下次启动将使用：\n%s\n\n请退出并重新启动程序。" % target,
                parent=parent)
            return
        if not messagebox.askyesno(
                "迁移数据目录",
                "将把当前 data.db、config.json 和 backups 复制到：\n%s\n\n"
                "原目录会保留作为备份。继续吗？" % target,
                parent=parent):
            return
        try:
            os.makedirs(target, exist_ok=True)
            self.store.backup(os.path.join(target, "data.db"))
            src_cfg = os.path.join(self.data_dir, "config.json")
            if os.path.exists(src_cfg):
                shutil.copy2(src_cfg, os.path.join(target, "config.json"))
            src_backups = os.path.join(self.data_dir, "backups")
            if os.path.isdir(src_backups):
                shutil.copytree(src_backups, os.path.join(target, "backups"),
                                dirs_exist_ok=True)
            config.set_data_dir(target)
            messagebox.showinfo(
                "迁移完成",
                "数据已复制到：\n%s\n\n请退出并重新启动程序。\n"
                "确认新目录正常后，旧目录可按资料管理要求清理。" % target,
                parent=parent)
        except Exception as exc:
            messagebox.showerror("迁移失败", str(exc), parent=parent)

    def open_log_folder(self):
        try:
            os.startfile(self.data_dir)  # noqa: S606 - local folder
        except Exception as exc:
            messagebox.showerror("无法打开文件夹", str(exc), parent=self.root)

    def clear_logs(self, parent=None):
        parent = parent or self.root
        if not messagebox.askyesno(
                "清空日志", "确定清空本机运行日志吗？事项数据不会受影响。",
                parent=parent):
            return
        root_logger = logging.getLogger()
        for handler in list(root_logger.handlers):
            if not isinstance(handler, logging.handlers.RotatingFileHandler):
                continue
            try:
                handler.close()
                root_logger.removeHandler(handler)
            except Exception:
                pass
        for path in glob.glob(os.path.join(self.data_dir, "app.log*")):
            try:
                os.remove(path)
            except OSError:
                pass
        _setup_logging(self.data_dir)
        log.info("logs cleared by user")
        messagebox.showinfo("日志已清空", "运行日志已清空。", parent=parent)

    def hotkey_status_text(self) -> str:
        thread = self.hotkey_thread
        if thread is not None and thread.is_alive() and thread.error is None:
            return "全局热键 %s 已成功注册。" % self.hotkey_text
        return "全局热键当前不可用；可在设置中更换组合键，或用窗口按钮登记。"

    def open_practice_workbook(self) -> str:
        path = practice.create_practice_workbook(self.data_dir)
        try:
            os.startfile(path)
        except (AttributeError, OSError) as exc:
            messagebox.showinfo(
                "练习底稿已生成",
                "文件已保存在本机：\n%s\n请用 Excel 或 WPS 手动打开。\n\n%s" %
                (path, exc), parent=self.root)
        return path

    def dismiss_onboarding(self):
        self.config.set("onboarding_completed", True)
        if self.onboarding_banner is not None:
            try:
                self.onboarding_banner.pack_forget()
            except Exception:
                pass
            self.onboarding_banner = None

    def open_onboarding(self):
        result = FirstRunDialog(self).show()
        if result == "settings":
            self.open_settings()

    def open_settings(self):
        SettingsDialog(self).show()

    def _show_context_menu(self, event):
        row = self.tree.identify_row(event.y)
        if row and row not in self.tree.selection():
            self.tree.selection_set(row)
            self.tree.focus(row)
        self.menu.tk_popup(event.x_root, event.y_root)

    def _selected_items(self) -> List[Item]:
        result = []
        for iid in self.tree.selection():
            try:
                item = self.items_by_id.get(int(iid))
            except Exception:
                item = None
            if item is not None:
                result.append(item)
        return result

    def _selected_item(self) -> Optional[Item]:
        items = self._selected_items()
        return items[0] if items else None

    # --------------------------------------------------------------- hotkey
    def _start_hotkey(self, initial: bool = False) -> Optional[str]:
        configured = self.hotkey_text
        error = self._start_hotkey_thread(configured)
        if error:
            self.hotkey_label.configure(text="登记热键不可用")
            self.capture_btn.configure(text="登记当前单元格")
            if initial and self.config.get("onboarding_completed", False):
                self.root.after(400, lambda: self._warn_hotkey(error))
            return error
        self.hotkey_label.configure(text="登记热键 %s" % self.hotkey_text)
        self.capture_btn.configure(text="登记当前单元格  %s" % self.hotkey_text)
        return None

    def _start_hotkey_thread(self, text: str) -> Optional[str]:
        if self.hotkey_thread is not None:
            try:
                self.hotkey_thread.stop()
                self.hotkey_thread.join(timeout=1.0)
            except Exception:
                pass
            self.hotkey_thread = None
        try:
            thread = HotkeyThread(
                text,
                lambda: self.ui_queue.put(
                    ("hotkey", win32util.get_foreground_window())))
            thread.start()
            if not thread.wait_ready(3.0):
                thread.stop()
                return "热键注册超时，请重试或更换组合键。"
        except Exception as exc:
            return "无法注册热键 %s：%s" % (text, exc)
        if thread.error is not None:
            return thread.error.args[0] if thread.error.args else str(thread.error)
        self.hotkey_thread = thread
        self.hotkey_text = thread.spec.text
        return None

    def _warn_hotkey(self, message: str):
        self._ensure_visible()
        messagebox.showwarning(
            "登记热键不可用",
            message + "\n\n您仍可以点击“登记当前单元格”按钮，"
                      "或在“设置”中更换热键。",
            parent=self.root)

    def apply_hotkey(self, text: str):
        try:
            spec = parse_hotkey(text)
        except ValueError as exc:
            return False, str(exc)
        old_text = self.hotkey_text
        error = self._start_hotkey_thread(spec.text)
        if error:
            # Try to restore the previous hotkey.
            self._start_hotkey_thread(old_text)
            return False, error
        self.config.set("hotkey", spec.text)
        self.hotkey_text = spec.text
        self.hotkey_label.configure(text="登记热键 %s" % spec.text)
        self.capture_btn.configure(text="登记当前单元格  %s" % spec.text)
        return True, ""

    # ------------------------------------------------------------ task queue
    def _submit(self, label: str, func, on_success, on_error):
        if self.busy:
            self.toast("请稍候", "上一步操作尚未完成", "warning")
            return
        self.busy = True
        self.set_status(label + "…")
        self.worker.submit(func, label, on_success, on_error)

    def _track_foreground(self):
        if self._closing:
            return
        try:
            fg = win32util.get_foreground_window()
            xlmain = win32util.resolve_sheet_window(fg)
            if xlmain and win32util.user32.IsWindowVisible(xlmain):
                self.last_sheet_hwnd = int(xlmain)
        except Exception:
            pass
        try:
            self.root.after(400, self._track_foreground)
        except Exception:
            pass

    def _poll_queue(self):
        if self._closing:
            return
        while True:
            try:
                msg = self.ui_queue.get_nowait()
            except queue.Empty:
                break
            if msg[0] == "hotkey":
                self.do_capture(from_hotkey=True,
                                foreground_hwnd=msg[1] if len(msg) > 1 else 0)
            elif msg[0] == "tray":
                action = msg[1] if len(msg) > 1 else ""
                if action == "show":
                    self._show_main_window()
                elif action == "capture":
                    self.do_capture(from_hotkey=False,
                                    foreground_hwnd=self.last_sheet_hwnd)
                elif action == "data":
                    self.open_data_folder()
                elif action == "exit":
                    self.on_close(real_exit=True)
                    return
            elif msg[0] == "task":
                _, _label, result, exc, on_success, on_error = msg
                self.busy = False
                self.set_status("就绪")
                try:
                    if exc is not None:
                        on_error(exc)
                    else:
                        on_success(result)
                except Exception:
                    log.exception("task callback failed")
                    self._ensure_visible()
                    messagebox.showerror(
                        "操作失败", "处理结果时发生错误，详情已写入日志。",
                        parent=self.root)
        if not self._closing:
            try:
                self.root.after(100, self._poll_queue)
            except Exception:
                pass

    def set_status(self, text: str):
        self.status_label.configure(text=text)

    def toast(self, title: str, message: str = "", kind: str = "info",
              duration: int = 3800, action_text: str = "", action=None,
              secondary_action_text: str = "", secondary_action=None):
        self.toast_obj.show(
            title, message, kind, duration, action_text=action_text,
            action=action, secondary_action_text=secondary_action_text,
            secondary_action=secondary_action)

    # -------------------------------------------------------------- helpers
    def _show_main_window(self):
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        except Exception:
            pass
        try:
            frame = self.root.wm_frame()
            hwnd = int(frame, 16) if frame else 0
            if hwnd:
                bring_to_front(hwnd)
        except Exception:
            pass

    def _ensure_visible(self):
        try:
            if self.root.state() in ("iconic", "withdrawn"):
                self.root.deiconify()
            self.root.lift()
        except Exception:
            pass

    def _dialog_open(self) -> bool:
        try:
            return self.root.grab_current() is not None
        except Exception:
            return False

    def _report_exception(self, exc_type, exc_value, tb):
        log.exception("unhandled Tk exception", exc_info=(exc_type, exc_value, tb))
        try:
            self._ensure_visible()
            messagebox.showerror(
                "程序错误", "发生未处理的错误，详情已写入日志。\n%s" % exc_value,
                parent=self.root)
        except Exception:
            pass

    def on_close(self, real_exit: bool = False):
        try:
            if self.root.state() not in ("iconic", "withdrawn"):
                self.config.set("window_geometry", self.root.geometry())
                self.config.set("window_rect", [
                    self.root.winfo_x(), self.root.winfo_y(),
                    self.root.winfo_width(), self.root.winfo_height()])
        except Exception:
            pass
        if not real_exit and self.close_to_tray and self.tray is not None:
            self.root.withdraw()
            log.info("window hidden to tray")
            if not self._tray_hint_shown:
                self._tray_hint_shown = True
                self.config.set("tray_hint_shown", True)
                self.toast(
                    "仍在后台运行",
                    "已最小化到托盘；右键托盘图标可显示窗口、登记最近单元格或退出。",
                    "info", 4500)
            return
        self._closing = True
        if self.tray is not None:
            try:
                self.tray.stop()
            except Exception:
                log.exception("tray shutdown failed")
            self.tray = None
        if self.hotkey_thread is not None:
            try:
                self.hotkey_thread.stop()
                self.hotkey_thread.join(timeout=1.0)
            except Exception:
                pass
        try:
            self.worker.stop()
            self.toast_obj.close()
        except Exception:
            log.exception("worker shutdown failed")
        try:
            self.store.close()
        except Exception:
            pass
        log.info("app closed")
        self.root.destroy()


def _setup_logging(data_dir: str):
    os.makedirs(data_dir, exist_ok=True)
    path = os.path.abspath(os.path.join(data_dir, "app.log"))
    handler = logging.handlers.RotatingFileHandler(
        path, maxBytes=1024 * 1024, backupCount=3, encoding="utf-8")
    handler.setFormatter(_RedactingFormatter(
        "%(asctime)s %(levelname)s %(name)s: %(message)s"))
    root_logger = logging.getLogger()
    for existing in list(root_logger.handlers):
        if (isinstance(existing, logging.handlers.RotatingFileHandler) and
                os.path.abspath(getattr(existing, "baseFilename", "")) == path):
            try:
                existing.close()
                root_logger.removeHandler(existing)
            except Exception:
                pass
    root_logger.setLevel(logging.INFO)
    root_logger.addHandler(handler)


def _set_dpi_awareness():
    try:
        # Per-monitor v2 (Windows 10 1703+); fall back to system aware.
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(
                ctypes.c_void_p(-4)):
            return
    except Exception:
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def _resource_path(relative: str) -> str:
    base = getattr(
        sys, "_MEIPASS",
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base, relative)


def main():
    _set_dpi_awareness()
    data_dir = config.default_data_dir()
    os.makedirs(data_dir, exist_ok=True)
    _setup_logging(data_dir)

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    mutex = kernel32.CreateMutexW(None, False, "Local\\WorkpaperTodo_SingleInstance")
    if ctypes.get_last_error() == 183:  # ERROR_ALREADY_EXISTS
        from .win32util import user32
        hwnd = user32.FindWindowW(None, __window_title__)
        if hwnd:
            bring_to_front(int(hwnd))
        else:
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo(__window_title__, "底稿未完成助手已经在运行。")
            root.destroy()
        return

    cfg = config.Config(data_dir)
    db_path = os.path.join(data_dir, "data.db")
    db = store.Store(db_path)

    root = tk.Tk()
    try:
        dpi = root.winfo_fpixels("1i")
        root.tk.call("tk", "scaling", dpi / 72.0)
    except Exception:
        pass

    icon_path = _resource_path(os.path.join("assets", "app.ico"))
    if os.path.exists(icon_path):
        try:
            root.iconbitmap(icon_path)
        except Exception:
            pass

    app = App(root, db, cfg, data_dir)
    root.mainloop()
    # keep mutex alive until process end
    _ = mutex


if __name__ == "__main__":
    main()
