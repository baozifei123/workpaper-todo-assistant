# -*- coding: utf-8 -*-
"""对话框：重复登记、内容变化、事项详情、设置。"""
from __future__ import annotations

import datetime as _dt
import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Optional

from . import __version__, win32util
from .models import (
    Item, STATUSES, STATUS_DONE, STATUS_TODO, STATUS_WAITING, display_address,
    parse_date, waiting_days,
)
from .win32util import HOTKEY_PRESETS

FONT = ("Microsoft YaHei UI", 10)
FONT_BOLD = ("Microsoft YaHei UI", 10, "bold")
FONT_SMALL = ("Microsoft YaHei UI", 9)
FONT_TITLE = ("Microsoft YaHei UI", 13, "bold")
BG = "#F9FAFB"
CARD = "#FFFFFF"


class BaseDialog(tk.Toplevel):
    def __init__(self, app, title: str, width: int, height: int):
        super().__init__(app.root)
        self.app = app
        self.result = None
        self.title(title)
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(480, 360)
        self.transient(app.root)
        self._center(width, height)
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.bind("<Escape>", lambda e: self._cancel())

    def _center(self, width: int, height: int):
        self.update_idletasks()
        work_area = None
        try:
            work_area = win32util.monitor_work_area_for_window(
                self.master.winfo_id())
        except Exception:
            work_area = None
        if work_area:
            left, top, right, bottom = work_area
            screen_w, screen_h = right - left, bottom - top
        else:
            left = top = 0
            screen_w = self.winfo_screenwidth()
            screen_h = self.winfo_screenheight()
        available_w = max(320, screen_w - 40)
        available_h = max(240, screen_h - 80)
        width = min(width, available_w)
        height = min(height, available_h)
        self.minsize(min(480, width), min(360, height))
        x = left + max(0, (screen_w - width) // 2)
        y = top + max(0, (screen_h - height) // 2 - 30)
        self.geometry("%dx%d+%d+%d" % (width, height, x, y))

    def _cancel(self):
        self.result = None
        self.destroy()

    def show(self):
        # Subclasses build their widgets after BaseDialog.__init__; process
        # geometry before entering the modal wait loop, otherwise the dialog
        # can be shown as an empty 1x1 frame on Windows.
        self.update_idletasks()
        self.grab_set()
        try:
            self.lift()
            self.focus_force()
        except Exception:
            pass
        self.wait_window(self)
        return self.result


def _button_row(parent, buttons, scale=1.0):
    row = tk.Frame(parent, bg=BG)
    row.pack(fill="x", pady=(12, 0))
    for text, command, primary in buttons:
        if primary:
            btn = tk.Button(
                row, text=text, command=command, relief="flat", bd=0,
                bg="#1570EF", fg="white", activebackground="#0B5BC4",
                activeforeground="white", font=FONT_BOLD,
                padx=int(16 * scale), pady=int(7 * scale), cursor="hand2")
        else:
            btn = tk.Button(
                row, text=text, command=command, relief="solid", bd=1,
                bg="#FFFFFF", fg="#344054", activebackground="#F2F4F7",
                font=FONT, padx=int(14 * scale), pady=int(6 * scale),
                cursor="hand2")
        btn.pack(side="right", padx=(8, 0))
    return row


def _wrapped_text(parent, text: str, height: int, bg: str = BG,
                  fg: str = "#344054", font=FONT):
    """Read-only character-wrapped text; avoids CJK clipping in Label widgets."""
    widget = tk.Text(
        parent, height=height, width=1, wrap="char", font=font, bg=bg, fg=fg,
        relief="flat", bd=0, highlightthickness=0, padx=0, pady=0)
    widget.insert("1.0", text)
    widget.configure(state="disabled")
    return widget


class FirstRunDialog(BaseDialog):
    """Non-destructive, three-step first-run guide."""

    def __init__(self, app):
        scale = app.scale
        super().__init__(app, "首次使用引导", int(700 * scale), int(500 * scale))
        self.page = 0
        self.path_var = tk.StringVar(value="练习底稿尚未生成。")
        self.body = tk.Frame(self, bg=BG, padx=int(22 * scale),
                             pady=int(18 * scale))
        self.body.pack(fill="both", expand=True)
        self.page_label = tk.Label(
            self.body, bg=BG, fg="#1570EF", font=FONT_SMALL, anchor="w")
        self.page_label.pack(fill="x")
        self.content = tk.Frame(self.body, bg=BG)
        self.content.pack(fill="both", expand=True, pady=(10, 0))

        footer = tk.Frame(self.body, bg=BG)
        footer.pack(fill="x", pady=(12, 0))
        tk.Button(footer, text="稍后", command=self._defer, bg="#FFFFFF",
                  fg="#344054", relief="solid", bd=1, font=FONT,
                  padx=12, pady=5, cursor="hand2").pack(side="left")
        self.back_button = tk.Button(
            footer, text="上一步", command=self._previous, bg="#FFFFFF",
            fg="#344054", relief="solid", bd=1, font=FONT,
            padx=12, pady=5, cursor="hand2")
        self.back_button.pack(side="right", padx=(8, 0))
        self.next_button = tk.Button(
            footer, text="下一步", command=self._next, bg="#1570EF",
            fg="white", activebackground="#0B5BC4", relief="flat", bd=0,
            font=FONT_BOLD, padx=14, pady=6, cursor="hand2")
        self.next_button.pack(side="right")
        self._render()

    def _render(self):
        for child in self.content.winfo_children():
            child.destroy()
        titles = ["1 / 3 · 本地数据与热键",
                  "2 / 3 · 用合成底稿练习登记",
                  "3 / 3 · 定位并手动更新状态"]
        self.page_label.configure(text=titles[self.page])
        if self.page == 0:
            text = (
                "数据目录：\n%s\n\n"
                "事项库是本地 SQLite 明文文件，包含程序文字和文件路径；"
                "请按事务所资料管理要求保管。助手不会上传数据，也不会向原底稿写入内容、保存或覆盖底稿。\n\n"
                "%s" % (self.app.data_dir, self.app.hotkey_status_text()))
            _wrapped_text(self.content, text, 10).pack(
                fill="both", expand=True)
            tk.Button(self.content, text="打开热键设置…",
                      command=self._open_settings, bg="#FFFFFF", fg="#344054",
                      relief="solid", bd=1, font=FONT_SMALL, padx=10, pady=4,
                      cursor="hand2").pack(anchor="w", pady=(10, 0))
        elif self.page == 1:
            text = (
                "点击下方按钮生成并打开一份本机练习文件。文件只包含虚构的审计程序，"
                "不会覆盖已有练习文件。\n\n"
                "在练习底稿中选中 B5，按 %s。成功提示可用“撤销”移除练习事项。\n"
                "登记是用户主动操作；启动引导不会自动打开表格或发送按键。" %
                self.app.hotkey_text)
            _wrapped_text(self.content, text, 7).pack(fill="x")
            tk.Button(
                self.content, text="生成并打开练习底稿",
                command=self._open_practice, bg="#12B76A", fg="white",
                activebackground="#0E9F5B", relief="flat", bd=0,
                font=FONT_BOLD, padx=14, pady=7, cursor="hand2").pack(
                    anchor="w", pady=(16, 8))
            tk.Label(self.content, textvariable=self.path_var, bg=BG,
                     fg="#667085", font=FONT_SMALL, wraplength=int(640 * self.app.scale),
                     justify="left", anchor="w").pack(fill="x")
        else:
            text = (
                "选择清单中的事项，双击或点“定位底稿”返回对应工作表和单元格。\n\n"
                "实际取得资料或执行程序后，再由您手动修改状态。收到日期只是记录；"
                "系统不会因为收到资料、单元格文字变化或定位成功而自动标记完成。\n\n"
                "练习结束后，请在练习事项提示条上点“撤销”；不要把练习事项当作正式项目记录。")
            _wrapped_text(self.content, text, 8).pack(
                fill="both", expand=True)
        self.back_button.configure(
            state=tk.NORMAL if self.page > 0 else tk.DISABLED)
        self.next_button.configure(
            text="完成" if self.page == 2 else "下一步",
            command=self._complete if self.page == 2 else self._next)

    def _open_practice(self):
        # Release the modal grab so the global hotkey can be tested in Excel/WPS.
        try:
            self.grab_release()
        except Exception:
            pass
        try:
            path = self.app.open_practice_workbook()
            self.path_var.set("已生成：%s" % path)
        except Exception as exc:
            messagebox.showerror("练习底稿生成失败", str(exc), parent=self)

    def _open_settings(self):
        self.result = "settings"
        self.destroy()

    def _previous(self):
        self.page = max(0, self.page - 1)
        self._render()

    def _next(self):
        self.page = min(2, self.page + 1)
        self._render()

    def _complete(self):
        self.app.dismiss_onboarding()
        self.result = "complete"
        self.destroy()

    def _defer(self):
        self.result = "later"
        self.destroy()


class DuplicateDialog(BaseDialog):
    def __init__(self, app, existing: Item):
        scale = app.scale
        super().__init__(app, "该位置已登记", int(560 * scale), int(330 * scale))
        body = tk.Frame(self, bg=BG, padx=int(20 * scale), pady=int(16 * scale))
        body.pack(fill="both", expand=True)
        tk.Label(body, text="该位置已有一条未完成事项", bg=BG, fg="#B42318",
                 font=FONT_TITLE, anchor="w").pack(fill="x")
        tk.Label(
            body,
            text="%s  ·  %s!%s" % (existing.file_name, existing.sheet_name,
                                   display_address(existing.cell_address)),
            bg=BG, fg="#344054", font=FONT_BOLD, anchor="w").pack(fill="x", pady=(6, 4))
        _wrapped_text(
            body,
            existing.program_text[:160] + ("…" if len(existing.program_text) > 160 else ""),
            3, fg="#475467").pack(fill="x")
        tk.Label(
            body, text="登记时间：%s    状态：%s" % (existing.created_at, existing.status),
            bg=BG, fg="#667085", font=FONT_SMALL, anchor="w").pack(fill="x", pady=(8, 0))
        _button_row(body, [
            ("取消", self._cancel, False),
            ("仍然新增", self._add, False),
            ("更新现有事项", self._update, False),
            ("查看现有事项", self._existing, True),
        ], scale)

    def _existing(self):
        self.result = "existing"
        self.destroy()

    def _add(self):
        self.result = "add"
        self.destroy()

    def _update(self):
        self.result = "update"
        self.destroy()


class MismatchDialog(BaseDialog):
    def __init__(self, app, item: Item, old: str, new: str,
                 relink_from: Optional[str] = None):
        scale = app.scale
        super().__init__(app, "单元格内容已变化", int(680 * scale), int(520 * scale))
        body = tk.Frame(self, bg=BG, padx=int(20 * scale), pady=int(16 * scale))
        body.pack(fill="both", expand=True)
        tk.Label(body, text="登记位置的内容与登记时不一致", bg=BG, fg="#B54708",
                 font=FONT_TITLE, anchor="w").pack(fill="x")
        _wrapped_text(
            body, "可能底稿已修改，或该位置已不是原来的审计程序。请核对后再决定。",
            2, fg="#475467").pack(fill="x", pady=(4, 8))
        if relink_from:
            context = (
                "重新关联：\n原文件：%s\n目标文件：%s\n位置：%s!%s" % (
                    relink_from, item.file_path, item.sheet_name,
                    display_address(item.cell_address)))
        else:
            context = "文件：%s\n位置：%s!%s" % (
                item.file_path, item.sheet_name, display_address(item.cell_address))
        _wrapped_text(body, context, 4, fg="#667085",
                      font=FONT_SMALL).pack(fill="x", pady=(0, 8))

        def text_block(title, content, color):
            frame = tk.LabelFrame(body, text=title, bg=BG, fg="#344054",
                                  font=FONT_BOLD, padx=8, pady=6)
            frame.pack(fill="both", expand=True, pady=(0, 8))
            widget = tk.Text(frame, height=4, wrap="word", font=FONT,
                             bg="#FFFFFF", fg=color, relief="solid", bd=1)
            widget.pack(fill="both", expand=True)
            widget.insert("1.0", content or "（空）")
            widget.configure(state="disabled")

        text_block("登记时的内容", old, "#344054")
        text_block("当前单元格内容", new, "#B42318")
        _button_row(body, [
            ("取消", self._cancel, False),
            ("更新快照并定位", self._update, False),
            ("仍然定位", self._jump, True),
        ], scale)

    def _jump(self):
        self.result = "jump"
        self.destroy()

    def _update(self):
        self.result = "update"
        self.destroy()


class RelinkDialog(BaseDialog):
    """Explicitly compare source and candidate paths before attempting a relink."""

    def __init__(self, app, item: Item):
        scale = app.scale
        super().__init__(app, "重新关联底稿", int(760 * scale), int(700 * scale))
        self.item = item
        self.target_path_var = tk.StringVar(value="")
        self.identity_var = tk.StringVar(
            value="请选择目标文件。文件指纹只是辅助线索，不能保证版本正确。")
        body = tk.Frame(self, bg=BG, padx=int(20 * scale), pady=int(16 * scale))
        body.pack(fill="both", expand=True)
        tk.Label(body, text="确认新底稿文件", bg=BG, fg="#344054",
                 font=FONT_TITLE, anchor="w").pack(fill="x")
        _wrapped_text(
            body,
            "系统无法替您判断移动后的文件是否为正确版本。请核对旧路径、新路径、工作表、单元格和登记快照；"
            "确认后系统还会打开文件并核对目标单元格。取消或核对失败时不会更新事项路径。",
            4, fg="#475467").pack(fill="x", pady=(6, 12))

        paths = tk.LabelFrame(body, text="文件路径", bg=BG, fg="#344054",
                              font=FONT_BOLD, padx=10, pady=8)
        paths.pack(fill="x")
        tk.Label(paths, text="原路径（当前不存在）", bg=BG, fg="#667085",
                 font=FONT_SMALL, anchor="w").pack(fill="x")
        _wrapped_text(paths, item.file_path, 2, fg="#344054",
                      font=FONT_SMALL).pack(fill="x", pady=(2, 8))
        tk.Label(paths, text="新路径", bg=BG, fg="#667085",
                 font=FONT_SMALL, anchor="w").pack(fill="x")
        target_row = tk.Frame(paths, bg=BG)
        target_row.pack(fill="x", pady=(2, 6))
        ttk.Entry(target_row, textvariable=self.target_path_var,
                  state="readonly", font=FONT_SMALL).pack(
                      side="left", fill="x", expand=True)
        tk.Button(target_row, text="选择文件…", command=self._choose_file,
                  bg="#FFFFFF", fg="#344054", relief="solid", bd=1,
                  font=FONT_SMALL, padx=9, pady=3, cursor="hand2").pack(
                      side="left", padx=(8, 0))
        tk.Label(paths, textvariable=self.identity_var, bg=BG, fg="#B54708",
                 font=FONT_SMALL, anchor="w", justify="left",
                 wraplength=int(680 * scale)).pack(fill="x")

        tk.Label(
            body,
            text="位置：%s!%s%s" % (
                item.sheet_name, display_address(item.cell_address),
                ("（合并区域 %s）" % item.merge_range) if item.merge_range else ""),
            bg=BG, fg="#344054", font=FONT_BOLD, anchor="w").pack(
                fill="x", pady=(12, 6))
        snapshot_box = tk.LabelFrame(
            body, text="登记时的内容（之后会与新文件中的单元格核对）",
            bg=BG, fg="#344054", font=FONT_BOLD, padx=8, pady=6)
        snapshot_box.pack(fill="both", expand=True)
        snapshot = tk.Text(snapshot_box, height=5, wrap="word", font=FONT,
                           bg="#FFFFFF", fg="#344054", relief="solid", bd=1)
        snapshot.pack(fill="both", expand=True)
        snapshot.insert("1.0", item.snapshot or "（登记快照为空）")
        snapshot.configure(state="disabled")

        row = tk.Frame(body, bg=BG)
        row.pack(fill="x", pady=(12, 0))
        tk.Button(row, text="取消", command=self._cancel, bg="#FFFFFF",
                  fg="#344054", relief="solid", bd=1, font=FONT,
                  padx=12, pady=5, cursor="hand2").pack(side="right", padx=(8, 0))
        self.continue_button = tk.Button(
            row, text="打开并核对目标", command=self._continue,
            state=tk.DISABLED, bg="#1570EF", fg="white",
            activebackground="#0B5BC4", relief="flat", bd=0,
            font=FONT_BOLD, padx=14, pady=6, cursor="hand2")
        self.continue_button.pack(side="right")

    def _choose_file(self):
        path = filedialog.askopenfilename(
            parent=self, title="选择移动或改名后的底稿",
            initialdir=os.path.dirname(self.item.file_path) or None,
            filetypes=[("Excel 工作簿", "*.xlsx *.xlsm *.xls *.xltx *.xltm"),
                       ("所有文件", "*.*")])
        if not path:
            return
        self.target_path_var.set(path)
        try:
            volume, index = win32util.file_identity(path)
        except Exception:
            volume, index = "", ""
        if self.item.has_file_identity and volume and index:
            if (volume, index) == (self.item.file_volume, self.item.file_index):
                message = "文件指纹与原文件一致（辅助线索）；仍需核对单元格内容。"
            else:
                message = ("文件指纹不同，可能是复制件或另一个文件。系统无法判断版本正确性；"
                           "请核对路径、工作表和内容后再继续。")
        else:
            message = ("无法完整核对文件指纹。请人工核对路径、工作表和单元格内容；"
                       "系统不会自动判断文件版本。")
        self.identity_var.set(message)
        self.continue_button.configure(state=tk.NORMAL)

    def _continue(self):
        path = self.target_path_var.get().strip()
        if not path:
            return
        self.result = path
        self.destroy()


class QuickFollowUpDialog(BaseDialog):
    """Focused editor for missing-material follow-up fields."""

    def __init__(self, app, item: Item, received_today: bool = False):
        scale = app.scale
        super().__init__(app, "快速跟进", int(680 * scale), int(620 * scale))
        self.item = item
        outer = tk.Frame(self, bg=BG, padx=int(20 * scale), pady=int(16 * scale))
        outer.pack(fill="both", expand=True)
        tk.Label(outer, text="快速跟进资料", bg=BG, fg="#344054",
                 font=FONT_TITLE, anchor="w").pack(fill="x")
        _wrapped_text(
            outer, "%s · %s!%s" % (
                item.file_name, item.sheet_name, display_address(item.cell_address)),
            2, fg="#667085", font=FONT_SMALL).pack(fill="x", pady=(4, 10))

        tk.Label(outer, text="缺失资料", bg=BG, fg="#344054",
                 font=FONT_BOLD, anchor="w").pack(fill="x", pady=(0, 4))
        self.missing_text = tk.Text(
            outer, height=4, wrap="word", font=FONT, relief="solid", bd=1,
            bg="#FFFFFF")
        self.missing_text.pack(fill="x")
        self.missing_text.insert("1.0", item.missing_info)

        fields = tk.Frame(outer, bg=BG)
        fields.pack(fill="x", pady=(10, 0))
        fields.columnconfigure(1, weight=1)
        fields.columnconfigure(3, weight=1)
        tk.Label(fields, text="资料负责人", bg=BG, fg="#344054",
                 font=FONT_SMALL).grid(row=0, column=0, sticky="w")
        self.owner_var = tk.StringVar(value=item.missing_owner)
        ttk.Entry(fields, textvariable=self.owner_var, font=FONT_SMALL).grid(
            row=0, column=1, sticky="ew", padx=(6, 14))
        tk.Label(fields, text="索取日期", bg=BG, fg="#344054",
                 font=FONT_SMALL).grid(row=0, column=2, sticky="w")
        self.requested_var = tk.StringVar(value=item.missing_requested_at)
        ttk.Entry(fields, textvariable=self.requested_var, font=FONT_SMALL,
                  width=14).grid(row=0, column=3, sticky="ew", padx=(6, 0))

        tk.Label(fields, text="收到日期", bg=BG, fg="#344054",
                 font=FONT_SMALL).grid(row=1, column=0, sticky="w", pady=(10, 0))
        initial_received = (_dt.date.today().isoformat() if received_today
                            else item.missing_received_at)
        self.received_var = tk.StringVar(value=initial_received)
        ttk.Entry(fields, textvariable=self.received_var, font=FONT_SMALL,
                  width=14).grid(row=1, column=1, sticky="ew", padx=(6, 14),
                                 pady=(10, 0))
        tk.Button(
            fields, text="资料已收到（今天）", command=self._mark_received_today,
            bg="#FFFFFF", fg="#344054", relief="solid", bd=1,
            font=FONT_SMALL, padx=8, pady=3, cursor="hand2").grid(
                row=1, column=2, columnspan=2, sticky="w", pady=(10, 0))

        tk.Label(outer, text="保存后的状态（请手动确认）", bg=BG, fg="#344054",
                 font=FONT_BOLD, anchor="w").pack(fill="x", pady=(14, 4))
        self.status_var = tk.StringVar(value=item.status)
        ttk.Combobox(outer, textvariable=self.status_var, values=STATUSES,
                     state="readonly", font=FONT, width=18).pack(anchor="w")
        _wrapped_text(
            outer,
            "记录收到日期不会自动改变状态。请选择您判断后的状态；如程序仍未执行，可继续保留“等待资料”。\n日期格式：YYYY-MM-DD。",
            3, fg="#667085", font=FONT_SMALL).pack(fill="x", pady=(8, 0))
        _button_row(outer, [
            ("取消", self._cancel, False),
            ("保存跟进", self._save, True),
        ], scale)

    def _mark_received_today(self):
        self.received_var.set(_dt.date.today().isoformat())

    def _save(self):
        values = {
            "missing_info": self.missing_text.get("1.0", "end").rstrip(),
            "missing_owner": self.owner_var.get().strip(),
            "missing_requested_at": self.requested_var.get().strip(),
            "missing_received_at": self.received_var.get().strip(),
            "status": self.status_var.get().strip(),
        }
        for label, key in (("索取日期", "missing_requested_at"),
                           ("收到日期", "missing_received_at")):
            if values[key] and parse_date(values[key]) is None:
                messagebox.showwarning(
                    "日期格式无效", "%s 请使用 YYYY-MM-DD 格式，或留空。" % label,
                    parent=self)
                return
        if values["status"] not in STATUSES:
            messagebox.showwarning("状态无效", "请选择有效的状态。", parent=self)
            return
        changes = {key: value for key, value in values.items()
                   if getattr(self.item, key) != value}
        if changes:
            self.app.update_item(self.item.id, changes)
        self.result = "saved"
        self.destroy()


class DetailDialog(BaseDialog):
    def __init__(self, app, item: Item):
        scale = app.scale
        super().__init__(app, "事项详情", int(800 * scale), int(860 * scale))
        self.item = item
        outer = tk.Frame(self, bg=BG, padx=int(20 * scale), pady=int(14 * scale))
        outer.pack(fill="both", expand=True)

        top = tk.Frame(outer, bg=BG)
        top.pack(fill="x")
        tk.Label(top, text="项目", bg=BG, fg="#344054", font=FONT_BOLD).grid(
            row=0, column=0, sticky="w")
        self.project_var = tk.StringVar(value=item.project)
        project_box = ttk.Combobox(top, textvariable=self.project_var,
                                   values=app.project_values(), font=FONT, width=28)
        project_box.grid(row=0, column=1, sticky="w", padx=(8, 20))
        tk.Label(top, text="状态", bg=BG, fg="#344054", font=FONT_BOLD).grid(
            row=0, column=2, sticky="w")
        self.status_var = tk.StringVar(value=item.status)
        status_box = ttk.Combobox(top, textvariable=self.status_var,
                                  values=STATUSES, state="readonly", font=FONT, width=12)
        status_box.grid(row=0, column=3, sticky="w", padx=(8, 0))

        tk.Label(outer, text="审计程序（登记内容）", bg=BG, fg="#344054",
                 font=FONT_BOLD, anchor="w").pack(fill="x", pady=(12, 4))
        self.program_text = tk.Text(outer, height=5, wrap="word", font=FONT,
                                    relief="solid", bd=1, bg="#FFFFFF")
        self.program_text.pack(fill="x")
        self.program_text.insert("1.0", item.program_text)

        loc = tk.LabelFrame(outer, text=" 底稿位置 ", bg=BG, fg="#344054",
                            font=FONT_BOLD, padx=10, pady=8)
        loc.pack(fill="x", pady=(12, 0))
        line1 = tk.Frame(loc, bg=BG)
        line1.pack(fill="x")
        tk.Label(line1, text="文件：", bg=BG, fg="#667085", font=FONT).pack(side="left")
        self.path_var = tk.StringVar(value=item.file_path)
        path_entry = ttk.Entry(line1, textvariable=self.path_var, font=FONT_SMALL)
        path_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        path_entry.configure(state="readonly")
        tk.Button(line1, text="打开文件夹", command=self._open_folder,
                  relief="solid", bd=1, bg="#FFFFFF", fg="#344054",
                  font=FONT_SMALL, padx=8, pady=2, cursor="hand2").pack(side="right")
        merge_text = ("，合并区域 %s" % item.merge_range) if item.merge_range else ""
        wait_text = ""
        if item.status == STATUS_WAITING:
            days = item.waiting_days
            wait_text = "    等待：%s" % (
                "%d 天" % days if days is not None else "未记录起始日")
        tk.Label(
            loc,
            text="位置：%s!%s%s    来源：%s%s" % (
                item.sheet_name, display_address(item.cell_address), merge_text,
                item.source_name, wait_text),
            bg=BG, fg="#344054", font=FONT, anchor="w").pack(fill="x", pady=(6, 0))
        tk.Label(
            loc,
            text="登记：%s    更新：%s    完成：%s" % (
                item.created_at, item.updated_at, item.completed_at or "—"),
            bg=BG, fg="#667085", font=FONT_SMALL, anchor="w").pack(fill="x")

        tk.Label(outer, text="登记时的单元格内容（用于定位核对）", bg=BG, fg="#344054",
                 font=FONT_BOLD, anchor="w").pack(fill="x", pady=(12, 4))
        snapshot = tk.Text(outer, height=3, wrap="word", font=FONT_SMALL,
                           relief="solid", bd=1, bg="#F2F4F7", fg="#475467")
        snapshot.pack(fill="x")
        snapshot.insert("1.0", item.snapshot)
        snapshot.configure(state="disabled")

        missing_box = tk.LabelFrame(
            outer, text=" 缺失资料 / 下一步 ", bg=BG, fg="#344054",
            font=FONT_BOLD, padx=10, pady=8)
        missing_box.pack(fill="x", pady=(12, 0))
        missing_box.columnconfigure(1, weight=1)
        missing_box.columnconfigure(3, weight=1)
        missing_box.columnconfigure(5, weight=1)
        tk.Label(missing_box, text="资料负责人", bg=BG, fg="#667085",
                 font=FONT_SMALL).grid(row=0, column=0, sticky="w")
        self.missing_owner_var = tk.StringVar(value=item.missing_owner)
        ttk.Entry(missing_box, textvariable=self.missing_owner_var,
                  font=FONT_SMALL).grid(row=0, column=1, sticky="ew",
                                        padx=(6, 14))
        tk.Label(missing_box, text="索取日期", bg=BG, fg="#667085",
                 font=FONT_SMALL).grid(row=0, column=2, sticky="w")
        self.requested_var = tk.StringVar(value=item.missing_requested_at)
        ttk.Entry(missing_box, textvariable=self.requested_var,
                  font=FONT_SMALL, width=12).grid(row=0, column=3, sticky="ew",
                                                  padx=(6, 14))
        tk.Label(missing_box, text="收到日期", bg=BG, fg="#667085",
                 font=FONT_SMALL).grid(row=0, column=4, sticky="w")
        self.received_var = tk.StringVar(value=item.missing_received_at)
        ttk.Entry(missing_box, textvariable=self.received_var,
                  font=FONT_SMALL, width=12).grid(row=0, column=5, sticky="ew",
                                                  padx=(6, 0))
        tk.Label(missing_box, text="日期格式 YYYY-MM-DD，可留空；收到日期仅作记录，"
                                  "不会自动改变状态。", bg=BG, fg="#667085",
                 font=FONT_SMALL, anchor="w").grid(
                     row=1, column=0, columnspan=6, sticky="w", pady=(3, 4))
        self.missing_text = tk.Text(missing_box, height=3, wrap="word", font=FONT,
                                    relief="solid", bd=1, bg="#FFFFFF")
        self.missing_text.grid(row=2, column=0, columnspan=6, sticky="ew")
        self.missing_text.insert("1.0", item.missing_info)

        tk.Label(outer, text="备注", bg=BG, fg="#344054",
                 font=FONT_BOLD, anchor="w").pack(fill="x", pady=(10, 4))
        self.note_text = tk.Text(outer, height=3, wrap="word", font=FONT,
                                 relief="solid", bd=1, bg="#FFFFFF")
        self.note_text.pack(fill="x")
        self.note_text.insert("1.0", item.note)

        _button_row(outer, [
            ("取消", self._cancel, False),
            ("变更历史", self._history, False),
            ("删除", self._delete, False),
            ("定位底稿", self._locate, False),
            ("保存", self._save, True),
        ], scale)

    def _values(self):
        return {
            "project": self.project_var.get().strip(),
            "status": self.status_var.get().strip(),
            "program_text": self.program_text.get("1.0", "end").rstrip(),
            "missing_info": self.missing_text.get("1.0", "end").rstrip(),
            "missing_owner": self.missing_owner_var.get().strip(),
            "missing_requested_at": self.requested_var.get().strip(),
            "missing_received_at": self.received_var.get().strip(),
            "note": self.note_text.get("1.0", "end").rstrip(),
        }

    def _save(self, close: bool = True):
        values = self._values()
        if values["status"] not in STATUSES:
            messagebox.showwarning("状态无效", "请选择有效的状态。", parent=self)
            return
        for label, key in (("索取日期", "missing_requested_at"),
                           ("收到日期", "missing_received_at")):
            text = values[key]
            if text and parse_date(text) is None:
                messagebox.showwarning(
                    "日期格式无效",
                    "%s 请使用 YYYY-MM-DD 格式，或留空。" % label,
                    parent=self)
                return
        self.app.update_item(self.item.id, values)
        self.result = "saved"
        if close:
            self.destroy()

    def _history(self):
        self.grab_release()
        try:
            HistoryDialog(self.app, self.item).show()
        finally:
            try:
                self.grab_set()
                self.lift()
            except Exception:
                pass

    def _locate(self):
        self._save(close=False)
        self.result = "locate"
        self.destroy()

    def _delete(self):
        if messagebox.askyesno(
                "删除事项", "删除后不可恢复，确定删除这条事项吗？", parent=self):
            self.app.delete_item(self.item.id, refresh=False)
            self.result = "deleted"
            self.destroy()

    def _open_folder(self):
        folder = os.path.dirname(self.item.file_path)
        if folder and os.path.isdir(folder):
            subprocess.Popen(["explorer", folder])
        else:
            messagebox.showinfo("文件夹不存在", "该文件夹当前不存在。", parent=self)


def _clip(value, limit: int = 100) -> str:
    text = (value or "").replace("\r", " ").replace("\n", " ")
    return text if len(text) <= limit else text[: limit - 1] + "…"


class HistoryDialog(BaseDialog):
    def __init__(self, app, item: Item):
        scale = app.scale
        super().__init__(app, "事项变更历史", int(780 * scale), int(540 * scale))
        body = tk.Frame(self, bg=BG, padx=int(18 * scale), pady=int(14 * scale))
        body.pack(fill="both", expand=True)
        tk.Label(
            body, text="%s · %s!%s" % (
                item.file_name, item.sheet_name, display_address(item.cell_address)),
            bg=BG, fg="#344054", font=FONT_BOLD, anchor="w").pack(fill="x")
        tk.Label(
            body, text="记录登记、状态、项目、缺失资料、备注和文件路径等变更。",
            bg=BG, fg="#667085", font=FONT_SMALL, anchor="w").pack(
                fill="x", pady=(2, 8))

        wrap = tk.Frame(body, bg=BG)
        wrap.pack(fill="both", expand=True)
        columns = ("time", "event", "old", "new")
        tree = ttk.Treeview(wrap, columns=columns, show="headings",
                            selectmode="browse")
        for col, title, width in (
                ("time", "时间", 140), ("event", "操作", 90),
                ("old", "原值", 220), ("new", "新值", 220)):
            tree.heading(col, text=title)
            tree.column(col, width=int(width * scale), anchor="w",
                        stretch=col in ("old", "new"))
        for ev in reversed(app.store.list_events(item.id)):
            tree.insert("", "end", values=(
                ev.created_at, ev.label, _clip(ev.old_value),
                _clip(ev.new_value)))
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        _button_row(body, [("关闭", self._cancel, True)], scale)


class ProjectArchiveDialog(BaseDialog):
    def __init__(self, app):
        scale = app.scale
        super().__init__(app, "项目归档管理", int(720 * scale), int(540 * scale))
        self.app = app
        body = tk.Frame(self, bg=BG, padx=int(18 * scale), pady=int(14 * scale))
        body.pack(fill="both", expand=True)
        tk.Label(body, text="项目归档", bg=BG, fg="#344054",
                 font=FONT_TITLE, anchor="w").pack(fill="x")
        tk.Label(
            body, text="归档只影响清单的默认显示，不会删除事项；"
                       "需要查看时可选择“查看事项”。",
            bg=BG, fg="#667085", font=FONT_SMALL, anchor="w",
            justify="left", wraplength=int(650 * scale)).pack(fill="x", pady=(2, 8))

        wrap = tk.Frame(body, bg=BG)
        wrap.pack(fill="both", expand=True)
        columns = ("project", "state", "count")
        self.tree = ttk.Treeview(wrap, columns=columns, show="headings",
                                 selectmode="browse")
        for col, title, width in (
                ("project", "项目", 420), ("state", "状态", 100),
                ("count", "事项数", 80)):
            self.tree.heading(col, text=title)
            self.tree.column(col, width=int(width * scale), anchor="w",
                             stretch=col == "project")
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        self._rows = []
        self._refresh()
        _button_row(body, [
            ("关闭", self._cancel, False),
            ("归档 / 取消归档", self._toggle, False),
            ("查看事项", self._view, True),
        ], scale)

    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        self._rows = self.app.store.project_infos()
        for index, info in enumerate(self._rows):
            self.tree.insert(
                "", "end", iid=str(index),
                values=(info.name, "已归档" if info.archived else "进行中",
                        info.item_count))
        if self._rows:
            self.tree.selection_set("0")
            self.tree.focus("0")

    def _selected(self):
        selection = self.tree.selection()
        if not selection:
            return None
        try:
            return self._rows[int(selection[0])]
        except (ValueError, IndexError):
            return None

    def _toggle(self):
        info = self._selected()
        if info is None:
            return
        self.app.store.set_project_archived(info.name, not info.archived)
        self.app.refresh()
        self._refresh()

    def _view(self):
        info = self._selected()
        if info is None:
            return
        name = info.name
        self.destroy()
        self.app.show_project_items(name)


class SettingsDialog(BaseDialog):
    def __init__(self, app):
        scale = app.scale
        super().__init__(app, "设置", int(720 * scale), int(720 * scale))
        body = tk.Frame(self, bg=BG, padx=int(20 * scale), pady=int(16 * scale))
        body.pack(fill="both", expand=True)

        tk.Label(body, text="登记热键", bg=BG, fg="#344054",
                 font=FONT_BOLD, anchor="w").pack(fill="x")
        hotkey_row = tk.Frame(body, bg=BG)
        hotkey_row.pack(fill="x", pady=(6, 0))
        self.hotkey_var = tk.StringVar(value=app.hotkey_text)
        box = ttk.Combobox(hotkey_row, textvariable=self.hotkey_var,
                           values=HOTKEY_PRESETS, font=FONT, width=20)
        box.pack(side="left")
        tk.Button(hotkey_row, text="应用热键", command=self._apply_hotkey,
                  relief="flat", bd=0, bg="#1570EF", fg="white",
                  activebackground="#0B5BC4", activeforeground="white",
                  font=FONT_BOLD, padx=12, pady=4, cursor="hand2").pack(
                      side="left", padx=(10, 0))
        tk.Label(body, text="在 Excel/WPS 中选中单元格后按此热键即可登记。"
                            "若热键被其他软件占用，请换一个组合。",
                 bg=BG, fg="#667085", font=FONT_SMALL, anchor="w",
                 justify="left", wraplength=int(660 * scale)).pack(fill="x", pady=(4, 0))

        tk.Label(body, text="界面与项目", bg=BG, fg="#344054",
                 font=FONT_BOLD, anchor="w").pack(fill="x", pady=(14, 4))
        self.close_tray_var = tk.BooleanVar(
            value=bool(app.config.get("close_to_tray", True)))
        tk.Checkbutton(
            body, text="关闭主窗口时最小化到托盘，热键继续可用",
            variable=self.close_tray_var, command=self._close_tray_changed,
            bg=BG, fg="#344054", activebackground=BG, font=FONT,
            anchor="w").pack(fill="x")
        self.auto_project_var = tk.BooleanVar(
            value=bool(app.config.get("auto_project_from_folder", True)))
        tk.Checkbutton(
            body, text="未选择项目时，按底稿所在文件夹名自动建议项目",
            variable=self.auto_project_var,
            command=lambda: app.config.set(
                "auto_project_from_folder", self.auto_project_var.get()),
            bg=BG, fg="#344054", activebackground=BG, font=FONT,
            anchor="w").pack(fill="x")
        tk.Button(body, text="项目归档管理…", command=self._project_archive,
                  relief="solid", bd=1, bg="#FFFFFF", fg="#344054", font=FONT,
                  padx=12, pady=4, cursor="hand2").pack(anchor="w", pady=(6, 0))

        tk.Label(body, text="数据与备份", bg=BG, fg="#344054",
                 font=FONT_BOLD, anchor="w").pack(fill="x", pady=(14, 4))
        path_row = tk.Frame(body, bg=BG)
        path_row.pack(fill="x")
        self.data_var = tk.StringVar(value=app.data_dir)
        entry = ttk.Entry(path_row, textvariable=self.data_var, font=FONT_SMALL)
        entry.pack(side="left", fill="x", expand=True)
        entry.configure(state="readonly")
        tk.Button(path_row, text="打开文件夹", command=app.open_data_folder,
                  relief="solid", bd=1, bg="#FFFFFF", fg="#344054", font=FONT_SMALL,
                  padx=8, pady=2, cursor="hand2").pack(side="right", padx=(8, 0))
        self.auto_backup_var = tk.BooleanVar(
            value=bool(app.config.get("auto_backup_enabled", True)))
        tk.Checkbutton(
            body, text="每天自动备份到数据目录下的 backups 文件夹",
            variable=self.auto_backup_var,
            command=lambda: app.config.set(
                "auto_backup_enabled", self.auto_backup_var.get()),
            bg=BG, fg="#344054", activebackground=BG, font=FONT,
            anchor="w").pack(fill="x", pady=(6, 0))
        backup_row = tk.Frame(body, bg=BG)
        backup_row.pack(fill="x", pady=(6, 0))
        for text, command in (
                ("普通备份…", self._backup),
                ("加密备份…", self._encrypted_backup),
                ("解密备份…", self._decrypt_backup),
                ("更改数据目录…", self._change_data_dir)):
            tk.Button(backup_row, text=text, command=command,
                      relief="solid", bd=1, bg="#FFFFFF", fg="#344054",
                      font=FONT, padx=10, pady=4, cursor="hand2").pack(
                          side="left", padx=(0, 8))
        tk.Label(
            body,
            text="加密备份使用 Windows 当前用户密钥，只能在同一台电脑、"
                 "同一 Windows 用户下恢复；请同时保留普通备份。",
            bg=BG, fg="#667085", font=FONT_SMALL, anchor="w",
            justify="left", wraplength=int(660 * scale)).pack(fill="x", pady=(4, 0))
        log_row = tk.Frame(body, bg=BG)
        log_row.pack(fill="x", pady=(6, 0))
        tk.Button(log_row, text="打开日志文件夹", command=app.open_log_folder,
                  relief="flat", bd=0, bg=BG, fg="#1570EF", font=FONT_SMALL,
                  cursor="hand2").pack(side="left")
        tk.Button(log_row, text="清空日志", command=self._clear_logs,
                  relief="flat", bd=0, bg=BG, fg="#667085", font=FONT_SMALL,
                  cursor="hand2").pack(side="left", padx=(14, 0))

        about = tk.LabelFrame(body, text=" 说明 ", bg=BG, fg="#344054",
                              font=FONT_BOLD, padx=10, pady=8)
        about.pack(fill="both", expand=True, pady=(12, 0))
        text = tk.Text(about, wrap="word", font=FONT_SMALL, bg="#FFFFFF",
                       fg="#475467", relief="flat", height=6)
        text.pack(fill="both", expand=True)
        text.insert("1.0", (
            "底稿未完成助手 %s\n\n"
            "• 数据只保存在本机，不会上传任何底稿或单元格内容。\n"
            "• 登记和定位只读取工作簿，不写入标记、不保存、不关闭您的底稿。\n"
            "• 事项状态需要您手动更新，工具不会自动判定“已完成”。\n"
            "• 数据为本地明文 SQLite 文件，请按事务所资料管理要求妥善保管和备份。\n"
            "• 支持 Windows 上的 Microsoft Excel 与 WPS 表格；"
            "单元格必须是单个单元格或合并单元格。"
        ) % __version__)
        text.configure(state="disabled")

        _button_row(body, [("关闭", self._cancel, True)], scale)

    def _apply_hotkey(self):
        ok, message = self.app.apply_hotkey(self.hotkey_var.get().strip())
        if ok:
            messagebox.showinfo("热键已更新", "登记热键已改为 %s。" % self.hotkey_var.get(),
                                parent=self)
        else:
            messagebox.showwarning("热键不可用", message, parent=self)

    def _backup(self):
        default_name = "底稿未完成事项备份_%s.db" % _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = filedialog.asksaveasfilename(
            parent=self, title="备份数据文件",
            initialdir=os.path.join(self.app.data_dir, "backups"),
            initialfile=default_name, defaultextension=".db",
            filetypes=[("SQLite 数据库", "*.db"), ("所有文件", "*.*")])
        if not dest:
            return
        try:
            self.app.backup_data(dest)
            messagebox.showinfo("备份完成", "数据已备份到：\n%s" % dest, parent=self)
        except Exception as exc:
            messagebox.showerror("备份失败", str(exc), parent=self)

    def _encrypted_backup(self):
        self.app.encrypted_backup(self)

    def _decrypt_backup(self):
        self.app.decrypt_backup(self)

    def _change_data_dir(self):
        self.app.change_data_dir(self)

    def _close_tray_changed(self):
        self.app.set_close_to_tray(self.close_tray_var.get())

    def _project_archive(self):
        self.grab_release()
        try:
            self.app.open_project_archive()
        finally:
            try:
                self.grab_set()
                self.lift()
            except Exception:
                pass

    def _clear_logs(self):
        self.app.clear_logs(self)
