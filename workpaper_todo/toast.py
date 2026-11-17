# -*- coding: utf-8 -*-
"""右下角轻量提示条（不打断当前操作）。"""
from __future__ import annotations

import tkinter as tk

_KIND_COLORS = {
    "success": "#12B76A",
    "info": "#2E90FA",
    "warning": "#F79009",
    "error": "#F04438",
}


class Toast:
    def __init__(self, root: tk.Misc, scale: float = 1.0):
        self.root = root
        self.scale = scale
        self.win = None
        self._after_id = None
        self._action = None
        self._secondary_action = None

    def show(self, title: str, message: str = "", kind: str = "info",
             duration: int = 3800, action_text: str = "",
             action=None, secondary_action_text: str = "",
             secondary_action=None):
        self.close()
        color = _KIND_COLORS.get(kind, _KIND_COLORS["info"])
        px = max(1, int(round(self.scale)))
        win = tk.Toplevel(self.root)
        win.withdraw()
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        try:
            win.attributes("-alpha", 0.98)
        except Exception:
            pass

        outer = tk.Frame(win, bg="#D0D5DD", padx=1, pady=1)
        outer.pack(fill="both", expand=True)
        card = tk.Frame(outer, bg="#FFFFFF")
        card.pack(fill="both", expand=True)
        strip = tk.Frame(card, bg=color, width=5 * px)
        strip.pack(side="left", fill="y")
        body = tk.Frame(card, bg="#FFFFFF", padx=14 * px, pady=10 * px)
        body.pack(side="left", fill="both", expand=True)

        tk.Label(body, text=title, bg="#FFFFFF", fg="#101828",
                 font=("Microsoft YaHei UI", 10, "bold"), anchor="w",
                 justify="left").pack(fill="x")
        if message:
            tk.Label(body, text=message, bg="#FFFFFF", fg="#475467",
                     font=("Microsoft YaHei UI", 9), anchor="w", justify="left",
                     wraplength=int(360 * self.scale)).pack(fill="x", pady=(3 * px, 0))
        actions = []
        if action_text and action:
            actions.append((action_text, action))
        if secondary_action_text and secondary_action:
            actions.append((secondary_action_text, secondary_action))
        if actions:
            self._action = action
            self._secondary_action = secondary_action
            action_row = tk.Frame(body, bg="#FFFFFF")
            action_row.pack(anchor="e", fill="x", pady=(4 * px, 0))
            for text, callback in actions:
                tk.Button(
                    action_row, text=text,
                    command=lambda cb=callback: self._run_action(cb),
                    bg="#FFFFFF", fg="#1570EF", activebackground="#F2F4F7",
                    activeforeground="#0B5BC4", relief="flat", bd=0,
                    font=("Microsoft YaHei UI", 9, "bold"), cursor="hand2",
                    padx=6 * px, pady=2).pack(side="right", padx=(4 * px, 0))

        win.update_idletasks()
        width = max(int(300 * self.scale), win.winfo_reqwidth())
        height = win.winfo_reqheight()
        screen_w = win.winfo_screenwidth()
        screen_h = win.winfo_screenheight()
        x = max(10, screen_w - width - int(26 * self.scale))
        y = max(10, screen_h - height - int(70 * self.scale))
        win.geometry("%dx%d+%d+%d" % (width, height, x, y))
        win.deiconify()
        self.win = win
        self._after_id = self.root.after(duration, self.close)

    def _run_action(self, selected_action=None):
        action = selected_action or self._action
        self.close()
        if action is not None:
            try:
                action()
            except Exception:
                pass

    def close(self):
        if self._after_id is not None:
            try:
                self.root.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        if self.win is not None:
            try:
                self.win.destroy()
            except Exception:
                pass
            self.win = None
        self._action = None
        self._secondary_action = None
