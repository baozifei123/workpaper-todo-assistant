# -*- coding: utf-8 -*-
"""应用配置（JSON，保存在本地数据目录）。"""
from __future__ import annotations

import json
import os
import shutil
from typing import Any, Dict

DEFAULT_HOTKEY = "Ctrl+Alt+M"

DEFAULTS: Dict[str, Any] = {
    "hotkey": DEFAULT_HOTKEY,
    "current_project": "",
    "filter_status": "待办",
    "filter_project": "全部",
    "filter_aging": "全部",
    "include_archived": False,
    "search_text": "",
    "window_geometry": "",
    "last_export_dir": "",
    "close_to_tray": True,
    "auto_project_from_folder": True,
    "auto_backup_enabled": True,
    "auto_backup_keep": 7,
    "last_auto_backup": "",
    "tray_hint_shown": False,
    "onboarding_completed": False,
    "quick_view": "待办",
    "sort_waiting": "",
}


def default_root() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "WorkpaperTodo")


def _location_file() -> str:
    return os.path.join(default_root(), "location.txt")


def default_data_dir() -> str:
    override = os.environ.get("WORKPAPER_TODO_DATA_DIR")
    if override:
        return override
    root = default_root()
    pointer = _location_file()
    if os.path.isfile(pointer):
        try:
            with open(pointer, "r", encoding="utf-8") as f:
                target = f.read().strip()
            if target and os.path.isdir(target):
                return target
        except Exception:
            pass
    return root


def set_data_dir(path: str) -> None:
    """Record a custom data directory; a restart picks it up."""
    root = default_root()
    os.makedirs(root, exist_ok=True)
    pointer = _location_file()
    target = os.path.abspath(path)
    if os.path.normcase(target) == os.path.normcase(os.path.abspath(root)):
        clear_data_dir()
        return
    tmp = pointer + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(target)
    os.replace(tmp, pointer)


def clear_data_dir() -> None:
    try:
        os.remove(_location_file())
    except FileNotFoundError:
        pass
    except OSError:
        pass


class Config:
    def __init__(self, data_dir: str):
        self.data_dir = data_dir
        os.makedirs(data_dir, exist_ok=True)
        self.path = os.path.join(data_dir, "config.json")
        self.data: Dict[str, Any] = dict(DEFAULTS)
        self.load()

    def load(self) -> None:
        if not os.path.exists(self.path):
            self.save()
            return
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                # Preserve pre-quick-view status/aging filters in existing configs.
                if "quick_view" not in loaded:
                    self.data["quick_view"] = "自定义"
                self.data.update(loaded)
        except Exception:
            # Keep a copy of a corrupt config for diagnostics.
            try:
                shutil.copyfile(self.path, self.path + ".bad")
            except Exception:
                pass
            self.data = dict(DEFAULTS)
            self.save()

    def save(self) -> None:
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, self.path)

    def get(self, key: str, default=None):
        return self.data.get(key, DEFAULTS.get(key, default))

    def set(self, key: str, value) -> None:
        self.data[key] = value
        self.save()
