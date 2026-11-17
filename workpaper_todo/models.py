# -*- coding: utf-8 -*-
"""数据模型和状态常量。"""
from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from typing import Optional

STATUS_TODO = "未完成"
STATUS_WAITING = "等待资料"
STATUS_DOING = "进行中"
STATUS_DONE = "已完成"

STATUSES = [STATUS_TODO, STATUS_WAITING, STATUS_DOING, STATUS_DONE]
STATUS_PRIORITY = {STATUS_TODO: 0, STATUS_WAITING: 1, STATUS_DOING: 2, STATUS_DONE: 3}
STATUS_TODO_FILTER = "待办"

SOURCE_EXCEL = "excel"
SOURCE_WPS = "wps"
SOURCE_NAMES = {SOURCE_EXCEL: "Excel", SOURCE_WPS: "WPS 表格"}

EVENT_LABELS = {
    "created": "登记",
    "status": "状态",
    "project": "项目",
    "program_text": "审计程序",
    "snapshot": "内容快照",
    "file_path": "文件路径",
    "file_name": "文件名",
    "merge_range": "合并区域",
    "missing_info": "缺失资料",
    "missing_owner": "资料负责人",
    "missing_requested_at": "索取日期",
    "missing_received_at": "收到日期",
    "note": "备注",
    "relocate": "重新关联",
    "locate": "定位",
}


def display_address(address: str) -> str:
    return (address or "").replace("$", "")


def parse_date(value: Optional[str]) -> Optional[_dt.date]:
    """Parse the date forms used by the UI; return None for blank/invalid."""
    text = (value or "").strip()
    if not text:
        return None
    try:
        return _dt.date.fromisoformat(text[:10])
    except ValueError:
        return None


def waiting_days(start: Optional[str], today: Optional[_dt.date] = None) -> Optional[int]:
    day = parse_date(start)
    if day is None:
        return None
    return max(0, ((today or _dt.date.today()) - day).days)


@dataclass
class SelectionInfo:
    file_path: str
    file_key: str
    file_name: str
    sheet_name: str
    cell_address: str          # anchor, e.g. A5 (no $)
    cell_address_abs: str      # e.g. $A$5 or $A$9:$C$9
    merge_range: Optional[str]  # e.g. A9:C9 or None
    snapshot: str
    source_app: str
    source_exe: str
    window_hwnd: int = 0
    captured_at: str = ""
    file_volume: str = ""
    file_index: str = ""


@dataclass
class Item:
    id: int = 0
    project: str = ""
    program_text: str = ""
    status: str = STATUS_TODO
    file_path: str = ""
    file_key: str = ""
    file_name: str = ""
    sheet_name: str = ""
    cell_address: str = ""
    merge_range: Optional[str] = None
    snapshot: str = ""
    missing_info: str = ""
    missing_owner: str = ""
    missing_requested_at: str = ""
    missing_received_at: str = ""
    note: str = ""
    source_app: str = ""
    source_exe: str = ""
    file_volume: str = ""
    file_index: str = ""
    waiting_since: Optional[str] = None
    created_at: str = ""
    updated_at: str = ""
    completed_at: Optional[str] = None

    @property
    def location_text(self) -> str:
        return "%s!%s" % (self.sheet_name, display_address(self.cell_address))

    @property
    def source_name(self) -> str:
        return SOURCE_NAMES.get(self.source_app, self.source_app or "未知")

    @property
    def waiting_start(self) -> Optional[str]:
        return self.missing_requested_at or self.waiting_since or self.updated_at

    @property
    def waiting_days(self) -> Optional[int]:
        if self.status != STATUS_WAITING:
            return None
        return waiting_days(self.waiting_start)

    @property
    def has_file_identity(self) -> bool:
        return bool(self.file_volume and self.file_index)

    @classmethod
    def from_row(cls, row) -> "Item":
        data = dict(row)
        return cls(**{k: data.get(k) for k in cls.__dataclass_fields__ if k in data})


@dataclass
class ItemEvent:
    id: int = 0
    item_id: int = 0
    event_type: str = ""
    old_value: Optional[str] = None
    new_value: Optional[str] = None
    created_at: str = ""

    @property
    def label(self) -> str:
        return EVENT_LABELS.get(self.event_type, self.event_type or "变更")

    @classmethod
    def from_row(cls, row) -> "ItemEvent":
        data = dict(row)
        return cls(**{k: data.get(k) for k in cls.__dataclass_fields__ if k in data})


@dataclass
class ProjectInfo:
    name: str
    archived: bool = False
    item_count: int = 0
