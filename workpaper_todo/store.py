# -*- coding: utf-8 -*-
"""SQLite 本地存储。"""
from __future__ import annotations

import csv
import datetime as _dt
import html
import logging
import os
from pathlib import Path
import shutil
import sqlite3
from typing import Dict, List, Optional

from . import nativeom, win32util
from .models import (
    Item, ItemEvent, ProjectInfo, SelectionInfo, STATUS_DONE,
    STATUS_PRIORITY, STATUS_TODO, STATUS_TODO_FILTER, STATUS_WAITING, STATUSES,
    display_address,
)

log = logging.getLogger(__name__)

SCHEMA_VERSION = 2

_ITEM_FIELDS = [
    "project", "program_text", "status", "file_path", "file_key", "file_name",
    "sheet_name", "cell_address", "merge_range", "snapshot", "missing_info",
    "missing_owner", "missing_requested_at", "missing_received_at", "note",
    "source_app", "source_exe", "file_volume", "file_index",
]

_EVENT_FIELDS = {
    "project", "program_text", "status", "file_path", "file_name",
    "merge_range", "snapshot", "missing_info", "missing_owner",
    "missing_requested_at", "missing_received_at", "note", "file_volume",
    "file_index",
}

_CREATE_SQL = """
CREATE TABLE IF NOT EXISTS items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project TEXT NOT NULL DEFAULT '',
    program_text TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT '未完成',
    file_path TEXT NOT NULL DEFAULT '',
    file_key TEXT NOT NULL DEFAULT '',
    file_name TEXT NOT NULL DEFAULT '',
    sheet_name TEXT NOT NULL DEFAULT '',
    cell_address TEXT NOT NULL DEFAULT '',
    merge_range TEXT,
    snapshot TEXT NOT NULL DEFAULT '',
    missing_info TEXT NOT NULL DEFAULT '',
    missing_owner TEXT NOT NULL DEFAULT '',
    missing_requested_at TEXT NOT NULL DEFAULT '',
    missing_received_at TEXT NOT NULL DEFAULT '',
    note TEXT NOT NULL DEFAULT '',
    source_app TEXT NOT NULL DEFAULT '',
    source_exe TEXT NOT NULL DEFAULT '',
    file_volume TEXT NOT NULL DEFAULT '',
    file_index TEXT NOT NULL DEFAULT '',
    waiting_since TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    completed_at TEXT
);
CREATE INDEX IF NOT EXISTS idx_items_key
    ON items(file_key, sheet_name, cell_address);
CREATE INDEX IF NOT EXISTS idx_items_status ON items(status);
CREATE INDEX IF NOT EXISTS idx_items_file_identity
    ON items(file_volume, file_index);

CREATE TABLE IF NOT EXISTS item_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    item_id INTEGER NOT NULL,
    event_type TEXT NOT NULL,
    old_value TEXT,
    new_value TEXT,
    created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_events_item ON item_events(item_id, id);

CREATE TABLE IF NOT EXISTS project_meta (
    name TEXT PRIMARY KEY,
    archived INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
"""

_FTS_SQL = """
CREATE VIRTUAL TABLE IF NOT EXISTS items_fts USING fts5(
    program_text, project, file_name, sheet_name, cell_address,
    missing_info, missing_owner, note,
    content='items', content_rowid='id',
    tokenize='trigram'
);
CREATE TRIGGER IF NOT EXISTS items_fts_ai AFTER INSERT ON items BEGIN
    INSERT INTO items_fts(
        rowid, program_text, project, file_name, sheet_name, cell_address,
        missing_info, missing_owner, note)
    VALUES (new.id, new.program_text, new.project, new.file_name,
            new.sheet_name, new.cell_address, new.missing_info,
            new.missing_owner, new.note);
END;
CREATE TRIGGER IF NOT EXISTS items_fts_ad AFTER DELETE ON items BEGIN
    INSERT INTO items_fts(
        items_fts, rowid, program_text, project, file_name, sheet_name,
        cell_address, missing_info, missing_owner, note)
    VALUES ('delete', old.id, old.program_text, old.project, old.file_name,
            old.sheet_name, old.cell_address, old.missing_info,
            old.missing_owner, old.note);
END;
CREATE TRIGGER IF NOT EXISTS items_fts_au AFTER UPDATE ON items BEGIN
    INSERT INTO items_fts(
        items_fts, rowid, program_text, project, file_name, sheet_name,
        cell_address, missing_info, missing_owner, note)
    VALUES ('delete', old.id, old.program_text, old.project, old.file_name,
            old.sheet_name, old.cell_address, old.missing_info,
            old.missing_owner, old.note);
    INSERT INTO items_fts(
        rowid, program_text, project, file_name, sheet_name, cell_address,
        missing_info, missing_owner, note)
    VALUES (new.id, new.program_text, new.project, new.file_name,
            new.sheet_name, new.cell_address, new.missing_info,
            new.missing_owner, new.note);
END;
"""

_V1_COLUMNS = [
    ("missing_owner", "TEXT NOT NULL DEFAULT ''"),
    ("missing_requested_at", "TEXT NOT NULL DEFAULT ''"),
    ("missing_received_at", "TEXT NOT NULL DEFAULT ''"),
    ("file_volume", "TEXT NOT NULL DEFAULT ''"),
    ("file_index", "TEXT NOT NULL DEFAULT ''"),
    ("waiting_since", "TEXT"),
]


def now_str() -> str:
    return _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _csv_text(value):
    """Keep user-provided cells as text when a CSV is opened in a spreadsheet."""
    if isinstance(value, str) and (value.startswith(("\t", "\r", "\n")) or
                                   value.lstrip(" \t\r\n").startswith(("=", "+", "-", "@"))):
        return "'" + value
    return value


class Store:
    def __init__(self, db_path: str):
        self.db_path = db_path
        os.makedirs(os.path.dirname(db_path) or ".", exist_ok=True)
        self.conn = sqlite3.connect(db_path)
        self.conn.row_factory = sqlite3.Row
        self.fts_enabled = False
        self._init_schema()

    def _init_schema(self):
        version = int(self.conn.execute("PRAGMA user_version").fetchone()[0])
        if version > SCHEMA_VERSION:
            raise RuntimeError("数据文件版本较新，请升级底稿未完成助手。")
        if version == 0:
            self.conn.executescript(_CREATE_SQL)
            self.conn.execute("PRAGMA user_version = %d" % SCHEMA_VERSION)
            self.conn.commit()
            self._init_fts(recreate=True)
        elif version < SCHEMA_VERSION:
            self._backup_before_migration(version)
            self._migrate_v1_to_v2()
            self.conn.execute("PRAGMA user_version = %d" % SCHEMA_VERSION)
            self.conn.commit()
            self._init_fts(recreate=True)
        else:
            self._init_fts()

    def _backup_before_migration(self, version: int):
        if not os.path.exists(self.db_path):
            return
        try:
            shutil.copy2(self.db_path, "%s.v%d.bak" % (self.db_path, version))
        except Exception:
            log.exception("could not back up database before migration")

    def _migrate_v1_to_v2(self):
        columns = {row[1] for row in self.conn.execute("PRAGMA table_info(items)")}
        for name, ddl in _V1_COLUMNS:
            if name not in columns:
                self.conn.execute("ALTER TABLE items ADD COLUMN %s %s" % (name, ddl))
        self.conn.executescript("""
            CREATE TABLE IF NOT EXISTS item_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_id INTEGER NOT NULL,
                event_type TEXT NOT NULL,
                old_value TEXT,
                new_value TEXT,
                created_at TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_events_item
                ON item_events(item_id, id);
            CREATE TABLE IF NOT EXISTS project_meta (
                name TEXT PRIMARY KEY,
                archived INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_items_file_identity
                ON items(file_volume, file_index);
        """)
        now = now_str()
        self.conn.execute(
            "UPDATE items SET waiting_since = updated_at "
            "WHERE status = ? AND (waiting_since IS NULL OR waiting_since = '')",
            (STATUS_WAITING,))
        for row in self.conn.execute(
                "SELECT DISTINCT project FROM items WHERE TRIM(project) != ''"):
            self.conn.execute(
                "INSERT OR IGNORE INTO project_meta(name, archived, created_at, updated_at) "
                "VALUES (?, 0, ?, ?)", (row[0], now, now))
        self.conn.execute(
            "INSERT INTO item_events(item_id, event_type, old_value, new_value, created_at) "
            "SELECT id, 'created', NULL, COALESCE(NULLIF(project, ''), '未分类'), created_at "
            "FROM items")
        self.conn.commit()

    def _init_fts(self, recreate: bool = False):
        try:
            if recreate:
                self.conn.executescript("""
                    DROP TRIGGER IF EXISTS items_fts_ai;
                    DROP TRIGGER IF EXISTS items_fts_ad;
                    DROP TRIGGER IF EXISTS items_fts_au;
                    DROP TABLE IF EXISTS items_fts;
                """)
            self.conn.executescript(_FTS_SQL)
            if recreate:
                self.conn.execute(
                    "INSERT INTO items_fts(items_fts) VALUES ('rebuild')")
                self.conn.commit()
            self.fts_enabled = True
        except sqlite3.Error:
            self.fts_enabled = False
            log.exception("FTS5 unavailable; falling back to LIKE search")

    def close(self):
        try:
            self.conn.close()
        except Exception:
            pass

    # --- create/update -----------------------------------------------------
    def add_item(self, info: SelectionInfo, project: str) -> int:
        now = info.captured_at or now_str()
        project = project or ""
        self._ensure_project(project, now)
        cur = self.conn.execute(
            """INSERT INTO items
               (project, program_text, status, file_path, file_key, file_name,
                sheet_name, cell_address, merge_range, snapshot, missing_info,
                missing_owner, missing_requested_at, missing_received_at,
                note, source_app, source_exe, file_volume, file_index,
                created_at, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                project, info.snapshot, STATUS_TODO, info.file_path,
                info.file_key, info.file_name, info.sheet_name,
                info.cell_address, info.merge_range, info.snapshot, "",
                "", "", "", "", info.source_app, info.source_exe,
                info.file_volume, info.file_index, now, now,
            ))
        item_id = int(cur.lastrowid)
        self._insert_event(item_id, "created", None, project or "未分类", now)
        self.conn.commit()
        return item_id

    def get_item(self, item_id: int) -> Optional[Item]:
        row = self.conn.execute("SELECT * FROM items WHERE id = ?", (item_id,)).fetchone()
        return Item.from_row(row) if row else None

    def find_active_duplicate(self, file_key: str, sheet_name: str,
                              cell_address: str, file_volume: str = "",
                              file_index: str = "") -> Optional[Item]:
        identity = bool(file_volume and file_index)
        row = self.conn.execute(
            """SELECT * FROM items
               WHERE (file_key = ? OR
                      (? = 1 AND file_volume = ? AND file_index = ?))
                 AND sheet_name = ? AND cell_address = ?
                 AND status != ?
               ORDER BY created_at DESC LIMIT 1""",
            (file_key, 1 if identity else 0, file_volume, file_index,
             sheet_name, cell_address, STATUS_DONE)).fetchone()
        return Item.from_row(row) if row else None

    def update_item(self, item_id: int, **fields) -> None:
        old_row = self.conn.execute(
            "SELECT * FROM items WHERE id = ?", (item_id,)).fetchone()
        if old_row is None:
            return
        allowed = {k: v for k, v in fields.items() if k in _ITEM_FIELDS}
        if not allowed:
            return
        old = dict(old_row)
        now = now_str()
        sets = ["%s = ?" % k for k in allowed]
        values = list(allowed.values())
        if "status" in allowed:
            new_status = allowed["status"]
            if new_status == STATUS_DONE:
                sets.append("completed_at = COALESCE(completed_at, ?)")
                values.append(now)
            else:
                sets.append("completed_at = NULL")
            if new_status == STATUS_WAITING and (
                    old.get("status") != STATUS_WAITING or
                    not old.get("waiting_since")):
                sets.append("waiting_since = COALESCE(waiting_since, ?)")
                values.append(now)
            elif new_status != STATUS_WAITING and old.get("status") == STATUS_WAITING:
                sets.append("waiting_since = NULL")
        sets.append("updated_at = ?")
        values.append(now)
        values.append(item_id)
        self.conn.execute(
            "UPDATE items SET %s WHERE id = ?" % ", ".join(sets), values)

        events = []
        for name, new_value in allowed.items():
            if name not in _EVENT_FIELDS:
                continue
            if old.get(name) != new_value:
                events.append((name, old.get(name), new_value))
        for event_type, old_value, new_value in events:
            self._insert_event(item_id, event_type, old_value, new_value, now)
        if "project" in allowed:
            self._ensure_project(str(allowed["project"] or ""), now)
        self.conn.commit()

    def set_status(self, item_id: int, status: str) -> None:
        self.update_item(item_id, status=status)

    def update_path(self, item_id: int, file_path: str,
                    file_volume: Optional[str] = None,
                    file_index: Optional[str] = None) -> None:
        if not file_volume or not file_index:
            computed_volume, computed_index = win32util.file_identity(file_path)
            file_volume = file_volume or computed_volume
            file_index = file_index or computed_index
        self.update_item(
            item_id,
            file_path=file_path,
            file_key=nativeom.path_key(file_path),
            file_name=os.path.basename(file_path),
            file_volume=file_volume,
            file_index=file_index,
        )

    def update_snapshot(self, item_id: int, snapshot: str) -> None:
        self.update_item(item_id, snapshot=snapshot)

    def delete_item(self, item_id: int) -> None:
        self.conn.execute("DELETE FROM item_events WHERE item_id = ?", (item_id,))
        self.conn.execute("DELETE FROM items WHERE id = ?", (item_id,))
        self.conn.commit()

    # --- history -----------------------------------------------------------
    def _insert_event(self, item_id: int, event_type: str,
                      old_value, new_value, created_at: str) -> None:
        self.conn.execute(
            "INSERT INTO item_events(item_id, event_type, old_value, new_value, created_at) "
            "VALUES (?, ?, ?, ?, ?)",
            (item_id, event_type,
             None if old_value is None else str(old_value),
             None if new_value is None else str(new_value), created_at))

    def log_event(self, item_id: int, event_type: str, old_value=None,
                  new_value=None) -> None:
        self._insert_event(item_id, event_type, old_value, new_value, now_str())
        self.conn.commit()

    def list_events(self, item_id: int) -> List[ItemEvent]:
        rows = self.conn.execute(
            "SELECT * FROM item_events WHERE item_id = ? ORDER BY id",
            (item_id,)).fetchall()
        return [ItemEvent.from_row(row) for row in rows]

    # --- projects ----------------------------------------------------------
    def _ensure_project(self, name: str, now: Optional[str] = None):
        name = (name or "").strip()
        if not name:
            return
        now = now or now_str()
        self.conn.execute(
            "INSERT OR IGNORE INTO project_meta(name, archived, created_at, updated_at) "
            "VALUES (?, 0, ?, ?)", (name, now, now))

    def project_infos(self) -> List[ProjectInfo]:
        names = {row[0] for row in self.conn.execute(
            "SELECT DISTINCT project FROM items WHERE TRIM(project) != ''")}
        names.update(row[0] for row in self.conn.execute(
            "SELECT name FROM project_meta WHERE TRIM(name) != ''"))
        counts = {row[0]: int(row[1]) for row in self.conn.execute(
            "SELECT project, COUNT(*) FROM items WHERE TRIM(project) != '' GROUP BY project")}
        archived = {row[0]: bool(row[1]) for row in self.conn.execute(
            "SELECT name, archived FROM project_meta")}
        return [ProjectInfo(name=name, archived=archived.get(name, False),
                            item_count=counts.get(name, 0))
                for name in sorted(names)]

    def set_project_archived(self, name: str, archived: bool) -> None:
        now = now_str()
        self._ensure_project(name, now)
        self.conn.execute(
            "UPDATE project_meta SET archived = ?, updated_at = ? WHERE name = ?",
            (1 if archived else 0, now, name))
        self.conn.commit()

    def project_archived(self, name: str) -> bool:
        row = self.conn.execute(
            "SELECT archived FROM project_meta WHERE name = ?", (name,)).fetchone()
        return bool(row[0]) if row else False

    # --- queries -----------------------------------------------------------
    def _fts_search(self, keyword: str) -> Optional[List[int]]:
        keyword = (keyword or "").strip()
        if not self.fts_enabled or len(keyword) < 3:
            return None
        query = '"' + keyword.replace('"', '""') + '"'
        try:
            rows = self.conn.execute(
                "SELECT rowid FROM items_fts WHERE items_fts MATCH ?",
                (query,)).fetchall()
        except sqlite3.OperationalError:
            log.exception("FTS query failed; falling back to LIKE")
            return None
        # An empty FTS result can be a tokenizer edge case; LIKE is a safe
        # fallback and still returns quickly for personal-sized databases.
        return [int(row[0]) for row in rows] or None

    def list_items(self, status_filter: Optional[str] = None,
                   project: Optional[str] = None,
                   keyword: Optional[str] = None,
                   aging_days: Optional[int] = None,
                   include_archived: bool = False,
                   unassigned_only: bool = False) -> List[Item]:
        clauses = []
        params: List[object] = []
        if status_filter and status_filter not in ("全部", ""):
            if status_filter == STATUS_TODO_FILTER:
                clauses.append("status != ?")
                params.append(STATUS_DONE)
            elif status_filter in STATUSES:
                clauses.append("status = ?")
                params.append(status_filter)
        if project and project not in ("全部", ""):
            clauses.append("project = ?")
            params.append(project)
        elif not include_archived:
            clauses.append(
                "project NOT IN (SELECT name FROM project_meta WHERE archived = 1)")
        if aging_days:
            clauses.append(
                "status = ? AND date(COALESCE(NULLIF(missing_requested_at, ''), "
                "NULLIF(waiting_since, ''), updated_at)) <= date('now', ?)")
            params.extend([STATUS_WAITING, "-%d day" % int(aging_days)])
        if unassigned_only:
            clauses.append("status != ? AND TRIM(missing_owner) = ''")
            params.append(STATUS_DONE)
        if keyword and keyword.strip():
            ids = self._fts_search(keyword)
            if ids is None:
                kw = "%" + keyword.strip() + "%"
                clauses.append(
                    "(program_text LIKE ? OR project LIKE ? OR file_name LIKE ? "
                    "OR sheet_name LIKE ? OR cell_address LIKE ? OR missing_info LIKE ? "
                    "OR missing_owner LIKE ? OR note LIKE ?)")
                params.extend([kw] * 8)
            elif ids:
                clauses.append("id IN (%s)" % ",".join("?" * len(ids)))
                params.extend(ids)
            else:
                clauses.append("0")
        where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
        order = ("ORDER BY CASE status WHEN '未完成' THEN 0 WHEN '等待资料' THEN 1 "
                 "WHEN '进行中' THEN 2 ELSE 3 END, created_at DESC")
        rows = self.conn.execute("SELECT * FROM items %s %s" % (where, order), params).fetchall()
        return [Item.from_row(r) for r in rows]

    def all_items(self) -> List[Item]:
        rows = self.conn.execute("SELECT * FROM items ORDER BY id").fetchall()
        return [Item.from_row(r) for r in rows]

    def distinct_projects(self, include_archived: bool = False) -> List[str]:
        if include_archived:
            names = {row[0] for row in self.conn.execute(
                "SELECT DISTINCT project FROM items WHERE TRIM(project) != ''")}
            names.update(row[0] for row in self.conn.execute(
                "SELECT name FROM project_meta WHERE TRIM(name) != ''"))
            return sorted(names)
        rows = self.conn.execute(
            """SELECT name FROM project_meta WHERE archived = 0 AND TRIM(name) != ''
               UNION
               SELECT DISTINCT project FROM items
               WHERE TRIM(project) != ''
                 AND project NOT IN (SELECT name FROM project_meta WHERE archived = 1)
               ORDER BY 1""").fetchall()
        return [row[0] for row in rows]

    def counts(self, include_archived: bool = True) -> Dict[str, int]:
        clause = ""
        if not include_archived:
            clause = (" WHERE project NOT IN "
                      "(SELECT name FROM project_meta WHERE archived = 1)")
        result = {s: 0 for s in STATUSES}
        total = 0
        for row in self.conn.execute(
                "SELECT status, COUNT(*) AS n FROM items%s GROUP BY status" % clause):
            result[row["status"]] = int(row["n"])
            total += int(row["n"])
        result["合计"] = total
        return result

    # --- export/backup -----------------------------------------------------
    def export_csv(self, dest_path: str) -> int:
        items = self.all_items()
        headers = [
            "ID", "状态", "项目", "审计程序", "底稿文件", "文件路径", "工作表",
            "单元格", "合并区域", "缺失资料", "备注", "来源", "登记时间",
            "更新时间", "完成时间", "资料负责人", "索取日期", "收到日期", "等待天数",
        ]
        with open(dest_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            for it in items:
                writer.writerow([_csv_text(value) for value in (
                    it.id, it.status, it.project, it.program_text, it.file_name,
                    it.file_path, it.sheet_name, it.cell_address,
                    it.merge_range or "", it.missing_info, it.note, it.source_name,
                    it.created_at, it.updated_at, it.completed_at or "",
                    it.missing_owner, it.missing_requested_at,
                    it.missing_received_at, it.waiting_days if it.waiting_days is not None else "",
                )])
        return len(items)

    def export_xlsx(self, dest_path: str) -> int:
        """Export a formatted Excel workbook with workbook hyperlinks."""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Alignment, Font, PatternFill
            from openpyxl.worksheet.hyperlink import Hyperlink
        except Exception as exc:
            raise RuntimeError("缺少 openpyxl，无法导出 Excel 复核清单") from exc
        items = self.all_items()
        wb = Workbook()
        ws = wb.active
        ws.title = "底稿未完成事项"
        headers = [
            "状态", "项目", "审计程序", "底稿文件", "位置", "打开文件",
            "缺失资料", "资料负责人", "索取日期", "收到日期", "等待天数",
            "备注", "登记时间", "完成时间",
        ]
        header_fill = PatternFill("solid", fgColor="1F4E78")
        header_font = Font(color="FFFFFF", bold=True)
        status_fill = {
            "未完成": PatternFill("solid", fgColor="FEF3F2"),
            "等待资料": PatternFill("solid", fgColor="FFFAEB"),
            "进行中": PatternFill("solid", fgColor="EFF8FF"),
            "已完成": PatternFill("solid", fgColor="F9FAFB"),
        }
        for col, title in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=title)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center",
                                       wrap_text=True)
        widths = {
            "A": 10, "B": 16, "C": 52, "D": 26, "E": 18, "F": 12,
            "G": 28, "H": 14, "I": 12, "J": 12, "K": 9, "L": 28,
            "M": 18, "N": 18,
        }
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
        for row, it in enumerate(items, 2):
            values = [
                it.status, _csv_text(it.project or "未分类"),
                _csv_text(it.program_text), it.file_name, it.location_text,
                "打开文件", _csv_text(it.missing_info), _csv_text(it.missing_owner),
                it.missing_requested_at, it.missing_received_at,
                it.waiting_days if it.waiting_days is not None else "",
                _csv_text(it.note), it.created_at, it.completed_at or "",
            ]
            for col, value in enumerate(values, 1):
                cell = ws.cell(row=row, column=col, value=value)
                cell.alignment = Alignment(vertical="top", wrap_text=col in (3, 7, 12))
            fill = status_fill.get(it.status)
            if fill is not None:
                ws.cell(row=row, column=1).fill = fill
            if it.file_path:
                try:
                    uri = Path(it.file_path).resolve().as_uri()
                except Exception:
                    uri = ""
                if uri:
                    link_cell = ws.cell(row=row, column=6)
                    location = "%s!%s" % (
                        it.sheet_name.replace("'", "''"),
                        display_address(it.cell_address))
                    link_cell.hyperlink = Hyperlink(
                        ref=link_cell.coordinate, target=uri, location=location)
                    link_cell.font = Font(color="0563C1", underline="single")
        ws.freeze_panes = "A2"
        if items:
            ws.auto_filter.ref = "A1:N%d" % (len(items) + 1)
        wb.save(dest_path)
        return len(items)

    def export_html(self, dest_path: str) -> int:
        items = self.all_items()
        status_class = {
            "未完成": "todo", "等待资料": "waiting",
            "进行中": "doing", "已完成": "done",
        }
        rows = []
        for it in items:
            link = ""
            if it.file_path:
                try:
                    uri = Path(it.file_path).resolve().as_uri()
                except Exception:
                    uri = ""
                if uri:
                    link = '<a href="%s">打开文件</a>' % html.escape(uri, quote=True)
            age = it.waiting_days
            cells = [
                '<span class="status %s">%s</span>' % (
                    status_class.get(it.status, ""), html.escape(it.status)),
                html.escape(_csv_text(it.project or "未分类")),
                html.escape(_csv_text(it.program_text)).replace("\n", "<br>"),
                html.escape(it.file_name),
                html.escape(it.location_text),
                link,
                html.escape(_csv_text(it.missing_info)).replace("\n", "<br>"),
                html.escape(_csv_text(it.missing_owner)),
                html.escape(it.missing_requested_at),
                html.escape(it.missing_received_at),
                "" if age is None else "%d 天" % age,
                html.escape(_csv_text(it.note)).replace("\n", "<br>"),
                html.escape(it.created_at),
                html.escape(it.completed_at or ""),
            ]
            rows.append("<tr>%s</tr>" % "".join("<td>%s</td>" % c for c in cells))
        document = """<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<title>底稿未完成事项复核清单</title>
<style>
body{font-family:"Microsoft YaHei UI",Arial,sans-serif;margin:24px;color:#101828}
h1{font-size:22px;margin:0 0 6px}
p.meta{color:#667085;font-size:12px;margin:0 0 16px}
table{border-collapse:collapse;width:100%;font-size:12px}
th,td{border:1px solid #D0D5DD;padding:6px 8px;vertical-align:top;text-align:left}
th{background:#F2F4F7;position:sticky;top:0}
tr:nth-child(even){background:#FAFAFA}
.status{display:inline-block;padding:2px 6px;border-radius:4px;white-space:nowrap}
.todo{background:#FEF3F2;color:#B42318}.waiting{background:#FFFAEB;color:#B54708}
.doing{background:#EFF8FF;color:#175CD3}.done{background:#F9FAFB;color:#475467}
a{color:#1570EF}
</style></head><body>
<h1>底稿未完成事项复核清单</h1>
<p class="meta">生成时间：__GENERATED__ · 共 __COUNT__ 条 · 本文件只包含本机数据，请按客户资料要求保管。链接只打开文件，精确位置请使用底稿未完成助手“定位底稿”。</p>
<table><thead><tr>
<th>状态</th><th>项目</th><th>审计程序</th><th>底稿文件</th><th>位置</th><th>打开</th>
<th>缺失资料</th><th>资料负责人</th><th>索取日期</th><th>收到日期</th><th>等待</th>
<th>备注</th><th>登记时间</th><th>完成时间</th>
</tr></thead><tbody>
__ROWS__
</tbody></table></body></html>"""
        document = (document
                    .replace("__GENERATED__", html.escape(now_str()))
                    .replace("__COUNT__", str(len(items)))
                    .replace("__ROWS__", "\n".join(rows)))
        with open(dest_path, "w", encoding="utf-8", newline="") as f:
            f.write(document)
        return len(items)

    def backup(self, dest_path: str) -> None:
        os.makedirs(os.path.dirname(dest_path) or ".", exist_ok=True)
        target = sqlite3.connect(dest_path)
        try:
            self.conn.backup(target)
        finally:
            target.close()
