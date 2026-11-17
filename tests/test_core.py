# -*- coding: utf-8 -*-
"""纯逻辑单元测试（不依赖 Excel/WPS）。"""
import csv
import datetime as _dt
import os
import queue
import sqlite3
import sys
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from workpaper_todo import (
    app as app_module, capture, config, jump, nativeom, secure, store, win32util,
)
from workpaper_todo.models import (
    Item, SelectionInfo, STATUS_DONE, STATUS_TODO, STATUS_WAITING,
    STATUS_DOING, display_address,
)
from workpaper_todo.win32util import parse_hotkey


def make_info(path=r"E:\底稿\应收账款.xlsx", sheet="应收账款",
              cell="A5", snapshot="检查期后回款情况", merge=None):
    return SelectionInfo(
        file_path=path,
        file_key=nativeom.path_key(path),
        file_name=os.path.basename(path),
        sheet_name=sheet,
        cell_address=cell,
        cell_address_abs="$" + cell,
        merge_range=merge,
        snapshot=snapshot,
        source_app="excel",
        source_exe="",
    )


# --- address / text --------------------------------------------------------
@pytest.mark.parametrize("raw,expected", [
    ("$A$5", "A5"),
    ("A5", "A5"),
    ("$A$9:$C$9", "A9:C9"),
    ("Sheet1!$B$2", "B2"),
    (" $d$15 ", "D15"),
])
def test_normalize_address(raw, expected):
    assert capture.normalize_address(raw) == expected


def test_normalize_address_invalid():
    with pytest.raises(capture.CaptureError):
        capture.normalize_address("not-a-cell")


def test_anchor_address():
    assert capture.anchor_address("A9:C9") == "A9"
    assert capture.anchor_address("A5") == "A5"


@pytest.mark.parametrize("value,expected", [
    (" hello ", " hello "),
    (None, ""),
    (123, "123"),
    ((("第一行", "", ""),), "第一行"),
    (("", "第二行"), "第二行"),
    ((), ""),
])
def test_flatten_text(value, expected):
    assert capture._flatten_text(value) == expected


def test_same_text_normalizes_line_endings_only():
    assert jump.same_text("a\r\nb", "a\nb")
    assert not jump.same_text("a  b\n c", "a b c")
    assert not jump.same_text("a b", "a c")


def test_display_address():
    assert display_address("$A$9:$C$9") == "A9:C9"


# --- path / hotkey ---------------------------------------------------------
def test_path_key_case_insensitive():
    assert nativeom.path_key(r"E:\Data\A.xlsx") == nativeom.path_key(r"e:\data\a.xlsx")
    assert nativeom.path_key("") == ""


@pytest.mark.parametrize("text,canonical", [
    ("Ctrl+Alt+M", "Ctrl+Alt+M"),
    ("ctrl+alt+m", "Ctrl+Alt+M"),
    ("Ctrl+Shift+F9", "Ctrl+Shift+F9"),
    ("Alt+Shift+J", "Alt+Shift+J"),
])
def test_parse_hotkey(text, canonical):
    spec = parse_hotkey(text)
    assert spec.text == canonical


@pytest.mark.parametrize("text", ["M", "Ctrl+Alt+", "Ctrl+Alt+?"])
def test_parse_hotkey_invalid(text):
    with pytest.raises(ValueError):
        parse_hotkey(text)


# --- store -----------------------------------------------------------------
@pytest.fixture()
def temp_store(tmp_path):
    st = store.Store(str(tmp_path / "data.db"))
    yield st
    st.close()


def test_store_add_duplicate_status_counts(temp_store):
    info = make_info()
    iid = temp_store.add_item(info, "项目A")
    item = temp_store.get_item(iid)
    assert item.status == STATUS_TODO
    assert item.snapshot == info.snapshot
    assert item.file_key == info.file_key
    assert temp_store.find_active_duplicate(
        info.file_key, info.sheet_name, info.cell_address).id == iid

    temp_store.set_status(iid, STATUS_DONE)
    assert temp_store.find_active_duplicate(
        info.file_key, info.sheet_name, info.cell_address) is None
    counts = temp_store.counts()
    assert counts[STATUS_DONE] == 1 and counts["合计"] == 1


def test_store_filters_and_search(temp_store):
    a = temp_store.add_item(make_info(snapshot="检查期后回款"), "项目A")
    b = temp_store.add_item(make_info(
        path=r"E:\底稿\存货.xlsx", sheet="存货", cell="B2",
        snapshot="存货监盘"), "项目B")
    temp_store.set_status(b, STATUS_WAITING)

    assert {i.id for i in temp_store.list_items("全部")} == {a, b}
    assert [i.id for i in temp_store.list_items(STATUS_WAITING)] == [b]
    # 待办按状态优先级排序：未完成在前，等待资料在后
    assert [i.id for i in temp_store.list_items("待办")] == [a, b]
    assert {i.id for i in temp_store.list_items(project="项目A")} == {a}
    assert {i.id for i in temp_store.list_items(keyword="监盘")} == {b}
    assert temp_store.distinct_projects() == ["项目A", "项目B"]


def test_store_update_path_and_snapshot(temp_store):
    iid = temp_store.add_item(make_info(), "项目A")
    new_path = r"E:\新目录\改名后.xlsx"
    temp_store.update_path(iid, new_path)
    item = temp_store.get_item(iid)
    assert item.file_path == new_path
    assert item.file_name == "改名后.xlsx"
    assert item.file_key == nativeom.path_key(new_path)

    temp_store.update_snapshot(iid, "新的内容")
    assert temp_store.get_item(iid).snapshot == "新的内容"


def test_store_delete(temp_store):
    iid = temp_store.add_item(make_info(), "项目A")
    temp_store.delete_item(iid)
    assert temp_store.get_item(iid) is None
    assert temp_store.all_items() == []


def test_store_export_csv_and_backup(temp_store, tmp_path):
    temp_store.add_item(make_info(), "项目A")
    csv_path = tmp_path / "out.csv"
    assert temp_store.export_csv(str(csv_path)) == 1
    raw = csv_path.read_bytes()
    assert raw.startswith(b"\xef\xbb\xbf")  # UTF-8 BOM for Excel
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    assert rows[0][0] == "ID"
    assert rows[1][1] == STATUS_TODO

    backup_path = tmp_path / "backup.db"
    temp_store.backup(str(backup_path))
    conn = sqlite3.connect(str(backup_path))
    try:
        assert conn.execute("select count(*) from items").fetchone()[0] == 1
    finally:
        conn.close()


def test_store_persistence(tmp_path):
    path = str(tmp_path / "persist.db")
    st = store.Store(path)
    iid = st.add_item(make_info(), "项目A")
    st.close()
    st2 = store.Store(path)
    try:
        assert st2.get_item(iid).snapshot == "检查期后回款情况"
    finally:
        st2.close()


def test_store_records_status_and_edit_history(temp_store):
    iid = temp_store.add_item(make_info(), "项目A")
    temp_store.set_status(iid, STATUS_WAITING)
    temp_store.update_item(iid, project="项目B", note="等待银行回函")
    events = temp_store.list_events(iid)
    assert [e.event_type for e in events] == ["created", "status", "project", "note"]
    assert events[1].old_value == STATUS_TODO
    assert events[1].new_value == STATUS_WAITING
    item = temp_store.get_item(iid)
    assert item.waiting_since
    assert item.waiting_days == 0


def test_store_fts_searches_chinese_and_tracks_updates(temp_store):
    iid = temp_store.add_item(make_info(snapshot="检查期后回款并核对银行回函"), "项目A")
    assert [i.id for i in temp_store.list_items(keyword="期后回款")] == [iid]
    assert [i.id for i in temp_store.list_items(keyword="回款")] == [iid]
    temp_store.update_item(iid, program_text="复核存货监盘差异")
    assert temp_store.list_items(keyword="期后回款") == []
    assert [i.id for i in temp_store.list_items(keyword="存货监盘")] == [iid]


def test_store_project_archive_filters(temp_store):
    a = temp_store.add_item(make_info(), "项目A")
    b = temp_store.add_item(make_info(
        path=r"E:\底稿\存货.xlsx", sheet="存货", cell="B2",
        snapshot="存货监盘"), "项目B")
    temp_store.set_project_archived("项目A", True)
    assert temp_store.distinct_projects() == ["项目B"]
    assert {i.id for i in temp_store.list_items()} == {b}
    assert {i.id for i in temp_store.list_items(project="项目A")} == {a}
    assert temp_store.counts(include_archived=False)["合计"] == 1
    infos = {p.name: p for p in temp_store.project_infos()}
    assert infos["项目A"].archived and infos["项目A"].item_count == 1


def test_store_waiting_aging_filter(temp_store):
    iid = temp_store.add_item(make_info(), "项目A")
    temp_store.set_status(iid, STATUS_WAITING)
    old_date = (_dt.date.today() - _dt.timedelta(days=10)).isoformat()
    temp_store.update_item(iid, missing_requested_at=old_date,
                           missing_owner="客户财务部")
    assert temp_store.get_item(iid).waiting_days >= 10
    assert {i.id for i in temp_store.list_items(aging_days=7)} == {iid}
    assert temp_store.list_items(aging_days=30) == []


def test_store_duplicate_by_file_identity_after_rename(temp_store):
    info = make_info()
    info.file_volume = "ABC123"
    info.file_index = "0000000100000002"
    iid = temp_store.add_item(info, "项目A")
    found = temp_store.find_active_duplicate(
        nativeom.path_key(r"E:\底稿\改名后.xlsx"), info.sheet_name,
        info.cell_address, info.file_volume, info.file_index)
    assert found and found.id == iid


def test_store_html_export_escapes_content_and_links(temp_store, tmp_path):
    info = make_info(snapshot="<script>alert(1)</script>", sheet="A&B")
    temp_store.add_item(info, "项目<A>")
    path = tmp_path / "review.html"
    assert temp_store.export_html(str(path)) == 1
    text = path.read_text(encoding="utf-8")
    assert "&lt;script&gt;alert(1)&lt;/script&gt;" in text
    assert "项目&lt;A&gt;" in text
    assert "file:///" in text
    assert "A&amp;B" in text


def test_store_xlsx_export_has_hyperlink(temp_store, tmp_path):
    from openpyxl import load_workbook
    info = make_info(snapshot="=1+1", sheet="应收账款")
    temp_store.add_item(info, "项目A")
    path = tmp_path / "review.xlsx"
    assert temp_store.export_xlsx(str(path)) == 1
    wb = load_workbook(path)
    try:
        ws = wb["底稿未完成事项"]
        assert ws["A1"].value == "状态"
        assert ws["C2"].value.startswith("'=")
        link = ws["F2"].hyperlink
        assert link is not None
        assert link.target.startswith("file:///")
        assert "应收账款" in link.location
        assert link.location.endswith("A5")
    finally:
        wb.close()


def test_store_migrates_v1_database(tmp_path):
    path = str(tmp_path / "old.db")
    conn = sqlite3.connect(path)
    conn.executescript("""
        CREATE TABLE items (
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
            note TEXT NOT NULL DEFAULT '',
            source_app TEXT NOT NULL DEFAULT '',
            source_exe TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            completed_at TEXT
        );
        INSERT INTO items(project, program_text, status, file_path, file_key,
            file_name, sheet_name, cell_address, snapshot, created_at, updated_at)
        VALUES ('旧项目', '旧程序', '等待资料', 'E:\\old.xlsx', 'e:\\old.xlsx',
            'old.xlsx', 'Sheet1', 'A1', '旧程序', '2026-01-01 10:00:00',
            '2026-01-02 10:00:00');
        PRAGMA user_version = 1;
    """)
    conn.commit()
    conn.close()

    st = store.Store(path)
    try:
        item = st.get_item(1)
        assert item.missing_owner == ""
        assert item.file_volume == ""
        assert item.waiting_since == "2026-01-02 10:00:00"
        assert st.distinct_projects() == ["旧项目"]
        assert st.project_infos()[0].item_count == 1
        assert st.fts_enabled
        assert [i.id for i in st.list_items(keyword="旧程序")] == [1]
    finally:
        st.close()
    assert os.path.exists(path + ".v1.bak")


def test_file_identity_stable_after_rename(tmp_path):
    old = tmp_path / "old.xlsx"
    new = tmp_path / "new.xlsx"
    old.write_bytes(b"synthetic")
    first = win32util.file_identity(str(old))
    assert first[0] and first[1]
    os.replace(old, new)
    assert win32util.file_identity(str(new)) == first


# --- config ----------------------------------------------------------------
def test_config_roundtrip(tmp_path):
    cfg = config.Config(str(tmp_path))
    assert cfg.get("hotkey") == config.DEFAULT_HOTKEY
    cfg.set("current_project", "项目A")
    cfg2 = config.Config(str(tmp_path))
    assert cfg2.get("current_project") == "项目A"


def test_config_corrupt(tmp_path):
    path = tmp_path / "config.json"
    path.write_text("not json", encoding="utf-8")
    cfg = config.Config(str(tmp_path))
    assert cfg.get("hotkey") == config.DEFAULT_HOTKEY
    assert (tmp_path / "config.json.bad").exists()


def test_data_dir_pointer_roundtrip(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "local"))
    target = tmp_path / "custom-data"
    target.mkdir()
    assert config.default_data_dir() == str(tmp_path / "local" / "WorkpaperTodo")
    config.set_data_dir(str(target))
    assert config.default_data_dir() == str(target)
    config.clear_data_dir()
    assert config.default_data_dir() == str(tmp_path / "local" / "WorkpaperTodo")


def test_suggest_project_from_folder():
    assert app_module.suggest_project(r"E:\项目A\底稿\应收账款.xlsx", "") == "项目A"
    assert app_module.suggest_project(r"E:\项目A\底稿\应收账款.xlsx", "项目B") == "项目B"
    assert app_module.suggest_project(r"E:\项目A\底稿\应收账款.xlsx", "", False) == "未分类"


def test_redact_absolute_paths():
    text = r"failed E:\客户\底稿\应收账款.xlsx and \\server\share\a.xlsx"
    redacted = app_module._redact_paths(text)
    assert "应收账款.xlsx" not in redacted
    assert "<path>" in redacted
    assert "failed" in redacted


def test_dpapi_secure_roundtrip(tmp_path):
    data = "审计底稿测试".encode("utf-8")
    protected = secure.protect(data)
    assert protected != data
    assert secure.unprotect(protected) == data
    src = tmp_path / "plain.db"
    enc = tmp_path / "backup.enc"
    out = tmp_path / "restored.db"
    src.write_bytes(data)
    secure.encrypt_file(str(src), str(enc))
    assert enc.read_bytes().startswith(secure.MAGIC)
    secure.decrypt_file(str(enc), str(out))
    assert out.read_bytes() == data


def test_auto_backup_creates_and_prunes(tmp_path, monkeypatch):
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    st = store.Store(str(data_dir / "data.db"))
    cfg = config.Config(str(data_dir))
    cfg.set("auto_backup_keep", 2)
    gui = app_module.App.__new__(app_module.App)
    gui.data_dir = str(data_dir)
    gui.store = st
    gui.config = cfg
    # Pretend older automatic backups exist.
    backup_dir = data_dir / "backups"
    backup_dir.mkdir()
    for name in ("auto-20260101-000000.db", "auto-20260102-000000.db",
                 "auto-20260103-000000.db"):
        (backup_dir / name).write_bytes(b"old")
    gui._auto_backup()
    files = sorted(backup_dir.glob("auto-*.db"))
    assert len(files) == 2
    assert cfg.get("last_auto_backup")
    st.close()


def test_change_data_dir_copies_database(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "local"))
    old = tmp_path / "old"
    old.mkdir()
    st = store.Store(str(old / "data.db"))
    st.add_item(make_info(), "项目A")
    cfg = config.Config(str(old))
    gui = app_module.App.__new__(app_module.App)
    gui.data_dir = str(old)
    gui.store = st
    gui.config = cfg
    gui.root = object()
    new = tmp_path / "new"
    new.mkdir()
    monkeypatch.setattr(app_module.filedialog, "askdirectory",
                        lambda **kwargs: str(new))
    monkeypatch.setattr(app_module.messagebox, "askyesno", lambda *a, **k: True)
    infos, errors = [], []
    monkeypatch.setattr(app_module.messagebox, "showinfo",
                        lambda *a, **k: infos.append(a))
    monkeypatch.setattr(app_module.messagebox, "showerror",
                        lambda *a, **k: errors.append(a))
    gui.change_data_dir()
    assert not errors
    assert (new / "data.db").is_file()
    assert (new / "config.json").is_file()
    assert config.default_data_dir() == str(new)
    copied = store.Store(str(new / "data.db"))
    try:
        assert len(copied.all_items()) == 1
    finally:
        copied.close()
    st.close()


def test_encrypted_backup_end_to_end(tmp_path, monkeypatch):
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    st = store.Store(str(data_dir / "data.db"))
    st.add_item(make_info(), "项目A")
    cfg = config.Config(str(data_dir))
    gui = app_module.App.__new__(app_module.App)
    gui.data_dir = str(data_dir)
    gui.store = st
    gui.config = cfg
    gui.root = object()
    dest = tmp_path / "backup.db.enc"
    monkeypatch.setattr(app_module.filedialog, "asksaveasfilename",
                        lambda **kwargs: str(dest))
    errors = []
    monkeypatch.setattr(app_module.messagebox, "showinfo", lambda *a, **k: None)
    monkeypatch.setattr(app_module.messagebox, "showerror",
                        lambda *a, **k: errors.append(a))
    gui.encrypted_backup()
    assert not errors
    assert dest.read_bytes().startswith(secure.MAGIC)
    restored = tmp_path / "restored.db"
    secure.decrypt_file(str(dest), str(restored))
    check = store.Store(str(restored))
    try:
        assert len(check.all_items()) == 1
    finally:
        check.close()
    st.close()


def test_snapshot_uses_full_cell_value_not_clipped_display():
    full_text = "核对完整的审计程序和资料说明" * 30
    cell = SimpleNamespace(Text="核对完整...", Value=full_text, Formula=full_text)
    sheet = SimpleNamespace(Range=lambda _addr: cell)
    workbook = SimpleNamespace(Worksheets=SimpleNamespace(Item=lambda _name: sheet))
    assert capture.read_cell_snapshot(workbook, "测试", "A5") == full_text


def test_window_parent_failure_must_not_use_active_workbook(monkeypatch):
    class Window:
        @property
        def Parent(self):
            raise RuntimeError("workbook unavailable")

        Application = SimpleNamespace(ActiveWorkbook=object())

    monkeypatch.setattr(nativeom.win32util, "find_child_by_class", lambda *_: [7])
    monkeypatch.setattr(nativeom.win32util.user32, "IsWindowVisible", lambda _: True)
    monkeypatch.setattr(nativeom, "native_window", lambda _: Window())
    with pytest.raises(nativeom.NativeOMError):
        nativeom.get_sheet_object(5)


def test_activating_unverifiable_workbook_does_not_succeed():
    class App:
        @property
        def ActiveWorkbook(self):
            raise RuntimeError("cannot read active workbook")

    wb = SimpleNamespace(Activate=lambda: None, FullName=r"E:\a.xlsx",
                         Windows=SimpleNamespace(Item=lambda _n: SimpleNamespace(
                             Activate=lambda: None)))
    with pytest.raises(jump.JumpError):
        jump._activate_workbook(wb, App())


def test_completed_timestamp_preserved_when_editing_details(temp_store, monkeypatch):
    iid = temp_store.add_item(make_info(), "项目A")
    monkeypatch.setattr(store, "now_str", lambda: "2026-01-01 10:00:00")
    temp_store.set_status(iid, STATUS_DONE)
    monkeypatch.setattr(store, "now_str", lambda: "2026-01-02 10:00:00")
    temp_store.update_item(iid, status=STATUS_DONE, note="补充说明")
    assert temp_store.get_item(iid).completed_at == "2026-01-01 10:00:00"
    temp_store.set_status(iid, STATUS_TODO)
    assert temp_store.get_item(iid).completed_at is None


@pytest.mark.parametrize("value", ["=1+1", "+1+1", "-1+1", "@SUM(A1)", "  =1+1"])
def test_csv_export_neutralizes_spreadsheet_formulas(temp_store, tmp_path, value):
    temp_store.add_item(make_info(snapshot=value), "项目A")
    path = tmp_path / "safe.csv"
    temp_store.export_csv(str(path))
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    assert rows[1][3].startswith("'")


def test_capture_rejects_selection_changed_during_read(monkeypatch):
    first = make_info(cell="A5")
    second = make_info(cell="A6")
    reads = iter([first, second])
    monkeypatch.setattr(capture.win32util, "resolve_sheet_window", lambda h: h)
    monkeypatch.setattr(capture.win32util.user32, "IsWindowVisible", lambda h: True)
    monkeypatch.setattr(capture.win32util, "get_window_pid", lambda h: 42)
    monkeypatch.setattr(capture.win32util, "get_process_path", lambda p: "EXCEL.EXE")
    monkeypatch.setattr(capture.nativeom, "get_sheet_object", lambda h: object())
    monkeypatch.setattr(capture, "_read_from_sheet_window", lambda _: next(reads))
    with pytest.raises(capture.CaptureError, match="选区发生变化"):
        capture.read_selection(10)


def test_capture_does_not_replace_missing_hotkey_window(monkeypatch):
    monkeypatch.setattr(capture.win32util, "resolve_sheet_window", lambda h: h)
    monkeypatch.setattr(capture.win32util, "get_foreground_window", lambda: 123)
    with pytest.raises(capture.CaptureError, match="未检测到"):
        capture.read_selection(0, require_foreground=True)


def test_hotkey_rejects_changed_foreground(monkeypatch):
    monkeypatch.setattr(capture.win32util, "resolve_sheet_window", lambda h: h)
    monkeypatch.setattr(capture.win32util.user32, "IsWindowVisible", lambda h: True)
    monkeypatch.setattr(capture.win32util, "get_foreground_window", lambda: 999)
    with pytest.raises(capture.CaptureError, match="窗口已切换"):
        capture.read_selection(10, require_foreground=True)


def test_snapshot_read_failure_is_not_treated_as_blank(monkeypatch):
    def fail(*_):
        raise RuntimeError("spreadsheet busy")
    monkeypatch.setattr(capture, "read_cell_snapshot", fail)
    with pytest.raises(jump.JumpError, match="无法读取"):
        jump.current_snapshot(object(), "Sheet1", "A1")


def test_duplicate_update_reuses_existing_item(temp_store, tmp_path, monkeypatch):
    existing_id = temp_store.add_item(make_info(snapshot="旧内容"), "项目A")
    gui = app_module.App.__new__(app_module.App)
    gui.store = temp_store
    gui.config = config.Config(str(tmp_path))
    gui.project_var = SimpleNamespace(get=lambda: "项目A")
    gui._ensure_visible = Mock()
    gui.refresh = Mock()
    gui.toast = Mock()
    gui.show_detail = Mock()
    monkeypatch.setattr(app_module, "DuplicateDialog",
                        lambda *args: SimpleNamespace(show=lambda: "update"))
    gui._handle_capture(make_info(snapshot="新内容"), "项目A")
    assert len(temp_store.all_items()) == 1
    assert temp_store.get_item(existing_id).snapshot == "新内容"
    gui.show_detail.assert_called_once_with(existing_id)


def test_mismatch_update_waits_for_success(monkeypatch):
    from workpaper_todo import app
    gui = app.App.__new__(app.App)
    gui.store = Mock()
    gui.locate_item = Mock()
    gui._ensure_visible = Mock()
    monkeypatch.setattr(app, "MismatchDialog", lambda *args: SimpleNamespace(
        show=lambda: "update"))
    original = Item(id=5, snapshot="原始内容")
    after_success = Mock()
    gui._handle_jump_error(original, jump.ContentMismatch("原始内容", "新内容"),
                           False, after_success)
    gui.store.update_snapshot.assert_not_called()
    candidate = gui.locate_item.call_args.args[0]
    assert original.snapshot == "原始内容"
    assert candidate.snapshot == "新内容"
    assert gui.locate_item.call_args.kwargs["snapshot_to_save"] == "新内容"
    assert gui.locate_item.call_args.kwargs["allow_mismatch"] is False


def test_relink_does_not_save_path_until_success(monkeypatch):
    from workpaper_todo import app
    gui = app.App.__new__(app.App)
    gui.store = Mock()
    gui.locate_item = Mock()
    gui._ensure_visible = Mock()
    gui.refresh = Mock()
    gui.toast = Mock()
    gui.root = object()
    target = r"E:\新位置.xlsx"

    class FakeRelinkDialog:
        def __init__(self, _app, _item):
            pass

        def show(self):
            return target

    monkeypatch.setattr(app, "RelinkDialog", FakeRelinkDialog)
    monkeypatch.setattr(app.win32util, "file_identity",
                        lambda _path: ("volume", "index"))
    monkeypatch.setattr(app.nativeom, "path_key", lambda path: path.lower())
    item = Item(id=5, file_path=r"E:\旧位置.xlsx", snapshot="记录内容")
    gui._handle_jump_error(item, jump.FileMissing("missing"), False, None)
    gui.store.update_path.assert_not_called()
    candidate = gui.locate_item.call_args.args[0]
    assert candidate.file_path == target
    assert gui.locate_item.call_args.kwargs["relink_from"] == item.file_path
    gui.locate_item.call_args.kwargs["after_success"](object())
    gui.store.update_path.assert_called_once_with(5, target, "volume", "index")
    gui.store.log_event.assert_called_once_with(
        5, "relocate", item.file_path, target)


def test_completed_shortcut_is_ignored_while_dialog_open():
    from workpaper_todo import app
    gui = app.App.__new__(app.App)
    gui._dialog_open = lambda: True
    gui.set_selected_status = Mock()
    gui._accel_done()
    gui.set_selected_status.assert_not_called()


def test_worker_shuts_down_after_processing():
    from workpaper_todo.app import Worker
    results = queue.Queue()
    worker = Worker(results)
    worker.submit(lambda: 42, "test", lambda _: None, lambda _: None)
    message = results.get(timeout=5)
    assert message[2] == 42 and message[3] is None
    worker.stop()
    assert not worker.thread.is_alive()


def test_unassigned_quick_view_includes_only_open_unowned_items(tmp_path):
    db = store.Store(str(tmp_path / "unassigned.db"))
    try:
        unassigned = db.add_item(
            make_info(path=r"E:\\practice\\open.xlsx", cell="A1"), "练习项目")
        db.update_item(unassigned, status=STATUS_WAITING,
                       missing_info="模拟银行回单", missing_owner="")
        assigned = db.add_item(
            make_info(path=r"E:\\practice\\assigned.xlsx", cell="A2"), "练习项目")
        db.update_item(assigned, missing_info="模拟明细", missing_owner="审计员A")
        completed = db.add_item(
            make_info(path=r"E:\\practice\\done.xlsx", cell="A3"), "练习项目")
        db.update_item(completed, status=STATUS_DONE, missing_owner="")

        rows = db.list_items(unassigned_only=True)
        assert [row.id for row in rows] == [unassigned]
    finally:
        db.close()


def test_practice_workbook_is_synthetic_and_never_overwrites(tmp_path):
    from openpyxl import load_workbook
    from workpaper_todo.practice import create_practice_workbook

    first = create_practice_workbook(str(tmp_path))
    second = create_practice_workbook(str(tmp_path))
    assert first != second
    assert os.path.isfile(first) and os.path.isfile(second)
    wb = load_workbook(first, read_only=True, data_only=True)
    try:
        ws = wb["练习底稿"]
        assert ws["A1"].value == "底稿未完成助手 · 练习底稿"
        assert ws["B5"].value.startswith("【练习】")
        assert "合成练习数据" in ws["A2"].value
    finally:
        wb.close()


def test_legacy_config_keeps_saved_filters_in_custom_view(tmp_path):
    cfg = config.Config(str(tmp_path))
    cfg.data.pop("quick_view", None)
    cfg.set("filter_status", STATUS_DOING)
    loaded = config.Config(str(tmp_path))
    assert loaded.get("quick_view") == "自定义"
    assert loaded.get("filter_status") == STATUS_DOING


def test_quick_view_params_and_manual_filter_reset():
    class Var:
        def __init__(self, value):
            self.value = value

        def get(self):
            return self.value

        def set(self, value):
            self.value = value

    gui = app_module.App.__new__(app_module.App)
    gui.quick_view_var = Var("待跟进")
    gui.status_filter_var = Var("待办")
    gui.aging_filter_var = Var("全部")
    gui._aging_days = lambda: None
    assert gui._effective_list_filters() == {
        "status_filter": STATUS_WAITING, "aging_days": None,
        "unassigned_only": False}
    gui.quick_view_var.set("无人负责")
    assert gui._effective_list_filters() == {
        "status_filter": "待办", "aging_days": None,
        "unassigned_only": True}
    gui.quick_view_var.set("等待>7天")
    assert gui._effective_list_filters() == {
        "status_filter": STATUS_WAITING, "aging_days": 7,
        "unassigned_only": False}
    gui.quick_view_var.set("自定义")
    gui.status_filter_var.set(STATUS_DOING)
    gui.aging_filter_var.set("等待>30天")
    gui._aging_days = lambda: 30
    assert gui._effective_list_filters() == {
        "status_filter": STATUS_DOING, "aging_days": 30,
        "unassigned_only": False}


def test_receipt_quick_action_does_not_change_status():
    from workpaper_todo.dialogs import QuickFollowUpDialog

    dialog = QuickFollowUpDialog.__new__(QuickFollowUpDialog)
    dialog.item = Item(status=STATUS_WAITING)
    dialog.received_var = Mock()
    dialog.status_var = Mock()
    dialog._mark_received_today()
    assert dialog.received_var.set.call_args.args[0] == _dt.date.today().isoformat()
    dialog.status_var.set.assert_not_called()


def test_toast_secondary_action_invokes_selected_callback():
    from workpaper_todo.toast import Toast

    toast = Toast.__new__(Toast)
    toast.close = Mock()
    calls = []
    Toast._run_action(toast, lambda: calls.append("view"))
    toast.close.assert_called_once_with()
    assert calls == ["view"]


def test_clear_filters_resets_all_controls_and_persists():
    class Var:
        def __init__(self, value):
            self.value = value

        def get(self):
            return self.value

        def set(self, value):
            self.value = value

    gui = app_module.App.__new__(app_module.App)
    gui.quick_view_var = Var("无人负责")
    gui.status_filter_var = Var(STATUS_WAITING)
    gui.project_filter_var = Var("旧项目")
    gui.aging_filter_var = Var("等待>30天")
    gui.include_archived_var = Var(True)
    gui.search_var = Var("旧关键词")
    gui.config = SimpleNamespace(data={}, save=Mock())
    gui.refresh = Mock()

    gui.clear_filters()

    assert gui.quick_view_var.get() == "待办"
    assert gui.status_filter_var.get() == "待办"
    assert gui.project_filter_var.get() == "全部"
    assert gui.aging_filter_var.get() == "全部"
    assert gui.include_archived_var.get() is False
    assert gui.search_var.get() == ""
    assert gui.config.data["quick_view"] == "待办"
    gui.config.save.assert_called_once_with()
    gui.refresh.assert_called_once_with()


def test_manual_status_filter_switches_quick_view_to_custom():
    gui = app_module.App.__new__(app_module.App)
    gui.quick_view_var = SimpleNamespace(
        get=lambda: "自定义", set=Mock())
    gui.status_filter_var = SimpleNamespace(get=lambda: STATUS_DOING)
    gui.project_filter_var = SimpleNamespace(get=lambda: "全部")
    gui.aging_filter_var = SimpleNamespace(get=lambda: "全部")
    gui.status_box = object()
    gui.aging_box = object()
    gui.config = SimpleNamespace(data={}, save=Mock())
    gui.refresh = Mock()

    gui._on_filter_change(SimpleNamespace(widget=gui.status_box))

    gui.quick_view_var.set.assert_called_once_with("自定义")
    assert gui.config.data["filter_status"] == STATUS_DOING
    gui.config.save.assert_called_once_with()
    gui.refresh.assert_called_once_with()


def test_waiting_column_sort_toggles_and_persists():
    gui = app_module.App.__new__(app_module.App)
    gui.waiting_sort = ""
    gui.config = SimpleNamespace(set=Mock())
    gui.tree = Mock()
    gui.refresh = Mock()

    gui.toggle_waiting_sort()
    assert gui.waiting_sort == "desc"
    gui.config.set.assert_called_with("sort_waiting", "desc")
    assert gui.tree.heading.call_args.kwargs["text"] == "等待 ↓"
    gui.toggle_waiting_sort()
    assert gui.waiting_sort == "asc"
    assert gui.tree.heading.call_args.kwargs["text"] == "等待 ↑"


def test_clamp_rect_keeps_window_inside_negative_coordinate_monitor():
    area = (2056, -1034, 4104, 118, 0)
    x, y, w, h = win32util.clamp_rect_to_work_area(
        (99999, 99999, 800, 600), area, 480, 360)
    assert x >= area[0] + 8
    assert y >= area[1] + 8
    assert x + w <= area[2] - 8
    assert y + h <= area[3] - 8

    x, y, w, h = win32util.clamp_rect_to_work_area(
        (area[0], area[1], 5000, 5000), area, 480, 360)
    assert w <= area[2] - area[0] - 16
    assert h <= area[3] - area[1] - 16
    assert x + w <= area[2] - 8
    assert y + h <= area[3] - 8


def test_intersecting_work_area_prefers_overlap_then_nearest():
    areas = [(0, 0, 100, 100, 1), (200, 0, 300, 100, 0)]
    assert win32util.intersecting_work_area((210, 10, 50, 50), areas) == areas[1]
    assert win32util.intersecting_work_area((999, 999, 10, 10), areas) == areas[1]
    assert win32util.intersecting_work_area((0, 0, 10, 10), []) is None
