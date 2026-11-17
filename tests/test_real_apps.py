# -*- coding: utf-8 -*-
"""真实 Excel / WPS 联动测试。

默认跳过。运行方式：
    set WPT_RUN_INTEGRATION=1
    python -m pytest tests/test_real_apps.py -v

测试会打开自己创建的临时工作簿并激活窗口，结束时只关闭该工作簿，
不会关闭或保存用户已打开的其他文档。
"""
import hashlib
import os
import subprocess
import sys
import time
import uuid

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

comtypes = pytest.importorskip("comtypes")
from openpyxl import Workbook

from workpaper_todo import capture, jump, nativeom, store, win32util
from workpaper_todo.models import Item

pytestmark = pytest.mark.skipif(
    os.environ.get("WPT_RUN_INTEGRATION") != "1",
    reason="设置 WPT_RUN_INTEGRATION=1 才运行真实软件联动测试")


@pytest.fixture(scope="session", autouse=True)
def com_apartment():
    comtypes.CoInitialize()
    yield
    try:
        comtypes.CoUninitialize()
    except Exception:
        pass


def make_workbook(path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "应收账款"
    ws["A1"] = "应收账款审计底稿（自动化测试）"
    ws["A5"] = "1. 获取应收账款明细表，复核加计正确性。" * 20
    ws.column_dimensions["A"].width = 3
    ws["A6"] = "2. 对期末余额较大的应收账款实施函证程序。"
    ws.merge_cells("A9:C9")
    ws["A9"] = "3. 评价坏账准备计提的合理性。"
    ws.merge_cells("A10:C10")
    ws["A10"] = "4. 检查坏账准备计算的准确性。"
    ws["B5"] = "D-1"
    ws2 = wb.create_sheet("存货")
    ws2["A5"] = "1. 参与存货监盘，记录盘点差异。"
    wb.save(path)


def _item_for(path: str, source_app: str) -> Item:
    return Item(file_path=path, file_key=nativeom.path_key(path),
                source_app=source_app)


def launch_and_wait(path: str, source_app: str, timeout: float = 35.0):
    exe = jump._resolve_exe(source_app)
    if not exe:
        pytest.skip("%s 未安装，跳过" % source_app)
    item = _item_for(path, source_app)
    subprocess.Popen([exe, path], cwd=os.path.dirname(path) or None)
    deadline = time.time() + timeout
    while time.time() < deadline:
        found = jump.find_open_workbook(item)
        if found:
            return found
        time.sleep(0.5)
    pytest.fail("%s 已安装但未能在 %.0f 秒内打开测试工作簿（可能有弹窗/占用）" % (source_app, timeout))


def activate_cell(sw, sheet: str, address: str):
    sw.workbook.Activate()
    ws = capture.get_worksheet(sw.workbook, sheet)
    ws.Activate()
    ws.Range(address).Select()


def selection_text(sw) -> str:
    return capture.normalize_address(str(sw.app.Selection.Address()))


def workbook_selection_text(workbook) -> str:
    """Selection of this workbook's own window.

    Application.Selection is shared by every workbook in an Excel/WPS
    instance, so another open workbook can make the global selection point at
    a different sheet/cell. Window.RangeSelection stays scoped to the workbook
    window that jump_to_item actually selected.
    """
    return capture.normalize_address(
        str(workbook.Windows(1).RangeSelection.Address()))


def find_window_for_workbook(workbook):
    target = nativeom.path_key(nativeom.workbook_path(workbook))
    for sw in nativeom.iter_sheet_windows():
        try:
            if nativeom.path_key(nativeom.workbook_path(sw.workbook)) == target:
                return sw
        except Exception:
            continue
    return None


def close_workbook(workbook):
    try:
        workbook.Close(False)
    except Exception:
        pass


@pytest.mark.parametrize("source_app", ["excel", "wps"])
def test_real_missing_file_then_relink(source_app, tmp_path):
    """A moved/renamed workbook must not be silently used; relinking must work."""
    path = str(tmp_path / ("wpt_move_%s_%s.xlsx" %
                           (source_app, uuid.uuid4().hex[:8])))
    moved_path = str(tmp_path / ("wpt_moved_%s_%s.xlsx" %
                                (source_app, uuid.uuid4().hex[:8])))
    make_workbook(path)
    sw, wb = launch_and_wait(path, source_app)
    specific = find_window_for_workbook(wb)
    if specific is not None:
        sw = specific
    st = store.Store(str(tmp_path / ("store_move_%s.db" % source_app)))
    moved_wb = None
    try:
        activate_cell(sw, "应收账款", "A5")
        info = capture.read_selection(sw.xlmain)
        iid = st.add_item(info, "自动化测试")
        item = st.get_item(iid)
        assert nativeom.path_key(item.file_path) == nativeom.path_key(path)
        before_identity = win32util.file_identity(path)
        assert before_identity[0] and before_identity[1]

        close_workbook(wb)
        wb = None
        deadline = time.time() + 8
        while time.time() < deadline and jump.find_open_workbook(item) is not None:
            time.sleep(0.25)
        assert jump.find_open_workbook(item) is None

        deadline = time.time() + 8
        while True:
            try:
                os.replace(path, moved_path)
                break
            except OSError:
                if time.time() >= deadline:
                    raise
                time.sleep(0.3)
        assert win32util.file_identity(moved_path) == before_identity

        # Old path is gone: jump must fail explicitly, not open some other file.
        with pytest.raises(jump.FileMissing):
            jump.jump_to_item(item)

        # The real app's relink path updates the stored path only after a
        # successful jump; simulate that persisted state and open the new file.
        st.update_path(iid, moved_path)
        relinked = st.get_item(iid)
        assert (relinked.file_volume, relinked.file_index) == before_identity
        outcome = jump.jump_to_item(relinked)
        assert nativeom.path_key(outcome.file_path) == nativeom.path_key(moved_path)
        found = jump.find_open_workbook(relinked)
        assert found is not None
        moved_sw, moved_wb = found
        assert nativeom.path_key(nativeom.workbook_path(moved_wb)) == \
            nativeom.path_key(moved_path)
        assert workbook_selection_text(moved_wb) == "A5"
    finally:
        if moved_wb is not None:
            close_workbook(moved_wb)
        if wb is not None:
            close_workbook(wb)
        found = jump.find_open_workbook(_item_for(moved_path, source_app))
        if found:
            close_workbook(found[1])
        st.close()
        for candidate in (path, moved_path):
            deadline = time.time() + 5
            while os.path.exists(candidate) and time.time() < deadline:
                try:
                    os.remove(candidate)
                    break
                except OSError:
                    time.sleep(0.3)


@pytest.mark.parametrize("source_app", ["excel", "wps"])
def test_real_flow(source_app, tmp_path):
    path = str(tmp_path / ("wpt_%s_%s.xlsx" % (source_app, uuid.uuid4().hex[:8])))
    make_workbook(path)
    sw, wb = launch_and_wait(path, source_app)
    specific = find_window_for_workbook(wb)
    if specific is not None:
        sw = specific
    st = store.Store(str(tmp_path / ("store_%s.db" % source_app)))
    other_wb = None
    try:
        # 1. normal cell capture
        activate_cell(sw, "应收账款", "A5")
        info = capture.read_selection(sw.xlmain)
        assert info.file_name == os.path.basename(path)
        assert info.cell_address == "A5"
        assert info.merge_range is None
        assert info.snapshot == "1. 获取应收账款明细表，复核加计正确性。" * 20
        assert info.source_app == source_app
        assert info.file_volume and info.file_index

        # 2. merged cell capture
        activate_cell(sw, "应收账款", "A9")
        merged = capture.read_selection(sw.xlmain)
        assert merged.cell_address == "A9"
        assert merged.merge_range == "A9:C9"
        assert merged.snapshot.startswith("3. 评价坏账准备")

        # 3. multi-cell selection rejected
        activate_cell(sw, "应收账款", "A5:B6")
        with pytest.raises(capture.CaptureError):
            capture.read_selection(sw.xlmain)

        # 多个连续合并区域不是一个程序单元格，不应只登记第一个。
        activate_cell(sw, "应收账款", "A9:C10")
        with pytest.raises(capture.CaptureError):
            capture.read_selection(sw.xlmain)

        # 4. store + duplicate detection
        iid = st.add_item(info, "自动化测试")
        assert st.find_active_duplicate(
            info.file_key, info.sheet_name, info.cell_address).id == iid
        item = st.get_item(iid)

        # 一个软件中同时打开另一份底稿：登记和跳转都必须匹配文件。
        other_path = str(tmp_path / ("other_%s_%s.xlsx" %
                                     (source_app, uuid.uuid4().hex[:8])))
        make_workbook(other_path)
        other_sw, other_wb = launch_and_wait(other_path, source_app)
        other_specific = find_window_for_workbook(other_wb)
        if other_specific is not None:
            other_sw = other_specific
        activate_cell(other_sw, "存货", "A5")
        other_info = capture.read_selection(other_sw.xlmain)
        assert nativeom.path_key(other_info.file_path) == nativeom.path_key(other_path)
        assert other_info.sheet_name == "存货"
        assert other_info.snapshot == "1. 参与存货监盘，记录盘点差异。"

        # 5. locate the first workbook while the second is active
        jump.jump_to_item(item)
        assert nativeom.path_key(nativeom.workbook_path(wb.Application.ActiveWorkbook)) == \
            nativeom.path_key(path)
        assert selection_text(sw) == "A5"

        # 6. content mismatch + forced jump
        bad = Item(**{**item.__dict__})
        bad.snapshot = "登记时不存在的文字"
        with pytest.raises(jump.ContentMismatch):
            jump.jump_to_item(bad)
        jump.jump_to_item(bad, allow_mismatch=True)
        assert selection_text(sw) == "A5"

        # 7. unsaved changes are neither saved to disk nor overwritten
        with open(path, "rb") as f:
            before_hash = hashlib.sha256(f.read()).digest()
        ws = capture.get_worksheet(wb, "应收账款")
        ws.Range("A20").Value[:] = "临时修改-不应被保存"
        activate_cell(sw, "应收账款", "A5")
        jump.jump_to_item(item)
        assert bool(wb.Saved) is False
        assert ws.Range("A20").Value[:] == "临时修改-不应被保存"
        with open(path, "rb") as f:
            assert hashlib.sha256(f.read()).digest() == before_hash
        ws.Range("A20").Value[:] = None

        # 8. unsaved new workbook is rejected
        new_wb = sw.app.Workbooks.Add()
        try:
            new_wb.ActiveSheet.Range("A1").Value[:] = "未保存"
            new_wb.Activate()
            new_win = find_window_for_workbook(new_wb)
            assert new_win is not None
            with pytest.raises(capture.CaptureError):
                capture.read_selection(new_win.xlmain)
        finally:
            close_workbook(new_wb)

        # 9. A closed test workbook is opened again and correctly located.
        close_workbook(wb)
        assert jump.find_open_workbook(item) is None
        wb = None
        reopened = jump.jump_to_item(item)
        assert reopened.cell_address == "A5"
        found_reopened = jump.find_open_workbook(item)
        assert found_reopened is not None
        sw, wb = found_reopened
        assert nativeom.path_key(nativeom.workbook_path(wb)) == nativeom.path_key(path)
        assert workbook_selection_text(wb) == "A5"

        # 10. persistence across a store reopen
        st.close()
        st2 = store.Store(str(tmp_path / ("store_%s.db" % source_app)))
        try:
            assert st2.get_item(iid).snapshot == info.snapshot
        finally:
            st2.close()
        st = None
    finally:
        if st is not None:
            st.close()
        if other_wb is not None:
            close_workbook(other_wb)
        if wb is None:
            found = jump.find_open_workbook(_item_for(path, source_app))
            wb = found[1] if found else None
        if wb is not None:
            close_workbook(wb)
        deadline = time.time() + 5
        while time.time() < deadline:
            try:
                os.remove(path)
                break
            except OSError:
                time.sleep(0.3)
