# -*- coding: utf-8 -*-
"""读取当前（或指定）表格窗口中的选区，生成待登记信息。"""
from __future__ import annotations

import datetime as _dt
import logging
import os
import re
from typing import Optional

from . import nativeom, win32util
from .models import SelectionInfo, display_address

log = logging.getLogger(__name__)

_ADDRESS_RE = re.compile(r"^[A-Z]{1,3}[0-9]+(?::[A-Z]{1,3}[0-9]+)?$")


class CaptureError(Exception):
    """A user-facing capture failure."""

    def __init__(self, message: str, detail: str = ""):
        super().__init__(message)
        self.user_message = message
        self.detail = detail


def _safe(func, default=None):
    try:
        return func()
    except Exception:
        return default


def _safe_retry(func, default=None, attempts: int = 4):
    try:
        return nativeom.retry(func, attempts=attempts, delay=0.18)
    except Exception:
        return default


def _call_if_needed(value):
    if type(value).__name__ == "NamedProperty":
        return value()
    if callable(value):
        try:
            return value()
        except Exception:
            return value
    return value


def _flatten_text(value) -> str:
    """Flatten WPS/Excel values (str, tuple of tuples, NamedProperty) to text."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, (tuple, list)):
        for item in value:
            text = _flatten_text(item)
            if text:
                return text
        return ""
    value = _call_if_needed(value)
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    return str(value).strip()


def normalize_address(raw: str) -> str:
    text = (raw or "").strip().replace("$", "").replace(" ", "")
    if "!" in text:
        text = text.rsplit("!", 1)[-1]
    text = text.upper()
    if not _ADDRESS_RE.match(text):
        raise CaptureError("无法识别当前单元格地址：%s" % raw)
    return text


def anchor_address(address: str) -> str:
    return (address or "").split(":", 1)[0]


def get_worksheet(app, sheet_name: str):
    try:
        return nativeom.retry(
            lambda: app.Worksheets.Item(sheet_name), attempts=4, delay=0.18)
    except Exception:
        return nativeom.retry(
            lambda: app.Worksheets(sheet_name), attempts=3, delay=0.18)


def worksheet_exists(app, sheet_name: str) -> bool:
    try:
        get_worksheet(app, sheet_name)
        return True
    except Exception:
        return False


def _cell_snapshot(cell) -> str:
    """Read the complete cell value, independent of the column width.

    Excel's Text property can display a truncated string or ####.  Capture and
    verification must use the same underlying value, never the display width.
    """
    try:
        value = nativeom.retry(lambda: _call_if_needed(cell.Value),
                               attempts=4, delay=0.18)
    except Exception as exc:
        raise CaptureError("无法读取单元格内容，请稍后重试。") from exc
    return _flatten_text(value)


def read_cell_snapshot(workbook, sheet_name: str, address: str) -> str:
    """Read the full value of the recorded cell (merged cell: its anchor)."""
    ws = get_worksheet(workbook, sheet_name)
    cell = nativeom.retry(lambda: ws.Range(address), attempts=4, delay=0.18)
    return _cell_snapshot(cell)


def _selection_snapshot(sel, sheet, anchor: str) -> str:
    cell = nativeom.retry(lambda: sheet.Range(anchor), attempts=4, delay=0.18)
    return _cell_snapshot(cell)


def _read_from_sheet_window(sw: nativeom.SheetWindow) -> SelectionInfo:
    workbook = sw.workbook
    app = sw.app

    full_path = nativeom.workbook_path(workbook)
    path_prop = _safe_retry(lambda: str(workbook.Path), "")
    if not full_path:
        name = nativeom.workbook_name(workbook)
        if not name and not path_prop:
            raise CaptureError(
                "无法读取当前工作簿信息。Excel/WPS 可能正忙或处于编辑状态，"
                "请按 Esc 退出编辑后再试。")
        if not path_prop:
            raise CaptureError("当前工作簿尚未保存，请先保存文件后再登记。")
        raise CaptureError("无法读取当前工作簿的文件路径，请稍后重试。")
    if full_path.lower().startswith(("http://", "https://")):
        raise CaptureError(
            "当前文件是云文档链接，工具只支持本机或共享文件夹中的已保存文件。")
    if nativeom.is_addin(workbook):
        raise CaptureError("当前是加载项工作簿，请切换到审计底稿工作簿。")
    if os.path.basename(full_path).lower().endswith((".xla", ".xlam")):
        raise CaptureError("当前是加载项工作簿，请切换到审计底稿工作簿。")
    if not os.path.isabs(full_path):
        raise CaptureError("当前工作簿尚未保存到本地路径，请先保存文件后再登记。")
    if not path_prop:
        raise CaptureError("当前工作簿尚未保存，请先保存文件后再登记。")

    sheet = _safe_retry(lambda: workbook.ActiveSheet)
    if sheet is None:
        raise CaptureError("无法确认当前工作簿的工作表，请重新选择单元格。")
    sheet_name = str(_safe_retry(lambda: sheet.Name, "") or "")
    if not sheet_name:
        raise CaptureError("无法读取当前工作表名称。")

    # A different Excel window or WPS tab can become active while reading.
    # Never substitute Application.Selection for this window's selection.
    sel = _safe_retry(lambda: sw.win.Selection)
    if sel is None:
        raise CaptureError("无法确认当前窗口的选区，请重新选择单元格后再试。")

    try:
        selected_sheet = nativeom.retry(lambda: sel.Parent, attempts=3, delay=0.18)
        selected_book = nativeom.retry(lambda: selected_sheet.Parent,
                                       attempts=3, delay=0.18)
        selected_name = str(nativeom.retry(lambda: selected_sheet.Name,
                                          attempts=3, delay=0.18))
        selected_path = nativeom.workbook_path(selected_book)
        if (selected_name != sheet_name or not selected_path or
                nativeom.path_key(selected_path) != nativeom.path_key(full_path)):
            raise CaptureError("工作簿或工作表已切换，请重新选择单元格后再登记。")
        areas_count = int(nativeom.retry(lambda: sel.Areas.Count,
                                         attempts=3, delay=0.18))
        cells_count = int(nativeom.retry(lambda: sel.Cells.Count,
                                         attempts=3, delay=0.18))
    except CaptureError:
        raise
    except Exception as exc:
        raise CaptureError(
            "无法确认当前选区，请确保选中一个单元格并退出编辑状态。") from exc
    if areas_count != 1:
        raise CaptureError("当前选择了多个区域，请只选中一个审计程序单元格。")
    merge = bool(_safe_retry(lambda: sel.MergeCells, False, attempts=3))

    raw_address = _safe_retry(lambda: sel.Address(), attempts=3)
    if not isinstance(raw_address, str) or not raw_address:
        raw_address = _safe_retry(lambda: sel.Address, attempts=3)
    if not isinstance(raw_address, str):
        raw_address = str(raw_address)
    address = normalize_address(raw_address)
    anchor = anchor_address(address)

    if cells_count > 1:
        if not merge:
            raise CaptureError(
                "当前选择了多个单元格。请只选中一个审计程序所在的单元格后再登记。")
        try:
            anchor_cell = nativeom.retry(lambda: sheet.Range(anchor),
                                         attempts=3, delay=0.18)
            merged_address = normalize_address(nativeom.retry(
                lambda: anchor_cell.MergeArea.Address(), attempts=3, delay=0.18))
        except Exception as exc:
            raise CaptureError("无法确认合并区域，请重新选择一个合并单元格。") from exc
        if merged_address != address:
            raise CaptureError("当前选择包含多个合并单元格，请只选择其中一个。")

    snapshot = _selection_snapshot(sel, sheet, anchor)
    if not snapshot.strip():
        raise CaptureError(
            "所选单元格为空或没有可读取的文字，请选择写明审计程序的单元格。")

    merge_range = address if ":" in address else None
    if merge and merge_range is None:
        merge_range = address

    saved = bool(_safe_retry(lambda: workbook.Saved, True, attempts=3))
    readonly = bool(_safe_retry(lambda: workbook.ReadOnly, False, attempts=3))
    file_volume, file_index = win32util.file_identity(full_path)

    return SelectionInfo(
        file_path=full_path,
        file_key=nativeom.path_key(full_path),
        file_name=os.path.basename(full_path),
        sheet_name=sheet_name,
        cell_address=anchor,
        cell_address_abs=str(raw_address),
        merge_range=merge_range,
        snapshot=snapshot,
        source_app=sw.source_app,
        source_exe=sw.source_exe,
        window_hwnd=sw.xlmain,
        captured_at=_dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        file_volume=file_volume,
        file_index=file_index,
    )


def read_selection(hwnd: Optional[int] = None,
                   require_foreground: bool = False) -> SelectionInfo:
    """Read and recheck a single Excel/WPS selection; fail if it changed."""
    try:
        target = hwnd if hwnd is not None else win32util.get_foreground_window()
        xlmain = win32util.resolve_sheet_window(target)
        if not xlmain:
            raise CaptureError(
                "未检测到 Excel 或 WPS 表格窗口。请先切换到表格，选中单元格后再登记。")
        if not win32util.user32.IsWindowVisible(xlmain):
            raise CaptureError("表格窗口当前不可见，请先切换到表格窗口。")

        def check_foreground():
            if (require_foreground and
                    win32util.resolve_sheet_window(win32util.get_foreground_window()) != xlmain):
                raise CaptureError("表格窗口已切换，请重新选择单元格后再登记。")

        check_foreground()
        pid = win32util.get_window_pid(xlmain)
        exe = win32util.get_process_path(pid)
        source = nativeom._source_app(exe)
        if source not in ("excel", "wps"):
            raise CaptureError(
                "当前窗口不是 Excel 或 WPS 表格，请切换到表格后再登记。")

        info = _read_from_sheet_window(nativeom.get_sheet_object(xlmain))
        current = _read_from_sheet_window(nativeom.get_sheet_object(xlmain))
        check_foreground()
        if ((info.file_key, info.sheet_name, info.cell_address, info.merge_range,
             info.snapshot, info.source_app, info.file_volume, info.file_index) !=
                (current.file_key, current.sheet_name, current.cell_address,
                 current.merge_range, current.snapshot, current.source_app,
                 current.file_volume, current.file_index)):
            raise CaptureError("工作簿或选区发生变化，请重新选择单元格后再登记。")
        return info
    except CaptureError:
        raise
    except Exception as exc:
        log.exception("capture failed")
        raise CaptureError(
            "读取当前单元格失败，请确认 Excel/WPS 未处于编辑或弹窗状态后重试。",
            detail=repr(exc)) from exc
