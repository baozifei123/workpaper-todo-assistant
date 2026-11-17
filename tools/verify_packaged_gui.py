# -*- coding: utf-8 -*-
"""Manually exercise the packaged GUI against generated Excel/WPS files.

Run only in an interactive Windows desktop.  This script creates its own
workbooks and database in a fresh temporary directory, checks the foreground
window before sending any input, and closes only its own workbooks/application.
"""
from __future__ import annotations

import ctypes
import json
import os
from pathlib import Path
import sqlite3
import subprocess
import sys
import tempfile
import time

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import comtypes
from openpyxl import Workbook
from PIL import Image
import win32con
import win32gui
import win32ui

from workpaper_todo import __window_title__, capture, jump, nativeom, win32util
from workpaper_todo.models import Item

EXE = ROOT / "dist" / "底稿未完成助手" / "底稿未完成助手.exe"
SCREENSHOT = ROOT / "testdata" / "gui_window_verification.png"
USER32 = ctypes.windll.user32


def wait_for(label, fn, seconds=25):
    deadline = time.monotonic() + seconds
    last_error = None
    while time.monotonic() < deadline:
        try:
            result = fn()
            if result:
                return result
        except Exception as exc:
            last_error = exc
        time.sleep(0.25)
    raise RuntimeError(f"{label} timed out; last error: {last_error!r}")


def focus(hwnd):
    """Focus only a known test app/workbook, never send keys on failure."""
    win32util.bring_to_front(hwnd)
    if win32util.get_foreground_window() != hwnd:
        USER32.SwitchToThisWindow(hwnd, True)
    if win32util.get_foreground_window() != hwnd:
        flags = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, flags)
        try:
            USER32.SetForegroundWindow(hwnd)
            if win32util.get_foreground_window() != hwnd:
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                x, y = (left + right) // 2, top + min(10, max(1, (bottom - top) // 10))
                hit = win32gui.WindowFromPoint((x, y))
                if win32gui.GetAncestor(hit, win32con.GA_ROOT) == hwnd:
                    USER32.SetCursorPos(x, y)
                    USER32.mouse_event(0x0002, 0, 0, 0, 0)
                    USER32.mouse_event(0x0004, 0, 0, 0, 0)
                    time.sleep(0.15)
        finally:
            win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                  flags | win32con.SWP_NOACTIVATE)
    deadline = time.monotonic() + 3
    while time.monotonic() < deadline:
        if win32util.get_foreground_window() == hwnd:
            return
        time.sleep(0.1)
    raise RuntimeError("Could not focus the intended test window; no input sent")


def keys(hwnd, *virtual_keys):
    if win32util.get_foreground_window() != hwnd:
        raise RuntimeError("Test window lost focus; no input sent")
    for key in virtual_keys:
        USER32.keybd_event(key, 0, 0, 0)
        time.sleep(0.04)
    for key in reversed(virtual_keys):
        USER32.keybd_event(key, 0, 2, 0)
        time.sleep(0.04)


def read_rows(db_path):
    if not db_path.is_file():
        return []
    with sqlite3.connect(str(db_path), timeout=1) as con:
        con.row_factory = sqlite3.Row
        return [dict(r) for r in con.execute(
            "SELECT id, file_path, sheet_name, cell_address, snapshot, status, source_app "
            "FROM items ORDER BY id")]


def screenshot_window(hwnd, path):
    """Use PrintWindow so no unrelated desktop windows are captured."""
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width, height = right - left, bottom - top
    if width <= 0 or height <= 0:
        raise RuntimeError("Test application window has no drawable area")
    src_handle = win32gui.GetWindowDC(hwnd)
    src = win32ui.CreateDCFromHandle(src_handle)
    mem = src.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(src, width, height)
    mem.SelectObject(bitmap)
    try:
        if not USER32.PrintWindow(hwnd, mem.GetSafeHdc(), 2):
            raise RuntimeError("PrintWindow did not capture the application")
        raw = bitmap.GetBitmapBits(True)
        image = Image.frombuffer("RGB", (width, height), raw, "raw", "BGRX", 0, 1)
        image.save(path)
        return image
    finally:
        win32gui.DeleteObject(bitmap.GetHandle())
        mem.DeleteDC()
        src.DeleteDC()
        win32gui.ReleaseDC(hwnd, src_handle)


def click_button(hwnd, image, color):
    """Find a uniquely coloured button in our own window and click its centre."""
    pixels = image.load()
    w, h = image.size
    matching = []
    for y in range(0, min(h, 320), 2):
        for x in range(0, w, 2):
            r, g, b = pixels[x, y][:3]
            if all(abs(actual - expected) < 5 for actual, expected in
                   zip((r, g, b), color)):
                matching.append((x, y))
    if len(matching) < 100:
        raise RuntimeError("Could not identify the button in the test window")
    # Ignore isolated same-colour pixels in the title-bar icon.  The button
    # forms a wide horizontal band; use the densest row and its neighbours.
    counts = {}
    for _, y in matching:
        counts[y] = counts.get(y, 0) + 1
    peak_y = max(counts, key=counts.get)
    matching = [(x, y) for x, y in matching if abs(y - peak_y) <= 12]
    left, top, _, _ = win32gui.GetWindowRect(hwnd)
    x = left + (min(p[0] for p in matching) + max(p[0] for p in matching)) // 2
    y = top + peak_y
    focus(hwnd)
    if win32gui.GetAncestor(win32gui.WindowFromPoint((x, y)), win32con.GA_ROOT) != hwnd:
        raise RuntimeError("The test button is obstructed; no click sent")
    USER32.SetCursorPos(x, y)
    time.sleep(0.1)
    USER32.mouse_event(0x0002, 0, 0, 0, 0)
    USER32.mouse_event(0x0004, 0, 0, 0, 0)


def build_workbook(path, source):
    wb = Workbook()
    ws = wb.active
    ws.title = "应收账款"
    ws["A1"] = "自动化测试底稿"
    ws["A5"] = f"{source} 自动化测试：检查期后回款并核对证据。"
    ws["A6"] = f"{source} 自动化测试：点击按钮登记。"
    wb.save(path)


def app_window(proc):
    hwnd = int(win32util.user32.FindWindowW(None, __window_title__) or 0)
    return hwnd if hwnd and win32util.get_window_pid(hwnd) == proc.pid else 0


def owned_workbook(path, source):
    item = Item(file_path=str(path), source_app=source)
    found = jump.find_open_workbook(item)
    if found and found[0].source_app == source:
        return found
    return None


def close_test_workbook(path, source):
    found = owned_workbook(path, source)
    if found:
        _, wb = found
        if nativeom.path_key(nativeom.workbook_path(wb)) != nativeom.path_key(str(path)):
            raise RuntimeError("Refusing to close a workbook outside the test")
        wb.Close(False)


def main():
    if not EXE.is_file():
        raise RuntimeError("The packaged executable has not been built")
    if win32util.user32.FindWindowW(None, __window_title__):
        raise RuntimeError("Another app instance is already running; no changes made")
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    comtypes.CoInitialize()
    tmp = Path(tempfile.mkdtemp(prefix="wpt_gui_acceptance_"))
    db_path = tmp / "data" / "data.db"
    data_dir = db_path.parent
    data_dir.mkdir()
    # Automated acceptance sends WM_CLOSE; disable close-to-tray so the test
    # process exits deterministically. Tray behaviour is exercised separately.
    (data_dir / "config.json").write_text(
        json.dumps({"close_to_tray": False, "onboarding_completed": True},
                   ensure_ascii=False),
        encoding="utf-8")
    env = dict(os.environ, WORKPAPER_TODO_DATA_DIR=str(data_dir))
    proc = subprocess.Popen([str(EXE)], cwd=str(EXE.parent), env=env)
    opened = []
    try:
        hwnd = wait_for("application window", lambda: app_window(proc))
        print("Packaged application started; isolated test database:", db_path, flush=True)
        for source in ("excel", "wps"):
            path = tmp / f"synthetic_{source}.xlsx"
            build_workbook(path, source)
            exe = jump._resolve_exe(source)
            if not exe:
                raise RuntimeError(f"{source} is unavailable; not counted as passed")
            subprocess.Popen([exe, str(path)], cwd=str(tmp))
            sw, wb = wait_for(f"{source} workbook", lambda: owned_workbook(path, source), 40)
            opened.append((path, source))
            wb.Activate()
            sheet = capture.get_worksheet(wb, "应收账款")
            sheet.Activate()
            sheet.Range("A5").Select()
            # find_open_workbook may return another window from the same
            # Excel instance; resolve the window of this exact workbook.
            target_hwnd = wait_for(
                f"{source} workbook window", lambda: jump.find_xlmain_for_workbook(wb))
            focus(target_hwnd)
            if (nativeom.path_key(nativeom.workbook_path(wb.Application.ActiveWorkbook)) !=
                    nativeom.path_key(str(path)) or
                    capture.normalize_address(str(wb.Application.Selection.Address())) != "A5"):
                raise RuntimeError("Test workbook or selection changed; no hotkey sent")
            keys(target_hwnd, 0x11, 0x12, 0x4D)  # Ctrl+Alt+M
            row = wait_for(f"{source} hotkey capture", lambda: next(
                (r for r in read_rows(db_path) if r["source_app"] == source), None), 20)
            expected = f"{source} 自动化测试：检查期后回款并核对证据。"
            assert nativeom.path_key(row["file_path"]) == nativeom.path_key(str(path))
            assert (row["sheet_name"], row["cell_address"], row["snapshot"], row["status"]) == (
                "应收账款", "A5", expected, "未完成")

            # The just-registered row is selected in the UI. Move the test
            # selection away, click the real Locate button, and verify COM state.
            sheet.Range("A1").Select()
            focus(hwnd)
            image = screenshot_window(hwnd, SCREENSHOT)
            click_button(hwnd, image, (21, 112, 239))
            wait_for(f"{source} locate", lambda: (
                nativeom.path_key(nativeom.workbook_path(wb.Application.ActiveWorkbook)) ==
                nativeom.path_key(str(path)) and
                capture.normalize_address(str(wb.Application.Selection.Address())) == "A5"), 20)
            focus(hwnd)
            keys(hwnd, 0x11, 0x10, 0x44)  # Ctrl+Shift+D
            wait_for(f"{source} completion", lambda: any(
                r["id"] == row["id"] and r["status"] == "已完成"
                for r in read_rows(db_path)), 10)
            print(f"PASS {source}: hotkey -> list -> Locate click -> completed", flush=True)

            if source == "wps":
                # Also verify the optional GUI capture button. The app tracks
                # the most recently focused spreadsheet window every 400 ms.
                sheet.Range("A6").Select()
                focus(target_hwnd)
                time.sleep(0.9)
                focus(hwnd)
                image = screenshot_window(hwnd, SCREENSHOT)
                click_button(hwnd, image, (18, 183, 106))
                try:
                    button_row = wait_for("WPS button capture", lambda: next(
                        (r for r in read_rows(db_path) if r["source_app"] == "wps"
                         and r["cell_address"] == "A6"), None), 20)
                except RuntimeError:
                    print("Capture-button diagnostics:", flush=True)
                    def describe(h, _):
                        if win32util.get_window_pid(h) != proc.pid or not win32gui.IsWindowVisible(h):
                            return
                        print("  window:", win32gui.GetWindowText(h), flush=True)
                        if win32gui.GetClassName(h) == "#32770":
                            def child(ch, _):
                                text = win32gui.GetWindowText(ch)
                                if text:
                                    print("    message:", text, flush=True)
                            win32gui.EnumChildWindows(h, child, None)
                    win32gui.EnumWindows(describe, None)
                    screenshot_window(hwnd, ROOT / "testdata" / "gui_button_error.png")
                    raise
                assert nativeom.path_key(button_row["file_path"]) == nativeom.path_key(str(path))
                assert button_row["snapshot"] == "wps 自动化测试：点击按钮登记。"
                focus(hwnd)
                keys(hwnd, 0x11, 0x10, 0x44)
                wait_for("WPS button-item completion", lambda: any(
                    r["id"] == button_row["id"] and r["status"] == "已完成"
                    for r in read_rows(db_path)), 10)
                print("PASS wps: capture by GUI button", flush=True)
            close_test_workbook(path, source)
            opened.remove((path, source))

        USER32.PostMessageW(hwnd, 0x0010, 0, 0)  # WM_CLOSE; only our app
        proc.wait(timeout=10)
        proc = subprocess.Popen([str(EXE)], cwd=str(EXE.parent), env=env)
        hwnd = wait_for("application restart", lambda: app_window(proc))
        rows = read_rows(db_path)
        assert len(rows) == 3 and all(r["status"] == "已完成" for r in rows)
        print("PASS restart: all three completed test records persisted", flush=True)
        print("Application-only screenshot:", SCREENSHOT, flush=True)
        print("Temporary test directory:", tmp, flush=True)
    finally:
        for path, source in opened:
            try:
                close_test_workbook(path, source)
            except Exception as exc:
                print("Could not close a generated test workbook:", exc, flush=True)
        if proc.poll() is None:
            current = app_window(proc)
            if current:
                USER32.PostMessageW(current, 0x0010, 0, 0)
            try:
                proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                proc.terminate()  # only the process started by this script
                proc.wait(timeout=5)


if __name__ == "__main__":
    main()
