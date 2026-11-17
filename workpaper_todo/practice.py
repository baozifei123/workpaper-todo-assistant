# -*- coding: utf-8 -*-
"""Generate local, synthetic workbooks for first-run practice."""
from __future__ import annotations

import os
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill


def create_practice_workbook(data_dir: str) -> str:
    """Create a new synthetic practice workbook without overwriting an old one."""
    folder = Path(data_dir) / "practice"
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / "练习底稿.xlsx"
    suffix = 2
    while target.exists():
        target = folder / ("练习底稿_%d.xlsx" % suffix)
        suffix += 1

    wb = Workbook()
    ws = wb.active
    ws.title = "练习底稿"
    ws["A1"] = "底稿未完成助手 · 练习底稿"
    ws["A1"].font = Font(name="Microsoft YaHei UI", size=15, bold=True,
                         color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="0F4C81")
    ws["A1"].alignment = Alignment(vertical="center")
    ws.merge_cells("A1:D1")
    ws.row_dimensions[1].height = 30
    ws["A2"] = "本文件为程序自动生成的合成练习数据，不含真实客户资料。"
    ws["A2"].font = Font(name="Microsoft YaHei UI", size=10, italic=True,
                         color="667085")
    ws.merge_cells("A2:D2")
    ws["A4"] = "练习步骤"
    ws["B4"] = "内容"
    for cell in ws[4]:
        cell.font = Font(name="Microsoft YaHei UI", bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1570EF")
    ws["A5"] = "1 · 选中并登记"
    ws["B5"] = "【练习】核对模拟供应商应付账款余额是否与虚构明细表一致。"
    ws["A6"] = "2 · 再选中定位"
    ws["B6"] = "【练习】检查虚构期末付款记录，并记录一项待补资料。"
    ws["A8"] = "测试提示"
    ws["B8"] = "登记练习事项后，可点击提示条上的“撤销”，不要保留为真实项目事项。"
    ws["A8"].font = Font(name="Microsoft YaHei UI", bold=True, color="B54708")
    ws["B8"].font = Font(name="Microsoft YaHei UI", color="B54708")
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 84
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.row_dimensions[5].height = 34
    ws.row_dimensions[6].height = 34
    ws.row_dimensions[8].height = 32
    for row in (5, 6, 8):
        ws.cell(row=row, column=2).alignment = Alignment(wrap_text=True,
                                                        vertical="center")
    ws.freeze_panes = "A4"
    ws.sheet_view.selection[0].activeCell = "B5"
    ws.sheet_view.selection[0].sqref = "B5"
    wb.save(str(target))
    return os.path.abspath(str(target))
