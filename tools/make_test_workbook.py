# -*- coding: utf-8 -*-
"""生成用于测试的审计底稿（不含任何真实客户资料）。"""
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

OUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "testdata")
os.makedirs(OUT_DIR, exist_ok=True)

wb = Workbook()
ws = wb.active
ws.title = "应收账款"
ws["A1"] = "应收账款审计底稿（测试用）"
ws["A1"].font = Font(bold=True, size=14)
ws["A3"] = "审计程序"
ws["B3"] = "索引号"
ws["C3"] = "执行情况"
ws["A5"] = "1. 获取应收账款明细表，复核加计正确性，并与报表数、总账数和明细账合计数核对相符。"
ws["B5"] = "D-1"
ws["A6"] = "2. 对期末余额较大的应收账款实施函证程序，函证比例不低于80%。"
ws["B6"] = "D-2"
ws["A7"] = "3. 检查期后回款情况，核对至银行流水及收款凭证。"
ws["B7"] = "D-3"
ws.merge_cells("A9:C9")
ws["A9"] = "4. 评价坏账准备计提的合理性，检查坏账准备计提政策是否一贯执行。"
ws["D9"] = "D-4"
ws.column_dimensions["A"].width = 90
ws.column_dimensions["B"].width = 12
ws.column_dimensions["C"].width = 20
for row in ws.iter_rows(min_row=5, max_row=9):
    for cell in row:
        cell.alignment = Alignment(wrap_text=True, vertical="center")

ws2 = wb.create_sheet("存货")
ws2["A1"] = "存货审计底稿（测试用）"
ws2["A3"] = "审计程序"
ws2["A5"] = "1. 参与存货监盘，记录盘点差异并追查原因。"
ws2["A6"] = "2. 对存货实施计价测试，检查跌价准备计提是否充分。"
ws2["B5"] = "E-1"
ws2["B6"] = "E-2"
ws2.column_dimensions["A"].width = 90
ws2.column_dimensions["B"].width = 12

path = os.path.join(OUT_DIR, "测试底稿.xlsx")
wb.save(path)
print(path)
