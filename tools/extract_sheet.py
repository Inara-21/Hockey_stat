# -*- coding: utf-8 -*-
"""Файл месяца из присланной книги: копия одного листа, остальные листы удаляются.
Ячейки и номера строк не меняются — строки в отчёте совпадают с листом присланной книги.

Запуск:  python3 -I tools/extract_sheet.py <книга.xlsx> <имя листа> data/source/ГГГГ-ММ_<конф>_div<N>.xlsx
После — записать, откуда взят лист, в data/source/origins.json.
"""
import sys, openpyxl
src, sheet, dst = sys.argv[1], sys.argv[2], sys.argv[3]
wb = openpyxl.load_workbook(src)
names = [ws.title for ws in wb.worksheets]
assert sheet in names, (sheet, names)
for ws in list(wb.worksheets):
    if ws.title != sheet:
        wb.remove(ws)
wb.active = 0
wb.save(dst)
# проверка: значения ячеек совпадают с исходным листом
a = openpyxl.load_workbook(src)[sheet]; b = openpyxl.load_workbook(dst).worksheets[0]
diff = sum(1 for row in a.iter_rows() for c in row if b.cell(c.row, c.column).value != c.value)
print(dst.rsplit("/", 1)[-1], "<-", sheet, f"| строк {a.max_row}, отличающихся ячеек: {diff}")
