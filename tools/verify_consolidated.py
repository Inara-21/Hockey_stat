# -*- coding: utf-8 -*-
"""Проверка итоговой книги reports/Vse_matchi_s_pravkami.xlsx.

На каждом листе матчи должны совпадать с исходником после всех правок реестра,
а каждая ячейка, не выделенная зелёным, — с исходником ячейка в ячейку (номера
строк и столбцов те же). Время начала в книге приведено к ЧЧ:ММ: такие ячейки
сравниваются по самому времени.

Запуск:  python3 tools/verify_consolidated.py      (последняя строка: ИТОГ)
"""
import sys, os, copy
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from openpyxl.cell.cell import MergedCell
import corrections
from hockey_data import read_workbook, read_sheet, parse_time, label_key
import datetime
from make_consolidated import sheet_name
from make_corrected_excel import FIXED
SRC = os.path.join(ROOT, "data", "source")
wb = openpyxl.load_workbook(os.path.join(ROOT, "reports", "Vse_matchi_s_pravkami.xlsx"))
KEYS = ("date", "time", "team1", "team2", "final_raw", "periods_raw", "roster1", "roster2", "judge", "secretary", "shootout")
bad = 0
for f in sorted(x for x in os.listdir(SRC) if x.endswith(".xlsx")):
    name, _ = sheet_name(f)
    ws = wb[name]
    exp, corr = corrections.apply(f, read_workbook(os.path.join(SRC, f))[0])
    got = read_sheet(ws)
    e = {g["row"]: g for g in exp}; gt = {g["row"]: g for g in got if not (g.get("comp") or "").startswith("Матч перенесён")}
    rows_missing = set(e) - set(gt); rows_extra = set(gt) - set(e)
    diffs = [(r, k) for r in set(e) & set(gt) for k in KEYS if e[r].get(k) != gt[r].get(k)]
    src = openpyxl.load_workbook(os.path.join(SRC, f)).worksheets[0]
    moved_in_rows = {m["fix"]["target_row"] + i for m in corr["moved_in"] for i in range(15)}
    shift = 15 * len(corr["moved_in"])
    last_end = max(g["row"] for g in read_sheet(src)) + 14
    cell_diff = 0
    time_fmt = 0
    for row in src.iter_rows():
        for c in row:
            if isinstance(c, MergedCell): continue
            r = c.row if c.row <= last_end else c.row + shift
            d = ws.cell(r, c.column)
            if label_key(src.cell(c.row, 1).value) == "time" and c.column == 2 and d.value != c.value:
                # время приведено к ЧЧ:ММ: само время то же, значение — время, формат hh:mm
                if (parse_time(d.value)[0] == parse_time(c.value)[0] and isinstance(d.value, datetime.time)
                        and d.number_format == "hh:mm"):
                    time_fmt += 1
                    continue
            if d.value != c.value and not (d.comment and d.fill.fgColor.rgb == FIXED.fgColor.rgb) and r not in moved_in_rows:
                if not any(m["game"]["row"] <= c.row < m["game"]["row"] + 15 for m in corr["moved_out"]):
                    cell_diff += 1
    ok = not rows_missing and not rows_extra and not diffs and not cell_diff
    bad += not ok
    print(f"{name:32} матчей {len(got):4} время к ЧЧ:ММ {time_fmt:4} {'OK' if ok else 'ОШИБКА'}", rows_missing or "", rows_extra or "", diffs[:5] or "", cell_diff or "")
print("ИТОГ:", "всё сходится" if not bad else f"расхождений в {bad} листах")
