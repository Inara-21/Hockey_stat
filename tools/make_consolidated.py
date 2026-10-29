# -*- coding: utf-8 -*-
"""Сводная книга Excel: все матчи со всеми внесёнными правками.

Каждый дивизион каждой конференции за каждый месяц — на своём листе. Лист —
копия исходного листа ячейка в ячейку: номера строк и столбцов совпадают с
исходным файлом. Внесены правки проверяющих (реестр data/corrections) и
исправленные файлы проверяющих (data/source; прежние версии — в архиве).
Исправленные ячейки выделены зелёным, в примечании — что было.
Первый лист — сводка со ссылками на листы.

Запуск:  python3 tools/make_consolidated.py
"""
import copy
import glob
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

import corrections
from error_report import MONTHS, SRC_DIR
from hockey_data import read_workbook, read_sheet, norm, parse_player, FIRST_COL
from make_corrected_excel import write_fixes, note, FIXED

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ARCHIVE = os.path.join(ROOT, "data", "source_archive")
OUT = os.path.join(ROOT, "reports", "Vse_matchi_s_pravkami.xlsx")
QUESTION = PatternFill("solid", fgColor="FFFFFF00")
CONF_ORDER = {"vostok": "ВОСТОК", "zapad": "ЗАПАД", "centr": "ЦЕНТР"}


def sheet_name(fname):
    m = re.match(r"(\d{4})-(\d{2})_([a-z]+)_div(\d+)", fname)
    y, mo, conf, div = int(m.group(1)), int(m.group(2)), m.group(3), m.group(4)
    return f"{CONF_ORDER[conf]} {div} дивизион {MONTHS[mo]} {y}", (list(CONF_ORDER).index(conf), int(div), y, mo)


def mark_reviewer_file(ws, fname):
    """Правки, пришедшие исправленным файлом проверяющего: сравнение с первой версией."""
    old_paths = sorted(glob.glob(os.path.join(ARCHIVE, "*", fname)))
    if not old_paths:
        return 0, None
    old = openpyxl.load_workbook(old_paths[0]).worksheets[0]
    n = 0
    for row in ws.iter_rows(min_row=1, max_row=max(ws.max_row, old.max_row),
                            max_col=max(ws.max_column, old.max_column)):
        for c in row:
            if isinstance(c, MergedCell):
                continue
            was = old.cell(c.row, c.column).value
            if was != c.value:
                note(c, "Исправлено проверяющим (исправленный файл от 13.10.2026).\n"
                        f"Было: {was if was not in (None, '') else 'пусто'}", FIXED)
                n += 1
    return n, os.path.basename(os.path.dirname(old_paths[0]))


def copy_sheet(src, dst, row_map=lambda r: r, rows=None):
    """Копия листа; row_map — куда ставить строку исходника; rows — только эти строки."""
    for row in src.iter_rows(min_row=rows[0] if rows else 1, max_row=rows[-1] if rows else src.max_row):
        for c in row:
            if isinstance(c, MergedCell):
                continue
            n = dst.cell(row=row_map(c.row), column=c.column, value=c.value)
            if c.has_style:
                n.font = copy.copy(c.font)
                n.fill = copy.copy(c.fill)
                n.border = copy.copy(c.border)
                n.alignment = copy.copy(c.alignment)
                n.number_format = c.number_format
                n.protection = copy.copy(c.protection)
            if c.comment:
                cm = Comment(c.comment.text, c.comment.author)
                cm.width, cm.height = c.comment.width, c.comment.height
                n.comment = cm
    for k, d in src.column_dimensions.items():
        if d.width:
            dst.column_dimensions[k].width = d.width
        dst.column_dimensions[k].hidden = d.hidden
    for k, d in src.row_dimensions.items():
        if d.height and (not rows or rows[0] <= k <= rows[-1]):
            dst.row_dimensions[row_map(k)].height = d.height
    if rows:
        return
    for rng in src.merged_cells.ranges:
        dst.merge_cells(start_row=row_map(rng.min_row), start_column=rng.min_col,
                        end_row=row_map(rng.max_row), end_column=rng.max_col)
    dst.sheet_format = copy.copy(src.sheet_format)
    dst.freeze_panes = src.freeze_panes


def write_rosters(ws, applied):
    """Верный состав команды: в строке состава меняются ячейки заменённых игроков."""
    for a in applied:
        g, fx = a["game"], a["fix"]
        r = g["labels_seen"][a["side"]] + 1
        cells = {parse_player(c.value): c for c in ws[r][FIRST_COL - 1:] if norm(c.value)}
        for (num, was), (_, now) in a["swaps"]:
            c = cells.get((num, was))
            if c is None:
                raise RuntimeError(f"{ws.title} стр. {r}: нет ячейки «{was} №{num}»")
            old = c.value
            c.value = f"{now} №{num}"
            note(c, f"Исправлено по ответу проверяющего ({fx.get('note', '')}).\nБыло: {old}", FIXED)


def main():
    out = openpyxl.Workbook()
    summary = out.active
    summary.title = "Сводка"
    files = sorted((f for f in os.listdir(SRC_DIR) if f.endswith(".xlsx")), key=lambda f: sheet_name(f)[1])
    rows = []
    for fname in files:
        name, _ = sheet_name(fname)
        src_path = os.path.join(SRC_DIR, fname)
        games, _ = read_workbook(src_path)
        _, corr = corrections.apply(fname, copy.deepcopy(games))
        if corr["conflicts"]:
            raise RuntimeError(f"{fname}: правки не сходятся с файлом: {corr['conflicts']}")
        for a in corr["applied"]:
            if a["fix"].get("team_roster"):
                continue                      # составы пишутся ниже, на готовый лист
            extra = set(a["fix"]) - {"file", "row", "final", "periods", "shootout", "was_final",
                                     "was_periods", "judge", "was_judge", "note", "source",
                                     "label", "label_row", "was_label", "date", "was_date"}
            if extra:
                raise RuntimeError(f"{fname} стр. {a['fix']['row']}: правку вида {extra} записывать в ячейки пока не умею")
        wb = openpyxl.load_workbook(src_path)
        ws = wb.worksheets[0]
        n_file, archived = mark_reviewer_file(ws, fname)
        write_fixes(wb, corr["applied"])
        for m in corr["misplaced"]:
            g = m["game"]
            note(ws.cell(g["labels_seen"]["comp"], 2),
                 "МАТЧ ПОД ВОПРОСОМ: " + m["fix"].get("note", "") +
                 ". Удаляем из этого дивизиона и переносим в другой — ждём ответ.", QUESTION)
        # перенесённые отсюда матчи: строки очищаются, номера строк не меняются
        for m in corr["moved_out"]:
            fx, g = m["fix"], m["game"]
            for r in range(g["row"], g["row"] + 15):
                for c in ws[r]:
                    if not isinstance(c, MergedCell):
                        c.value = None
                        c.comment = None
            to_name, _ = sheet_name(fx["move_to"])
            a = ws.cell(g["row"], 1)
            a.value = f"Матч перенесён на лист {to_name}, стр. {fx['target_row']}"
            note(a, f"Исправлено по ответу проверяющего ({fx.get('note', '')}).\n"
                    f"Было: {g['date'].strftime('%d.%m.%Y')} {g['time']} {g['team1']} — {g['team2']}, "
                    f"итог {g['final_raw']}", FIXED)
        dst = out.create_sheet(name)
        if corr["moved_in"]:
            # перенесённые сюда матчи — сразу после последнего матча; ниже строки сдвигаются
            last_end = max(g["row"] for g in games) + 14
            shift = 15 * len(corr["moved_in"])
            copy_sheet(ws, dst, row_map=lambda r: r if r <= last_end else r + shift)
            for m in corr["moved_in"]:
                fx = m["fix"]
                src_ws = openpyxl.load_workbook(os.path.join(SRC_DIR, fx["file"])).worksheets[0]
                rows_ = list(range(fx["row"], fx["row"] + 15))
                copy_sheet(src_ws, dst, row_map=lambda r, d=fx["target_row"] - fx["row"]: r + d, rows=rows_)
                comp = dst.cell(m["game"]["labels_seen"]["comp"], 2)
                was = comp.value
                comp.value = fx["comp"]
                from_name, _ = sheet_name(fx["file"])
                note(comp, f"Матч перенесён с листа {from_name}, стр. {fx['row']}, по ответу проверяющего "
                           f"({fx.get('note', '')}).\nБыло: {was}", FIXED)
        else:
            copy_sheet(ws, dst)
        write_rosters(dst, [a for a in corr["applied"] if a["fix"].get("team_roster")])
        n_games = len(read_sheet(dst))
        n_fixed = sum(1 for row in dst.iter_rows() for c in row
                      if c.comment and c.fill.fgColor.rgb == FIXED.fgColor.rgb)
        rows.append(dict(name=name, file=fname, comp=games[0]["comp"], month=ws.title, games=n_games,
                         fixes=n_fixed, misplaced=len(corr["misplaced"]), reviewer_file=bool(archived),
                         moved_in=len(corr["moved_in"]), moved_out=len(corr["moved_out"])))

    # --- сводка
    font = "Arial"
    thin = Side(style="thin", color="FFC7CBDC")
    box = Border(top=thin, bottom=thin, left=thin, right=thin)
    summary["A1"] = "ВСЕ МАТЧИ С ВНЕСЁННЫМИ ПРАВКАМИ"
    summary["A1"].font = Font(name=font, size=14, bold=True)
    summary["A2"] = (f"С последними правками от {corrections.LAST_DATE}. На каждом листе — копия исходного файла: "
                     "номера строк и столбцов совпадают с исходником.")
    summary["A3"] = "Зелёная ячейка — исправлена, в примечании к ней указано, что было. Жёлтая — матч под вопросом."
    for r in (2, 3):
        summary.cell(r, 1).font = Font(name=font, size=10, color="FF5A5F7A")
    head = ["Лист", "Соревнование", "Месяц", "Матчей на листе", "Исправлено ячеек", "Примечание"]
    for j, h in enumerate(head, 1):
        c = summary.cell(5, j, h)
        c.font = Font(name=font, size=10, bold=True)
        c.fill = PatternFill("solid", fgColor="FFEEF0F7")
        c.border = box
        c.alignment = Alignment(vertical="top", wrap_text=True)
    for i, r in enumerate(rows, 6):
        notes = []
        if r["reviewer_file"]:
            notes.append("исправленный файл проверяющего от 13.10.2026")
        if r["misplaced"]:
            notes.append(f"в том числе {r['misplaced']} матча под вопросом (не из этого дивизиона)")
        if r["moved_in"]:
            notes.append(f"добавлены матчи, перенесённые с другого листа: {r['moved_in']}")
        if r["moved_out"]:
            notes.append(f"перенесено на другой лист матчей: {r['moved_out']} (строки очищены)")
        vals = [r["name"], r["comp"], r["month"], r["games"], r["fixes"], "; ".join(notes)]
        for j, v in enumerate(vals, 1):
            c = summary.cell(i, j, v)
            c.font = Font(name=font, size=10)
            c.border = box
            c.alignment = Alignment(vertical="top", wrap_text=(j == 6))
        link = summary.cell(i, 1)
        link.hyperlink = Hyperlink(ref=link.coordinate, location=f"'{r['name']}'!A1")
        link.font = Font(name=font, size=10, color="FF2B2F77", underline="single")
    last = 6 + len(rows)
    summary.cell(last, 1, "Итого").font = Font(name=font, size=10, bold=True)
    summary.cell(last, 4, f"=SUM(D6:D{last - 1})").font = Font(name=font, size=10, bold=True)
    summary.cell(last, 5, f"=SUM(E6:E{last - 1})").font = Font(name=font, size=10, bold=True)
    summary.cell(last + 2, 1, "Число матчей подсчитано программой проверки по блокам матчей на каждом листе.").font = \
        Font(name=font, size=9, color="FF5A5F7A")
    for col, w in zip("ABCDEF", (32, 34, 14, 16, 17, 60)):
        summary.column_dimensions[col].width = w
    out.calculation.fullCalcOnLoad = True
    out.save(OUT)
    print(f"готово: {OUT}")
    for r in rows:
        print(f"   {r['name']:32} матчей {r['games']:>4}  исправлено ячеек {r['fixes']:>3}  {r['file']}")
    return rows


if __name__ == "__main__":
    main()
