# -*- coding: utf-8 -*-
"""Реестр исправлений по ответам проверяющих.

Исходные Excel-файлы не меняются. Исправления хранятся в data/corrections/*.json
и накладываются на прочитанные матчи при каждой проверке. В каждой правке
записано прежнее значение: если исходный файл с тех пор изменился и значение
не совпадает, правка не применяется, а расхождение попадает в отчёт.
"""
import glob
import json
import os

from hockey_data import parse_score

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CORR_DIR = os.path.join(ROOT, "data", "corrections")
SHOOTOUT_NOTE = "Б Буллиты"


def _load():
    fixes, rules, sources = [], {}, []
    for path in sorted(glob.glob(os.path.join(CORR_DIR, "*.json"))):
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        sources.append(data.get("source", os.path.basename(path)))
        rules.update(data.get("rules", {}))
        for fx in data.get("fixes", []):
            fixes.append(dict(fx, source=data.get("source", "")))
    return fixes, rules, sources


FIXES, RULES, SOURCES = _load()


def _scores(items):
    return [parse_score(x) for x in items]


def describe(g):
    """Счёт матча одной строкой: «итог 5:4 (Б Буллиты), периоды 2:0 3:1 0:4»."""
    s = f"итог {g['final_raw'] or '—'}"
    if g["final_notes"]:
        s += f" ({' '.join(g['final_notes'])})"
    s += ", периоды " + (" ".join(g["periods_raw"]) if g["periods_raw"] else "не заполнены")
    return s


def apply(fname, games):
    """-> (games, applied, conflicts). Применяет правки к матчам одного файла."""
    mine = [fx for fx in FIXES if fx["file"] == fname]
    if not mine:
        return games, [], []
    by_row = {g["row"]: g for g in games}
    applied, conflicts, drop = [], [], set()
    for fx in mine:
        g = by_row.get(fx["row"])
        if g is None:
            conflicts.append(dict(fix=fx, game=None, why="в файле нет матча с такой строкой"))
            continue
        if "was_final" in fx and parse_score(fx["was_final"]) != (g["goals1"], g["goals2"]):
            conflicts.append(dict(fix=fx, game=g, why=f"итог в файле {g['final_raw']}, "
                                                     f"а правка рассчитана на {fx['was_final']}"))
            continue
        if "was_periods" in fx and _scores(fx["was_periods"]) != _scores(g["periods_raw"]):
            conflicts.append(dict(fix=fx, game=g, why="счёт по периодам в файле отличается от того, "
                                                     "на который рассчитана правка"))
            continue
        before = describe(g)
        if fx.get("exclude"):
            drop.add(fx["row"])
            applied.append(dict(fix=fx, game=g, before=before, after="матч исключён из файла"))
            continue
        if "final" in fx:
            g["final_raw"] = fx["final"]
            g["goals1"], g["goals2"] = parse_score(fx["final"])
        if "periods" in fx:
            g["periods_raw"] = list(fx["periods"])
        if fx.get("shootout") and not g["shootout"]:
            g["final_notes"] = list(g["final_notes"]) + [SHOOTOUT_NOTE]
            g["shootout"] = True
        applied.append(dict(fix=fx, game=g, before=before, after=describe(g)))
    return [g for g in games if g["row"] not in drop], applied, conflicts
