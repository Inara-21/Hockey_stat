# -*- coding: utf-8 -*-
"""Реестр исправлений по ответам проверяющих.

Исходные Excel-файлы не меняются. Исправления хранятся в data/corrections/*.json
и накладываются на прочитанные матчи при каждой проверке. В каждой правке
записано прежнее значение: если исходный файл с тех пор изменился и значение
не совпадает, правка не применяется, а расхождение попадает в отчёт.
"""
import fnmatch
import glob
import json
import os

from hockey_data import parse_score

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CORR_DIR = os.path.join(ROOT, "data", "corrections")
SHOOTOUT_NOTE = "Б Буллиты"


def _load():
    fixes, rules, sources, dates, scoped = [], {}, [], [], []
    for path in sorted(glob.glob(os.path.join(CORR_DIR, "*.json"))):
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        sources.append(data.get("source", os.path.basename(path)))
        if data.get("applied_on"):
            dates.append(data["applied_on"])
        rules.update(data.get("rules", {}))
        # правила действуют на файлы из scope (по умолчанию — на все)
        scoped.append((data.get("scope", ["*"]), data.get("rules", {})))
        for fx in data.get("fixes", []):
            fixes.append(dict(fx, source=data.get("source", "")))
    return fixes, rules, sources, dates, scoped


FIXES, RULES, SOURCES, DATES, SCOPED_RULES = _load()


def rules_for(fname):
    """Ответы организаторов, относящиеся к этому файлу."""
    out = {}
    for patterns, rules in SCOPED_RULES:
        if any(fnmatch.fnmatch(fname, pt) for pt in patterns):
            out.update(rules)
    return out
# дата последних внесённых правок, ДД.ММ.ГГГГ
LAST_DATE = max(DATES, key=lambda d: d.split(".")[::-1]) if DATES else None


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
    """-> (games, итог). Применяет правки к матчам одного файла.

    итог — словарь: applied (применено сейчас), already (уже стояло в файле),
    conflicts (не применено: значение в файле не то, на которое рассчитана
    правка), misplaced (матчи не из этого дивизиона — открытый вопрос).
    """
    mine = [fx for fx in FIXES if fx["file"] == fname]
    if not mine:
        return games, dict(applied=[], conflicts=[], misplaced=[], already=[], questions={}, answers=[])
    by_row = {g["row"]: g for g in games}
    applied, conflicts, drop, misplaced, already, questions, answers = [], [], set(), [], [], {}, []
    for fx in mine:
        if fx.get("answer"):
            answers.append(fx)
            continue
        g = by_row.get(fx["row"])
        if g is None:
            conflicts.append(dict(fix=fx, game=None, why="в файле нет матча с такой строкой"))
            continue
        if fx.get("answer"):
            # Ответ без конкретной правки (например, «исправил у себя») — к ошибке.
            answers.append(fx)
            continue
        if "player" in fx:
            # Номер игрока: в составах этого матча у игрока с прежним номером ставим новый.
            hits = [(k, i) for k in ("roster1", "roster2") for i, (n, nm) in enumerate(g[k]) if nm == fx["player"]]
            if hits and all(g[k][i][0] == fx["number"] for k, i in hits):
                already.append(dict(fix=fx, game=g))
                continue
            todo = [(k, i) for k, i in hits if g[k][i][0] == fx.get("was_number", "")]
            if not todo:
                conflicts.append(dict(fix=fx, game=g, why=f"у игрока {fx['player']} в этом матче нет "
                                                         f"номера {fx.get('was_number') or '(пусто)'}"))
                continue
            for k, i in todo:
                g[k][i] = (fx["number"], fx["player"])
            applied.append(dict(fix=fx, game=g,
                                before=f"{fx['player']} — {('№' + fx['was_number']) if fx.get('was_number') else 'без номера'}",
                                after=f"{fx['player']} — №{fx['number']}"))
            continue
        if "roster" in fx:
            # Состав команды целиком (скопированный состав другой команды).
            key = "roster1" if g["team1"] == fx["team"] else "roster2" if g["team2"] == fx["team"] else None
            new_roster = [tuple(x) for x in fx["roster"]]
            if key is None:
                conflicts.append(dict(fix=fx, game=g, why=f"в матче нет команды {fx['team']}"))
                continue
            if g[key] == new_roster:
                already.append(dict(fix=fx, game=g))
                continue
            other = g["roster2" if key == "roster1" else "roster1"]
            if fx.get("was_copy_of_other") and set(g[key]) != set(other):
                conflicts.append(dict(fix=fx, game=g, why="состав в файле уже не совпадает с составом соперника"))
                continue
            g[key] = new_roster
            applied.append(dict(fix=fx, game=g, before=f"состав команды {fx['team']}: копия состава соперника",
                                after=f"состав команды {fx['team']}: её постоянный состав ({len(new_roster)} игроков)"))
            continue
        if fx.get("question"):
            # Ответ проверяющего непонятен — правки нет, к ошибке добавляется вопрос.
            questions[fx["row"]] = fx["question"]
            continue
        if fx.get("misplaced"):
            # Матч не из этого дивизиона; куда его деть — открытый вопрос.
            # Из проверки дивизиона он исключается, но в файле остаётся.
            g["misplaced"] = fx.get("note", "")
            drop.add(fx["row"])
            misplaced.append(dict(fix=fx, game=g))
            continue
        # Правку можно применять повторно: если в файле уже стоит исправленное
        # значение (например, это исправленная копия), это не расхождение.
        done_final = "final" not in fx or parse_score(fx["final"]) == (g["goals1"], g["goals2"])
        done_periods = "periods" not in fx or _scores(fx["periods"]) == _scores(g["periods_raw"])
        done_shoot = not fx.get("shootout") or g["shootout"]
        if done_final and done_periods and done_shoot:
            already.append(dict(fix=fx, game=g))
            continue
        if "was_final" in fx and not done_final and parse_score(fx["was_final"]) != (g["goals1"], g["goals2"]):
            conflicts.append(dict(fix=fx, game=g, why=f"итог в файле {g['final_raw']}, "
                                                     f"а правка рассчитана на {fx['was_final']}"))
            continue
        if "was_periods" in fx and not done_periods and _scores(fx["was_periods"]) != _scores(g["periods_raw"]):
            conflicts.append(dict(fix=fx, game=g, why="счёт по периодам в файле отличается от того, "
                                                     "на который рассчитана правка"))
            continue
        before = describe(g)
        if "final" in fx:
            g["final_raw"] = fx["final"]
            g["goals1"], g["goals2"] = parse_score(fx["final"])
        if "periods" in fx:
            g["periods_raw"] = list(fx["periods"])
        if fx.get("shootout") and not g["shootout"]:
            g["final_notes"] = list(g["final_notes"]) + [SHOOTOUT_NOTE]
            g["shootout"] = True
        applied.append(dict(fix=fx, game=g, before=before, after=describe(g)))
    kept = [g for g in games if g["row"] not in drop]
    return kept, dict(applied=applied, conflicts=conflicts, misplaced=misplaced, already=already,
                      questions=questions, answers=answers)
