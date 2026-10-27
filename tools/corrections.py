# -*- coding: utf-8 -*-
"""Реестр исправлений по ответам проверяющих.

Исходные Excel-файлы не меняются. Исправления хранятся в data/corrections/*.json
и накладываются на прочитанные матчи при каждой проверке. В каждой правке
записано прежнее значение: если исходный файл с тех пор изменился и значение
не совпадает, правка не применяется, а расхождение попадает в отчёт.
"""
import copy
import fnmatch
import glob
import json
import os

from hockey_data import parse_score, read_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CORR_DIR = os.path.join(ROOT, "data", "corrections")
SHOOTOUT_NOTE = "Б Буллиты"


def _load():
    fixes, rules, sources, dates, scoped = [], {}, [], [], []
    global GENERAL_QUESTIONS
    GENERAL_QUESTIONS = []
    for path in sorted(glob.glob(os.path.join(CORR_DIR, "*.json"))):
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        sources.append(data.get("source", os.path.basename(path)))
        if data.get("applied_on"):
            dates.append(data["applied_on"])
        rules.update(data.get("rules", {}))
        # правила действуют на файлы из scope (по умолчанию — на все)
        scoped.append((data.get("scope", ["*"]), data.get("rules", {})))
        GENERAL_QUESTIONS += [q for q in data.get("general_questions", [])
                              if not q.get("answered") and not q.get("superseded")]
        for fx in data.get("fixes", []):
            fixes.append(dict(fx, source=data.get("source", "")))
    return fixes, rules, sources, dates, scoped


GENERAL_QUESTIONS = []
FIXES, RULES, SOURCES, DATES, SCOPED_RULES = _load()


_SOURCE_CACHE = {}


def _source_games(fname):
    """Матчи исходного файла (для переноса матча из одного файла в другой)."""
    if fname not in _SOURCE_CACHE:
        _SOURCE_CACHE[fname] = read_workbook(os.path.join(ROOT, "data", "source", fname))[0]
    return _SOURCE_CACHE[fname]


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
    incoming = [fx for fx in FIXES if fx.get("move_to") == fname and not fx.get("superseded")]
    if not mine and not incoming:
        return games, dict(applied=[], conflicts=[], misplaced=[], already=[], questions={}, answers=[],
                           open_questions=[], moved_out=[], moved_in=[])
    by_row = {g["row"]: g for g in games}
    applied, conflicts, drop, misplaced, already, questions, answers, open_q = [], [], set(), [], [], {}, [], []
    moved_out, moved_in, team_rosters = [], [], []
    for fx in mine:
        if fx.get("superseded"):
            # правка закрыта: например, проверяющий прислал исправленный файл
            continue
        if fx.get("answer"):
            answers.append(fx)
            continue
        if fx.get("open_question"):
            # вопрос к проверяющему по его ответу — выводится в отчёте отдельной строкой
            open_q.append(fx)
            continue
        if fx.get("team_roster"):
            team_rosters.append(fx)          # после переноса матчей: им тоже нужна правка
            continue
        g = by_row.get(fx["row"])
        if fx.get("move_to"):
            # Матч переносится в другой файл: здесь его больше нет.
            if g is None:
                already.append(dict(fix=fx, game=None))          # уже убран из файла
            elif [g["team1"], g["team2"]] != fx["teams"]:
                conflicts.append(dict(fix=fx, game=g, why="в этой строке другой матч"))
            else:
                drop.add(fx["row"])
                moved_out.append(dict(fix=fx, game=g))
            continue
        if g is None:
            conflicts.append(dict(fix=fx, game=None, why="в файле нет матча с такой строкой"))
            continue
        if "judge" in fx:
            if g["judge"] == fx["judge"]:
                already.append(dict(fix=fx, game=g))
            elif g["judge"] != fx.get("was_judge", ""):
                conflicts.append(dict(fix=fx, game=g, why=f"судья в файле: {g['judge'] or 'не указан'}"))
            else:
                g["judge"] = fx["judge"]
                applied.append(dict(fix=fx, game=g, before="судья не указан", after=f"судья {fx['judge']}"))
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
    # Матчи, перенесённые сюда из другого файла, — в конец, в заранее назначенные строки.
    for fx in incoming:
        here = next((g for g in games if g["row"] == fx["target_row"]), None)
        if here is not None:
            if [here["team1"], here["team2"]] == fx["teams"]:
                already.append(dict(fix=fx, game=here))
            else:
                conflicts.append(dict(fix=fx, game=here, why=f"строка {fx['target_row']} уже занята другим матчем"))
            continue
        origin = next((g for g in _source_games(fx["file"]) if g["row"] == fx["row"]), None)
        if origin is None or [origin["team1"], origin["team2"]] != fx["teams"]:
            conflicts.append(dict(fix=fx, game=None, why=f"в файле {fx['file']} нет переносимого матча"))
            continue
        g = copy.deepcopy(origin)
        shift = fx["target_row"] - g["row"]
        g["labels_seen"] = {k: r + shift for k, r in g["labels_seen"].items()}
        g["row"], g["comp"] = fx["target_row"], fx["comp"]
        g["moved_from"] = (fx["file"], fx["row"])
        games = games + [g]
        moved_in.append(dict(fix=fx, game=g, origin=origin))

    for fx in team_rosters:
        # Верный состав команды от проверяющего: во всех матчах команды в файле
        # игроки под этими номерами заменяются. Порядок игроков — как в файле.
        was = [tuple(x) for x in fx["was_roster"]]
        swap = {(r["number"], r["was"]): (r["number"], r["player"]) for r in fx["replace"]}
        now = [swap.get(p, p) for p in was]
        for g in games:
            for key, team in (("roster1", g["team1"]), ("roster2", g["team2"])):
                if team != fx["team"]:
                    continue
                if g[key] == now:
                    already.append(dict(fix=fx, game=g))
                elif g[key] != was:
                    conflicts.append(dict(fix=fx, game=g, why=f"состав команды {team} в этом матче "
                                                             f"не тот, на который рассчитана правка"))
                else:
                    g[key] = list(now)
                    swaps = [(p, swap[p]) for p in was if p in swap]
                    applied.append(dict(fix=fx, game=g, side=key, swaps=swaps,
                                        before=f"{team}: " + ", ".join(f"№{a[0]} {a[1]}" for a, _ in swaps),
                                        after=f"{team}: " + ", ".join(f"№{b[0]} {b[1]}" for _, b in swaps)))

    kept = [g for g in games if g["row"] not in drop]
    return kept, dict(applied=applied, conflicts=conflicts, misplaced=misplaced, already=already,
                      questions=questions, answers=answers, open_questions=open_q,
                      moved_out=moved_out, moved_in=moved_in)
