# -*- coding: utf-8 -*-
"""Чтение матчей из исходных таблиц.

Поля ищутся ПО ПОДПИСИ в левом столбце, а не по номеру строки: в выгрузках
встречаются переставленные строки («Команда 2» раньше «Команда 1»), опечатки
в подписях и разные форматы дат, времени и счёта. Значения переносятся как
есть — ничего не пересчитывается и не исправляется.
"""
import datetime
import re

BLOCK_ROWS = 15
FIRST_COL, LAST_COL_DEFAULT = 2, 8

# Варианты написания подписей. Ключ — внутреннее имя поля.
LABELS = {
    "comp":      ("название соревнования", "навание соревнования"),
    "date":      ("дата",),
    "time":      ("время начала матча", "время"),
    "team1":     ("команда 1", "команда1"),
    "team2":     ("команда 2", "команда2"),
    "roster1":   ("состав команды 1", "состав команды1"),
    "roster2":   ("состав команды 2", "состав команды2"),
    "secretary": ("секретари", "секретарь"),
    "judge":     ("судьи", "судья"),
    "final":     ("итоговый счёт матча", "итоговый счет матча", "итоговый счёт"),
    "periods":   ("счёт по периодам", "счет по периодам"),
}
# Поля, значения которых лежат в СЛЕДУЮЩЕЙ строке под подписью
VALUE_ON_NEXT_ROW = {"roster1", "roster2", "secretary", "judge"}
# Поля, занимающие несколько столбцов
MULTI_COLUMN = {"roster1", "roster2", "periods"}
# Поля, у которых правее значения могут стоять пометки (напр. «Б Буллиты»)
WITH_NOTES = {"final"}

SCORE_SEPARATORS = ("*", ":", "-", "–", "—", "х", "x")


def norm(v):
    """Текст ячейки: без неразрывных пробелов и лишних пробельных символов.

    Юникод-нормализацию NFKC здесь применять нельзя: она превращает «№» в «No»,
    и номера игроков перестают разбираться.
    """
    if v is None:
        return ""
    s = str(v).replace("\xa0", " ")
    return " ".join(s.split())


def label_key(text):
    """Подпись -> внутреннее имя поля, либо None."""
    t = norm(text).lower().rstrip(":").strip()
    for key, variants in LABELS.items():
        if t in variants:
            return key
    return None


def parse_date(v):
    """-> (date | None, вид). Принимает и дату Excel, и текст."""
    if isinstance(v, datetime.datetime):
        return datetime.date(v.year, v.month, v.day), "дата"
    if isinstance(v, datetime.date):
        return v, "дата"
    t = norm(v)
    if not t:
        return None, "пусто"
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%y"):
        try:
            return datetime.datetime.strptime(t, fmt).date(), "текст"
        except ValueError:
            pass
    return None, "неразобрано"


def parse_time(v):
    """-> (строка ЧЧ:ММ | None, вид)."""
    if isinstance(v, datetime.datetime):
        return v.strftime("%H:%M"), "время"
    if isinstance(v, datetime.time):
        return v.strftime("%H:%M"), "время"
    t = norm(v)
    if not t:
        return None, "пусто"
    m = re.fullmatch(r"(\d{1,2})[:.\-](\d{2})(?::\d{2})?", t)
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
        if 0 <= h < 24 and 0 <= mi < 60:
            return f"{h:02d}:{mi:02d}", "текст"
    return None, "неразобрано"


def parse_score(raw):
    """'7:3' или '3*5' -> (7, 3). Не разобралось -> (None, None)."""
    t = norm(raw)
    if not t:
        return None, None
    for sep in SCORE_SEPARATORS:
        if sep in t:
            left, _, right = t.partition(sep)
            try:
                return int(left.strip()), int(right.strip())
            except ValueError:
                return None, None
    return None, None


def score_text(raw):
    """Счёт для печати: '7:3' -> '7 : 3'. Неразобранное — без изменений."""
    a, b = parse_score(raw)
    return f"{a} : {b}" if a is not None else norm(raw)


def parse_player(raw):
    """'Иванов Иван №17' -> ('17', 'Иванов Иван'). Без № — весь текст как ФИО."""
    t = norm(raw)
    if not t:
        return "", ""
    if "№" in t:
        name, _, num = t.rpartition("№")
        return num.strip(), name.strip()
    return "", t


def find_blocks(ws):
    """Строки, с которых начинаются блоки игр."""
    starts = []
    for r in range(1, ws.max_row + 1):
        if label_key(ws.cell(row=r, column=1).value) == "comp":
            starts.append(r)
    return starts


def read_sheet(ws):
    """Матчи одного листа. Каждый — словарь; row — строка начала блока."""
    starts = find_blocks(ws)
    last_col = max(ws.max_column, LAST_COL_DEFAULT)
    games = []

    for i, start in enumerate(starts):
        end = starts[i + 1] - 1 if i + 1 < len(starts) else min(ws.max_row, start + BLOCK_ROWS - 1)
        game = {"sheet": ws.title, "row": start, "labels_seen": {}}

        for r in range(start, end + 1):
            key = label_key(ws.cell(row=r, column=1).value)
            if key is None or key in game.get("labels_seen", {}):
                continue
            game["labels_seen"][key] = r
            value_row = r + 1 if key in VALUE_ON_NEXT_ROW else r
            if value_row > end:
                continue
            if key in MULTI_COLUMN:
                vals = [norm(ws.cell(row=value_row, column=c).value)
                        for c in range(FIRST_COL, last_col + 1)]
                game[key] = [v for v in vals if v]
            else:
                game[key] = ws.cell(row=value_row, column=FIRST_COL).value
                if key in WITH_NOTES:
                    # правее значения могут стоять пометки об исходе матча
                    notes = [norm(ws.cell(row=value_row, column=c).value)
                             for c in range(FIRST_COL + 1, last_col + 1)]
                    game[key + "_notes"] = [n for n in notes if n]

        # нормализация
        game["comp"] = norm(game.get("comp"))
        game["date"], game["date_kind"] = parse_date(game.get("date"))
        game["time"], game["time_kind"] = parse_time(game.get("time"))
        game["team1"] = norm(game.get("team1"))
        game["team2"] = norm(game.get("team2"))
        game["secretary"] = norm(game.get("secretary"))
        game["judge"] = norm(game.get("judge"))
        game["final_raw"] = norm(game.get("final"))
        game["final_notes"] = game.get("final_notes") or []
        game["goals1"], game["goals2"] = parse_score(game["final_raw"])
        game["shootout"] = any("буллит" in n.lower() for n in game["final_notes"])
        game["overtime"] = any("оверта" in n.lower() for n in game["final_notes"])
        game["periods_raw"] = game.get("periods") or []
        game["roster1"] = [parse_player(p) for p in (game.get("roster1") or [])]
        game["roster2"] = [parse_player(p) for p in (game.get("roster2") or [])]
        games.append(game)

    return games


def read_workbook(path):
    """Все матчи книги. Листы без блоков игр пропускаются."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    games, skipped = [], []
    for ws in wb.worksheets:
        found = read_sheet(ws)
        if found:
            games.extend(found)
        else:
            skipped.append((ws.title, ws.max_row))
    return games, skipped


def winner(game):
    """Победитель по итоговому счёту. Поля в исходнике нет."""
    g1, g2 = game["goals1"], game["goals2"]
    if g1 is None or g2 is None:
        return ""
    if g1 > g2:
        return game["team1"]
    if g2 > g1:
        return game["team2"]
    return "ничья"
