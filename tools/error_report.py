# -*- coding: utf-8 -*-
"""Общий отчёт об ошибках по всем исходным файлам.

Проверяет каждый файл из папки (по умолчанию data/new) и собирает единый
отчёт. Исходные файлы не меняются никогда. Если исправление ошибки однозначно
следует из самих данных, оно предлагается в тексте ошибки на подтверждение.

Запуск:
    python3 tools/error_report.py                 # все файлы из data/new
    python3 tools/error_report.py путь/к/файлу    # один файл
"""
import datetime
import difflib
import os
import re
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from hockey_data import read_workbook, parse_score

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(ROOT, "data", "source")
OUT = os.path.join(ROOT, "reports", "Otchet_ob_oshibkah.txt")

REQUIRED = ["comp", "date", "time", "team1", "team2", "roster1", "roster2",
            "secretary", "judge", "final", "periods"]
LABEL_RU = {"comp": "Название соревнования", "date": "Дата", "time": "Время начала матча",
            "team1": "Команда 1", "team2": "Команда 2", "roster1": "Состав команды 1",
            "roster2": "Состав команды 2", "secretary": "Секретари", "judge": "Судьи",
            "final": "Итоговый счёт матча", "periods": "Счёт по периодам"}

ERROR, QUESTION = "ОШИБКА", "ВОПРОС"

# Как спрашивать по каждому виду вопроса. Текст находки содержит только факты,
# сам вопрос задаётся один раз на вид — и в тексте, и в Word.
ASK = {
    "игроки выходят за разные команды":
        "Допускается ли, что игрок выступает за несколько команд, в каждой под своим "
        "номером? Если это предусмотрено регламентом соревнований, вопрос снимается. "
        "Если нет — это ошибки в составах, и их нужно исправить.",
    "игрок за обе команды в одном матче":
        "Допускается ли, что игрок в одном матче выступает за обе команды, у каждой под "
        "своим номером? Если это предусмотрено регламентом соревнований, вопрос "
        "снимается. Если нет — это ошибки в составах, и их нужно исправить.",
}
DEFAULT_ASK = "Это нормальная практика или ошибка в данных? Просим подтвердить."

# Виды находок, сгруппированные по смыслу — для блока «Главное» в начале отчёта.
THEMES = [
    ("Буллиты",
     ["похоже на буллиты, но пометки нет"],
     "Если подтвердить правило простановки пометки, эти находки закрываются разом."),
    ("Счёт",
     ["счёт не сходится с периодами", "итоговый счёт записан наоборот",
      "неправдоподобный счёт", "пустые периоды"],
     "Итоговый счёт не сходится с тем, что записано по периодам."),
    ("Принадлежность матчей",
     ["команда встречается в единичных матчах", "дата из другого месяца",
      "разные названия соревнования"],
     "Затрагивает то, к какому соревнованию относятся матчи."),
    ("Составы команд",
     ["составы команд перепутаны местами", "один состав за разные команды",
      "состав из игроков другой команды", "состав скопирован в другую команду",
      "игрок в двух матчах одновременно"],
     "Один и тот же состав записан за разные команды."),
    ("Оформление составов",
     ["повтор игрового номера", "номер игрока меняется внутри команды",
      "игрок без номера", "игрок без ФИО",
      "посторонние символы в ФИО", "состав больше обычного"],
     "Дефекты записи состава, на результаты матчей не влияют."),
    ("Судьи и секретари",
     ["нет судьи", "нет секретаря", "судья в двух матчах одновременно",
      "секретарь в двух матчах одновременно"],
     "Не заполнены поля протокола."),
]

MONTHS = {1: "январь", 2: "февраль", 3: "март", 4: "апрель", 5: "май", 6: "июнь",
          7: "июль", 8: "август", 9: "сентябрь", 10: "октябрь", 11: "ноябрь", 12: "декабрь"}


def plural(n, one, few, many):
    """1 игрок, 3 игрока, 21 игрок, 11 игроков."""
    if n % 100 in range(11, 15):
        return many
    last = n % 10
    return one if last == 1 else few if last in (2, 3, 4) else many


def n_matches(n):
    """Предложный падеж: «в 1 матче», «в 3 матчах», «в 21 матче»."""
    word = "матче" if n % 10 == 1 and n % 100 != 11 else "матчах"
    return f"{n} {word}"


def in_matches(n):
    """«Встречается в 8 матчах» / пусто, если матч один — строка и так видна."""
    if n <= 1:
        return ""
    word = "матче" if n % 10 == 1 and n % 100 != 11 else "матчах"
    return f"Встречается в {n} {word}. "


def period_sums(g):
    if any(parse_score(p)[0] is None for p in g["periods_raw"]):
        return None, None
    return (sum(parse_score(p)[0] for p in g["periods_raw"]),
            sum(parse_score(p)[1] for p in g["periods_raw"]))


NAME_PREFIX = re.compile(r"^\s*\d+\s*[).．.]\s*")


def repair(games):
    """Находит ошибки, исправление которых однозначно следует из самого файла.

    Например, если в игровой день у всех остальных матчей один и тот же судья,
    то и у этого матча он тот же. Каждая такая ошибка всё равно попадает в
    отчёт: с предложенным исправлением и вопросом, подтверждает ли его
    проверяющий. Дальнейшие проверки идут по исправленным данным, чтобы одна
    ошибка не описывалась дважды. Если исправление неоднозначно, ничего не
    меняем — пропуск остаётся обычной ошибкой.
    """
    done = []

    by_day_judge, by_day_sec = defaultdict(set), defaultdict(set)
    num_of, name_of = defaultdict(set), defaultdict(set)
    for g in games:
        if g["judge"]:
            by_day_judge[(g["comp"], g["date"])].add(g["judge"])
        if g["secretary"]:
            by_day_sec[(g["comp"], g["date"])].add(g["secretary"])
        for team, roster in ((g["team1"], g["roster1"]), (g["team2"], g["roster2"])):
            for n, nm in roster:
                if n and nm:
                    num_of[(team, nm)].add(n)
                    name_of[(team, n)].add(nm)

    for g in games:
        for field, src, what in (("judge", by_day_judge, "судья"),
                                 ("secretary", by_day_sec, "секретарь")):
            if g[field]:
                continue
            cand = src.get((g["comp"], g["date"]), set())
            if len(cand) == 1:
                value = next(iter(cand))
                g[field] = value
                Cap = what.capitalize()
                done.append(dict(game=g, kind=f"нет {'судьи' if field == 'judge' else 'секретаря'}",
                                 text=f"{Cap} не указан. Во всех остальных матчах этого игрового "
                                      f"дня {what} — {value}.",
                                 fix=f"Очевидное исправление: {what} — {value}. Подтверждаете?"))

        for side, (team, roster) in enumerate(
                ((g["team1"], g["roster1"]), (g["team2"], g["roster2"]))):
            fixed = []
            for n, nm in roster:
                # нумерация списка, попавшая в ФИО: «1) Козлов Артем Владимирович»
                # нумерация списка перед ФИО («1) Козлов …») — просто убираем,
                # в отчёт не выносим; номер игрока проверяется дальше как обычно
                if nm and NAME_PREFIX.match(nm):
                    nm = NAME_PREFIX.sub("", nm).strip() or nm
                if nm and not n and len(num_of.get((team, nm), ())) == 1:
                    n = next(iter(num_of[(team, nm)]))
                    done.append(dict(game=g, kind="игрок без номера",
                                     text=f"{nm} ({team}) — не указан игровой номер. Во всех "
                                          f"остальных матчах за эту команду у него номер {n}.",
                                     fix=f"Очевидное исправление: номер {n}. Подтверждаете?"))
                elif n and not nm and len(name_of.get((team, n), ())) == 1:
                    nm = next(iter(name_of[(team, n)]))
                    done.append(dict(game=g, kind="игрок без ФИО",
                                     text=f"В составе команды {team} у номера {n} не указано ФИО. "
                                          f"Во всех остальных матчах за эту команду под этим "
                                          f"номером играет {nm}.",
                                     fix=f"Очевидное исправление: {nm}. Подтверждаете?"))
                fixed.append((n, nm))
            g["roster1" if side == 0 else "roster2"] = fixed

    # Состав одной команды целиком скопирован в другую: в одном матче за обе
    # команды записаны одни и те же игроки. Восстанавливаем, только если это
    # однозначно: у одной команды этот состав стоит во всех её матчах, а у
    # другой во всех остальных матчах файла один и тот же свой состав.
    # Частичные замены (один-два игрока) не трогаем никогда.
    def players(r):
        return frozenset((n, nm) for n, nm in r if nm)

    appear = defaultdict(list)
    for g in games:
        appear[g["team1"]].append((g, "roster1"))
        appear[g["team2"]].append((g, "roster2"))

    def own_roster(team, skip):
        """Состав команды, если он одинаков во всех её матчах, кроме skip."""
        others = [h[k] for h, k in appear[team] if h is not skip]
        if len(others) < 2 or len({players(r) for r in others}) != 1:
            return None
        return others[0]

    for g in games:
        r1, r2 = players(g["roster1"]), players(g["roster2"])
        if not r1 or r1 != r2:
            continue
        found = []
        for key, team, other in (("roster1", g["team1"], g["team2"]),
                                 ("roster2", g["team2"], g["team1"])):
            mine, theirs = own_roster(team, g), own_roster(other, g)
            if mine and theirs and players(mine) != r1 and players(theirs) == r1:
                found.append((key, team, other, mine))
        if len(found) != 1:
            continue
        key, team, other, mine = found[0]
        n_other = len(appear[team]) - 1
        g[key] = list(mine)
        done.append(dict(
            game=g, kind="один состав за разные команды",
            text=(f"За обе команды записаны одни и те же игроки — состав команды {other}. "
                  f"У команды {team} во всех остальных {n_other} матчах файла один и тот же "
                  f"свой состав: "
                  + ", ".join(f"{nm} №{n}" if n else nm for n, nm in mine if nm) + "."),
            fix=(f"Очевидное исправление: записать команде {team} её постоянный состав. "
                 f"Подтверждаете?")))

    return done


def collect(path):
    """-> (games, problems, summary). problems — список словарей."""
    games, skipped = read_workbook(path)
    fname = os.path.basename(path)
    repaired = repair(games)
    problems = []

    def add(kind, game, what, text, rows=None, fix="", ready=True):
        """fix — то, что стоит в последнем столбце, где отвечает проверяющий:
        предложенное исправление с вопросом о подтверждении (ready=True) или
        вопрос без готового исправления (ready=False)."""
        problems.append(dict(kind=kind, file=fname, game=game, what=what, text=text,
                             rows=sorted(rows) if rows else [], fix=fix,
                             proposed=bool(fix) and ready))

    for r in repaired:
        add(ERROR, r["game"], r["kind"], r["text"], fix=r["fix"])

    # Повторяющиеся дефекты состава копим и выводим одной записью со списком
    # матчей: иначе одна и та же проблема размножается на десятки находок.
    no_number = defaultdict(list)
    dup_numbers = defaultdict(list)

    def rows_text(rows):
        shown = ", ".join(f"стр. {r}" for r in rows[:8])
        return shown + (f" и ещё {len(rows) - 8}" if len(rows) > 8 else "")

    # ---------- по каждому матчу
    for g in games:
        for f in REQUIRED:
            if f not in g["labels_seen"]:
                add(ERROR, g, "в блоке нет строки",
                    f"В блоке матча отсутствует строка {LABEL_RU[f]}.")

        if not g["comp"]:
            add(ERROR, g, "пустое соревнование", "Не указано название соревнования.")
        if g["date"] is None:
            add(ERROR, g, "дата не разобрана",
                f"Дата не читается, записано: {g['date_kind']}.")
        if g["time"] is None:
            add(ERROR, g, "время не разобрано",
                f"Время не читается, записано: {g['time_kind']}.")
        if not g["team1"] or not g["team2"]:
            add(ERROR, g, "пустая команда",
                f"Не заполнено название команды: {g['team1'] or '(пусто)'} — {g['team2'] or '(пусто)'}.")
        elif g["team1"] == g["team2"]:
            add(ERROR, g, "команда против себя",
                f"В обеих графах одна и та же команда: {g['team1']}. Кто был соперником?")
        if not g["roster1"] or not g["roster2"]:
            add(ERROR, g, "пустой состав",
                f"Состав не заполнен: команда А — {len(g['roster1'])} игроков, "
                f"команда В — {len(g['roster2'])}.")
        if not g["judge"]:
            add(ERROR, g, "нет судьи", "Судья не указан. Кто судил матч?")
        if not g["secretary"]:
            add(ERROR, g, "нет секретаря", "Секретарь не указан. Кто вёл протокол?")

        if not g["final_raw"]:
            add(ERROR, g, "пустой итоговый счёт", "Итоговый счёт не заполнен.")
        elif g["goals1"] is None:
            add(ERROR, g, "итог не разобран",
                f"Итоговый счёт не читается, записано: {g['final_raw']}.")
        if not g["periods_raw"]:
            add(ERROR, g, "пустые периоды", "Счёт по периодам не заполнен.")
        for p in g["periods_raw"]:
            if parse_score(p)[0] is None:
                add(ERROR, g, "период не разобран",
                    f"Счёт периода не читается, записано: {p}.")

        sh, sa = period_sums(g)
        has_sum = sh is not None and bool(g["periods_raw"])

        # неправдоподобные значения счёта; если есть сумма по периодам, это та же
        # опечатка, что и расхождение с периодами, — одна находка, а не две
        MAX_PLAUSIBLE = 15
        implausible = False
        for side, val in (("А", g["goals1"]), ("В", g["goals2"])):
            if val is not None and val > MAX_PLAUSIBLE:
                implausible = True
                if has_sum:
                    add(ERROR, g, "неправдоподобный счёт",
                        f"У команды {side} в итоге {val} голов — неправдоподобно много. "
                        f"Сумма по периодам — {sh}:{sa}.",
                        fix=f"Похоже на описку: должно быть {sh}:{sa} вместо "
                            f"{g['goals1']}:{g['goals2']}. Подтверждаете?")
                else:
                    add(ERROR, g, "неправдоподобный счёт",
                        f"У команды {side} в итоге {val} голов — неправдоподобно много. "
                        f"Похоже на опечатку в записи счёта.")

        # расхождение периодов и итога
        # при пустых периодах расхождение уже описано отдельной находкой
        if (not implausible and g["goals1"] is not None and has_sum
                and (sh, sa) != (g["goals1"], g["goals2"])):
            tie_plus_one = sh == sa and (g["goals1"] - sh) + (g["goals2"] - sa) == 1
            if g["shootout"] or g["overtime"]:
                if not tie_plus_one:
                    add(ERROR, g, "счёт не сходится с периодами",
                        "Счёт не сходится с суммой по периодам.")
            elif tie_plus_one:
                add(ERROR, g, "похоже на буллиты, но пометки нет",
                    "После трёх периодов ничья, а пометки Б Буллиты нет.",
                    fix="Похоже на победу по буллитам: поставить пометку Б Буллиты. "
                        "Подтверждаете?")
            elif (g["goals1"], g["goals2"]) == (sa, sh):
                by_periods = g["team1"] if sh > sa else g["team2"]
                by_final = g["team1"] if g["goals1"] > g["goals2"] else g["team2"]
                add(ERROR, g, "итоговый счёт записан наоборот",
                    f"Сумма по периодам — {sh}:{sa}, по ней победила команда {by_periods}. "
                    f"Итоговый счёт записан {g['goals1']}:{g['goals2']}, по нему победила "
                    f"команда {by_final}. Числа те же, но переставлены местами.",
                    fix=f"Кто победил на самом деле — {by_periods} или {by_final}? Что записано "
                        f"неверно: итоговый счёт, счёт по периодам или порядок команд?",
                    ready=False)
            else:
                add(ERROR, g, "счёт не сходится с периодами",
                    "Счёт не сходится с суммой по периодам.")

        for team, roster in ((g["team1"], g["roster1"]), (g["team2"], g["roster2"])):
            nums = [n for n, _ in roster]
            for n, c in Counter(nums).items():
                if c > 1 and n:
                    who_n = tuple(sorted(nm for nu, nm in roster if nu == n))
                    dup_numbers[(team, n, who_n)].append(g["row"])
            for n, nm in roster:
                if not nm:
                    add(ERROR, g, "игрок без ФИО",
                        f"В составе {team} есть номер {n}, но фамилия не указана. "
                        f"Кто это?")
                elif not n:
                    no_number[(team, nm)].append(g["row"])

    for (team, nm), rows in sorted(no_number.items()):
        add(ERROR, None, "игрок без номера",
            f"{nm} ({team}) — не указан игровой номер. {in_matches(len(rows))}"
            f"Какой у него номер?", rows)
    for (team, n, who_n), rows in sorted(dup_numbers.items()):
        add(ERROR, None, "повтор игрового номера",
            f"В составе {team} номер {n} стоит сразу у двух игроков: "
            f"{' и '.join(who_n)}. {in_matches(len(rows))}"
            f"У кого из них этот номер верный?", rows)

    # ---------- по файлу целиком
    key = lambda g: (g["comp"], str(g["date"]), g["time"], g["team1"], g["team2"])
    seen = {}
    for g in games:
        k = key(g)
        if k in seen:
            add(ERROR, g, "матч встречается дважды",
                f"Такой же матч уже записан в строке {seen[k]['row']}. Это два разных матча или запись продублирована?")
        else:
            seen[k] = g

    slots = defaultdict(list)
    for g in games:
        if g["date"] and g["time"]:
            slots[(g["date"], g["time"])].append(g)
    for (d, t), gs in sorted(slots.items()):
        cnt = Counter()
        for g in gs:
            cnt[g["team1"]] += 1
            cnt[g["team2"]] += 1
        for team, c in cnt.items():
            if c > 1:
                add(ERROR, gs[0], "команда в двух матчах одновременно",
                    f"{d.strftime('%d.%m.%Y')} в {t} команда {team} указана сразу в {c} матчах. "
                    f"Какое время верное?")

    # Один файл — один месяц и одно соревнование: дата или название из другого
    # месяца или другого турнира видна прямо по файлу, спрашивать не о чем.
    month = re.match(r"(\d{4})-(\d{2})", fname)
    if month:
        y, m = int(month.group(1)), int(month.group(2))
        off = [g for g in games if g["date"] and (g["date"].year, g["date"].month) != (y, m)]
        if off:
            add(ERROR, None, "дата из другого месяца",
                f"Файл за {m:02d}.{y}, а в {n_matches(len(off))} стоит другая дата: "
                + ", ".join(sorted({g["date"].strftime("%d.%m.%Y") for g in off}))
                + ". Эти матчи относятся к этому файлу?",
                rows=[g["row"] for g in off])

    comps = Counter(g["comp"] for g in games if g["comp"])
    if len(comps) > 1:
        main = comps.most_common(1)[0][0]
        odd = [g for g in games if g["comp"] and g["comp"] != main]
        add(ERROR, None, "разные названия соревнования",
            f"В файле указано несколько соревнований. Чаще всего — {main}; "
            f"в {n_matches(len(odd))} написано: "
            + ", ".join(f"{c} ({n})" for c, n in comps.most_common()[1:])
            + ". Это один турнир или в файл попали чужие матчи?",
            rows=[g["row"] for g in odd])

    # Один человек не может быть в двух местах сразу: это проверяется по
    # дате и времени, без вопросов организаторам.
    for (d, t), gs in sorted(slots.items()):
        if len(gs) < 2:
            continue
        where = defaultdict(set)
        for g in gs:
            for n, nm in g["roster1"] + g["roster2"]:
                if nm:
                    where[nm].add(g["row"])
        both = {nm: rs for nm, rs in where.items() if len(rs) > 1}
        if both:
            add(ERROR, gs[0], "игрок в двух матчах одновременно",
                f"{d.strftime('%d.%m.%Y')} в {t} идут разные матчи, и в составы сразу "
                f"двух из них вписаны: " + ", ".join(sorted(both)) + ". Какое время верное?",
                rows=sorted({r for rs in both.values() for r in rs}))

        for field, what in (("judge", "судья"), ("secretary", "секретарь")):
            seen_at = defaultdict(set)
            for g in gs:
                if g[field]:
                    seen_at[g[field]].add(g["row"])
            busy = {p: rs for p, rs in seen_at.items() if len(rs) > 1}
            if busy:
                add(ERROR, gs[0], f"{what} в двух матчах одновременно",
                    f"{d.strftime('%d.%m.%Y')} в {t} один и тот же {what} "
                    + ", ".join(sorted(busy))
                    + " записан сразу в несколько матчей. Какое время верное?",
                    rows=sorted({r for rs in busy.values() for r in rs}))

    # Один и тот же игрок в составах обеих команд одного матча. Если совпадает
    # почти весь состав — это скопированный состав (ошибка). Если совпадают
    # один-два игрока, у каждой команды под своим номером, — это вопрос:
    # допускается ли такое правилами.
    namesakes = defaultdict(lambda: dict(rows=set(), sides=Counter()))
    copied = []
    for g in games:
        r1 = {nm: n for n, nm in g["roster1"] if nm}
        r2 = {nm: n for n, nm in g["roster2"] if nm}
        both = set(r1) & set(r2)
        if not both:
            continue
        if len(both) * 2 >= min(len(r1), len(r2)):
            copied.append(g)
            continue
        for nm in both:
            namesakes[nm]["rows"].add(g["row"])
            namesakes[nm]["sides"][(g["team1"], r1[nm])] += 1
            namesakes[nm]["sides"][(g["team2"], r2[nm])] += 1

    ptm = defaultdict(Counter)
    player_rows = defaultdict(set)
    team_num_rows = defaultdict(lambda: defaultdict(list))   # (игрок, команда) -> номер -> строки
    for g in games:
        for team, roster in ((g["team1"], g["roster1"]), (g["team2"], g["roster2"])):
            for n, nm in roster:
                ptm[nm][team] += 1
                player_rows[nm].add(g["row"])
                if nm and n:
                    team_num_rows[(nm, team)][n].append(g["row"])

    # Состав, заметно больше обычного для этого файла.
    sizes = Counter()
    for g in games:
        for roster in (g["roster1"], g["roster2"]):
            if roster:
                sizes[len(roster)] += 1
    if sizes:
        typical = sizes.most_common(1)[0][0]
        oversized = defaultdict(list)
        for g in games:
            for team, roster in ((g["team1"], g["roster1"]), (g["team2"], g["roster2"])):
                if roster and len(roster) >= typical + 3:
                    oversized[(team, len(roster))].append(g["row"])
        for (team, n), rows in sorted(oversized.items()):
            add(ERROR, None, "состав больше обычного",
                f"В составе {team} заявлено {n} игроков, тогда как у остальных команд "
                f"обычно {typical}. {in_matches(len(rows))}"
                f"Это допустимо или в состав попали лишние игроки?", rows)

    # Один и тот же состав под разными названиями команд.
    # Это объективный факт, в отличие от догадки про «основную команду»:
    # в младших дивизионах игроки и правда кочуют, и счёт «сколько игроков
    # обычно выступают за другую команду» даёт ложные срабатывания.
    roster_teams = defaultdict(list)
    for g in games:
        for team, roster in ((g["team1"], g["roster1"]), (g["team2"], g["roster2"])):
            if team and len(roster) >= 3:
                roster_teams[frozenset(nm for _, nm in roster)].append((g["row"], team))

    swapped_rosters = set()
    conflicts = []
    for players, entries in roster_teams.items():
        if len({t for _, t in entries}) < 2:
            continue
        swapped_rosters |= players
        by_team = defaultdict(list)
        for row, team in entries:
            by_team[team].append(row)
        counts = sorted((len(rows), t) for t, rows in by_team.items())
        # строки, где состав записан за команду, под которой он встречается реже,
        # — это и есть подозрительные матчи
        if counts[0][0] == counts[-1][0]:
            anomaly = frozenset(r for rows in by_team.values() for r in rows)
        else:
            anomaly = frozenset(by_team[counts[0][1]])
        conflicts.append(dict(players=players, by_team=by_team, anomaly=anomaly))

    # Перестановка составов в матче даёт два зеркальных конфликта с одними и
    # теми же подозрительными строками — показываем это одной записью.
    grouped = defaultdict(list)
    for c in conflicts:
        grouped[c["anomaly"]].append(c)

    for anomaly, group in sorted(grouped.items(), key=lambda x: sorted(x[0])):
        teams = sorted({t for c in group for t in c["by_team"]})
        rows = sorted(anomaly)
        if len(group) >= 2 and len(teams) == 2:
            add(ERROR, None, "составы команд перепутаны местами",
                ("В этих матчах" if len(rows) > 1 else "В этом матче")
                + f" составы команд {teams[0]} и {teams[1]} поменяны "
                f"местами: игроки, указанные здесь в составе команды {teams[0]}, "
                f"в остальных матчах играют в составе команды {teams[1]}, и наоборот. "
                + ("В каком из матчей команды указаны верно?" if len(rows) > 1
                   else "Какие составы верные?"), rows)
            continue
        for c in group:
            by_team = c["by_team"]
            counts = sorted(((len(r), t) for t, r in by_team.items()), reverse=True)
            main_team = counts[0][1]
            odd = [(t, r) for t, r in by_team.items() if t != main_team]
            odd_rows = sorted(r for _, rr in odd for r in rr)
            odd_names = ", ".join(t for t, _ in odd)
            add(ERROR, None, "один состав за разные команды",
                f"Эти же {len(c['players'])} игроков в {len(by_team[main_team])} других "
                f"матчах указаны в составе команды {main_team}, а здесь — в составе "
                f"команды {odd_names}. Если состав принадлежит команде {main_team}, "
                f"то здесь указана не та команда. Какая команда играла на самом деле?",
                odd_rows)

    # имена с посторонними символами (нумерация, скобки, цифры в начале)
    bad_names = defaultdict(list)
    for g in games:
        for team, roster in ((g["team1"], g["roster1"]), (g["team2"], g["roster2"])):
            for _, nm in roster:
                if nm and (nm[0].isdigit() or nm[0] in "()[].,-"):
                    bad_names[(team, nm)].append(g["row"])
    for (team, nm), rows in sorted(bad_names.items()):
        clean = nm.lstrip("0123456789()[].,- ")
        add(ERROR, None, "посторонние символы в ФИО",
            f"В составе {team} игрок записан как {nm} — в начале лишние символы. "
            f"Должно быть {clean}?", rows)

    # Однотипные находки по игрокам — одним вопросом на всю группу.
    movers = [(nm, teams) for nm, teams in ptm.items()
              if len(teams) > 1 and nm not in swapped_rosters]
    if movers:
        lines_m = "; ".join(
            f"{nm}: {', '.join(f'{t} — {c}' for t, c in teams.most_common())}"
            for nm, teams in sorted(movers))
        word = plural(len(movers), "игрок", "игрока", "игроков")
        add(QUESTION, None, "игроки выходят за разные команды",
            f"За разные команды выходят {len(movers)} {word}. {lines_m}.",
            rows=sorted({r for nm, _ in movers for r in player_rows[nm]}))

    # Номер, сменившийся вместе с командой, покрыт вопросом о переходах между
    # командами. Ошибка — только когда номер меняется внутри одной команды:
    # называем номер, с которым игрок записан в большинстве матчей за неё.
    odd_num = defaultdict(list)          # (команда, строки) -> описания игроков
    for (nm, team), by_num in sorted(team_num_rows.items()):
        if len(by_num) < 2:
            continue
        usual = max(by_num, key=lambda n: len(by_num[n]))
        for n, rs in by_num.items():
            if n == usual:
                continue
            odd_num[(team, tuple(sorted(set(rs))))].append(
                f"{nm} — здесь номер {n}, в остальных матчах за эту команду — {usual}")
    for (team, rs), who in sorted(odd_num.items()):
        add(ERROR, None, "номер игрока меняется внутри команды",
            f"Номер игрока не совпадает с его номером в других матчах за команду {team}. "
            + "; ".join(who) + ". Какой номер верный?",
            rows=list(rs))

    # Команда, у которой матчей на порядок меньше, чем у остальных, — скорее
    # всего чужое или неверно указанное название.
    team_games = Counter()
    team_rows = defaultdict(list)
    for g in games:
        for team in (g["team1"], g["team2"]):
            if team:
                team_games[team] += 1
                team_rows[team].append(g["row"])
    if len(team_games) > 2:
        counts = sorted(team_games.values())
        median = counts[len(counts) // 2]
        rare = [(t, c) for t, c in sorted(team_games.items()) if c * 10 < median]
        if rare:
            listed = ", ".join(f"{t} — {c}" for t, c in rare)
            rare_rows = sorted({r for t, _ in rare for r in team_rows[t]})
            add(ERROR, None, "команда встречается в единичных матчах",
                f"У остальных команд файла порядка {median} матчей, а здесь совсем мало: "
                f"{listed}. Эти команды относятся к этому соревнованию или матчи попали "
                f"в файл по ошибке?", rare_rows)

    names = sorted({g["team1"] for g in games} | {g["team2"] for g in games})
    for i, a in enumerate(names):
        for b in names[i + 1:]:
            if difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio() > 0.85:
                add(ERROR, None, "похожие названия команд",
                    f"Названия {a} и {b} очень похожи. Это две разные команды или опечатка?")

    # Плотность календаря (сколько матчей в день, пустые дни) не проверяем:
    # расписание — не наша зона ответственности.
    byday = Counter(g["date"] for g in games if g["date"])

    # разбивка для сводной таблицы: соревнование + месяц -> число матчей
    breakdown = Counter()
    undated = 0
    for g in games:
        if g["date"]:
            breakdown[(g["comp"], g["date"].year, g["date"].month)] += 1
        else:
            undated += 1

    wb = openpyxl.load_workbook(path)
    summary = dict(
        breakdown=breakdown, undated=undated,
        file=fname, games=len(games), skipped=skipped,
        sheets=dict(Counter(g["sheet"] for g in games)),
        comps=dict(Counter(g["comp"] for g in games)),
        teams=len(names), players=len(ptm),
        shootouts=sum(1 for g in games if g["shootout"]),
        rosters=dict(sorted(Counter((len(g["roster1"]), len(g["roster2"])) for g in games).items())),
        dates=(min(byday), max(byday)) if byday else None,
        days=len(byday),
        hidden_rows=sum(1 for ws in wb.worksheets for d in ws.row_dimensions.values() if d.hidden),
        formulas=sum(1 for ws in wb.worksheets for row in ws.iter_rows()
                     for c in row if isinstance(c.value, str) and c.value.startswith("=")),
    )
    summary["repaired"] = repaired
    covered = {r for x in problems
               if x["what"] in ("составы команд перепутаны местами", "один состав за разные команды")
               for r in (x["rows"] or ([x["game"]["row"]] if x["game"] else []))}
    for g in copied:
        if g["row"] not in covered:
            add(ERROR, g, "состав скопирован в другую команду",
                f"Составы команд {g['team1']} и {g['team2']} в этом матче почти полностью "
                f"совпадают. Состав одной из команд скопирован по ошибке. Какой состав верный?")

    if namesakes:
        parts = []
        for nm, d in sorted(namesakes.items()):
            sides = "; ".join(f"{t} — №{n}" for (t, n) in sorted(d["sides"]))
            parts.append(f"{nm} ({sides}; в {n_matches(len(d['rows']))})")
        add(QUESTION, None, "игрок за обе команды в одном матче",
            "Один и тот же игрок в одном матче записан в составы обеих команд, "
            "у каждой под своим номером: " + "; ".join(parts) + ".",
            rows=sorted({r for d in namesakes.values() for r in d["rows"]}))

    return games, problems, summary


def fmt_rows(rows, limit=12):
    """Номера строк для первого столбца: «стр. 1, 16, 61 и ещё 5»."""
    if not rows:
        return ""
    shown = ", ".join(str(r) for r in rows[:limit])
    tail = f" и ещё {len(rows) - limit}" if len(rows) > limit else ""
    return f"{'стр.' if len(rows) == 1 else 'стр.'} {shown}{tail}"


def render(all_summaries, all_problems, path):
    """Отчёт: сначала ошибки по файлам, затем вопросы отдельным разделом.

    Ошибка — то, что данные противоречат сами себе. Вопрос — то, что ошибкой
    считать нельзя, пока не подтверждено правилами соревнования.
    """
    L = []
    say = L.append
    today = datetime.date.today().strftime("%d.%m.%Y")
    total_games = sum(s["games"] for s in all_summaries)
    errors = [q for q in all_problems if q["kind"] == ERROR]
    questions = [q for q in all_problems if q["kind"] == QUESTION]

    say("=" * 78)
    say(f"ОТЧЁТ О ПРОВЕРКЕ ИСХОДНЫХ ДАННЫХ          Дата проверки: {today}")
    say("=" * 78)
    proposed = sum(1 for q in errors if q.get("proposed"))
    say("Исходные файлы не изменялись.")
    say(f"Файлов: {len(all_summaries)}   матчей: {total_games}   "
        f"ошибок: {len(errors)} (из них с готовым исправлением на подтверждение: "
        f"{proposed})   вопросов: {len(questions)}")

    kinds = Counter(q["what"] for q in errors)
    if kinds:
        say("")
        say("=" * 78)
        say("ГЛАВНОЕ")
        say("=" * 78)
        say("")
        say("По видам:")
        for what, c in kinds.most_common():
            say(f"   {c:4}  {what}")

        say("")
        say("На что обратить внимание:")
        used, blocks = set(), []
        for title, whats, comment in THEMES:
            n = sum(kinds.get(w, 0) for w in whats)
            if not n:
                continue
            used.update(whats)
            blocks.append((n, title, whats, comment))
        for n, title, whats, comment in sorted(blocks, key=lambda x: -x[0]):
            detail = ", ".join(f"{kinds[w]} — {w}" for w in whats if kinds.get(w))
            say("")
            say(f"   {title}: {n} из {len(errors)}")
            say(f"      {detail}.")
            if comment:
                say(f"      {comment}")
        other = {w: c for w, c in kinds.items() if w not in used}
        if other:
            say("")
            say(f"   Прочее: {sum(other.values())}")
            say(f"      {', '.join(f'{c} — {w}' for w, c in other.items())}.")

    by_file = defaultdict(list)
    for q in all_problems:
        by_file[q["file"]].append(q)

    say("")
    say("=" * 78)
    say("ОШИБКИ ПО ФАЙЛАМ")
    say("=" * 78)

    for s in all_summaries:
        comps = list(s["comps"])
        period = ", ".join(f"{MONTHS[m]} {y}"
                           for (c, y, m) in sorted(s["breakdown"], key=lambda k: (k[1], k[2])))
        say("")
        say("-" * 78)
        say(f"{s['file']} — {', '.join(comps) if comps else 'соревнование не указано'}"
            + (f", {period}" if period else ""))
        line = (f"{s['games']} матчей, команд {s['teams']}, игроков {s['players']}, "
                f"побед по буллитам {s['shootouts']}")
        if s["undated"]:
            line += f", без даты {s['undated']}"
        say(line)
        if s["skipped"]:
            say("листы без матчей: "
                + ", ".join(f"{t} ({r} строк)" for t, r in s["skipped"]))
        say("-" * 78)

        group = [q for q in by_file.get(s["file"], []) if q["kind"] == ERROR]
        if not group:
            say("ошибок не найдено")
            continue
        say(f"ОШИБКИ ({len(group)}):")
        multi_comp = len(comps) > 1
        for i, q in enumerate(group, 1):
            g = q["game"]
            say("")
            if g is None:
                where = fmt_rows(q.get("rows"))
                say(f"{i}. {where}   {q['what']}" if where else f"{i}. {q['what']}")
                say(f"   {q['text']}")
                continue
            dd = g["date"].strftime("%d.%m") if g["date"] else "—"
            say(f"{i}. стр. {g['row']}   {dd} {g['time'] or '—'}   "
                f"{g['team1'] or '—'} — {g['team2'] or '—'}"
                + (f"   [{g['comp']}]" if multi_comp else ""))
            if g["final_raw"]:
                mark = f" ({' '.join(g['final_notes'])})" if g["final_notes"] else ""
                row = f"   итог {g['final_raw']}{mark}"
                if g["periods_raw"]:
                    sh, sa = period_sums(g)
                    row += f", периоды {' '.join(g['periods_raw'])}"
                    if sh is not None:
                        row += f" = {sh}:{sa}"
                say(row)
            say(f"   {q['text']}")
            if q.get("fix"):
                say(f"   >> {q['fix']}")

    if questions:
        say("")
        say("=" * 78)
        say("ВОПРОСЫ — НЕ ОШИБКИ, ТРЕБУЮТ ПОДТВЕРЖДЕНИЯ")
        say("=" * 78)
        say("")
        say("Ошибкой это считать нельзя, пока не подтверждено правилами соревнования.")
        say("Ниже перечислено, где именно встречается.")
        by_what = defaultdict(list)
        for q in questions:
            by_what[q["what"]].append(q)
        for what, items in by_what.items():
            say("")
            say("-" * 78)
            say(what.upper())
            say("-" * 78)
            say(ASK.get(what, DEFAULT_ASK))
            n = 0
            for s in all_summaries:
                mine = [q for q in items if q["file"] == s["file"]]
                if not mine:
                    continue
                n += 1
                period = ", ".join(f"{MONTHS[m]} {y}" for (c, y, m) in
                                   sorted(s["breakdown"], key=lambda k: (k[1], k[2])))
                comps = ", ".join(s["comps"]) if s["comps"] else "—"
                say("")
                say(f"{n}. {comps}, {period}")
                for q in mine:
                    say(f"   {q['text']}")

    total = Counter()
    for s in all_summaries:
        total.update(s["breakdown"])

    say("")
    say("=" * 78)
    say("СВОДНАЯ ТАБЛИЦА")
    say("=" * 78)
    if total:
        wcomp = max(30, max(len(c) for c, _, _ in total))
        say(f"  {'Соревнование':<{wcomp}}  {'Месяц':<16}  {'Игр':>6}")
        say(f"  {'-' * wcomp}  {'-' * 16}  {'-' * 6}")
        for (comp, y, m), n in sorted(total.items(), key=lambda x: (x[0][0], x[0][1], x[0][2])):
            say(f"  {comp:<{wcomp}}  {MONTHS[m] + ' ' + str(y):<16}  {n:>6}")
        say(f"  {'-' * wcomp}  {'-' * 16}  {'-' * 6}")
        say(f"  {'ИТОГО':<{wcomp}}  {'':<16}  {sum(total.values()):>6}")

    say("")
    say(f"ИТОГО: матчей — {total_games}, ошибок — {len(errors)}, "
        f"вопросов — {len(questions)}")
    say("=" * 78)

    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L) + "\n")
    return len(errors), len(questions)


def main(args):
    if args:
        files = [a for a in args if os.path.isfile(a)]
    else:
        files = [os.path.join(SRC_DIR, f) for f in sorted(os.listdir(SRC_DIR))
                 if f.lower().endswith((".xlsx", ".xlsm", ".xls"))]
    if not files:
        print(f"Не найдено файлов для проверки в {SRC_DIR}")
        return 1

    all_problems, all_summaries = [], []
    for path in files:
        games, problems, summary = collect(path)
        all_problems += problems
        all_summaries.append(summary)
        e = sum(1 for q in problems if q["kind"] == ERROR)
        print(f"  {os.path.basename(path)}: матчей {len(games)}, ошибок {e}, "
              f"вопросов {len(problems) - e}")

    e, q = render(all_summaries, all_problems, OUT)
    print(f"\nвсего: ошибок {e}, вопросов {q}")
    print(f"отчёт: {OUT}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
