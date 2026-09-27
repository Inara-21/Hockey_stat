# -*- coding: utf-8 -*-
"""Общий отчёт об ошибках по всем исходным файлам.

Проверяет каждый файл из папки (по умолчанию data/new) и собирает единый
отчёт. Ничего не исправляет: только находит и описывает.

Запуск:
    python3 tools/error_report.py                 # все файлы из data/new
    python3 tools/error_report.py путь/к/файлу    # один файл
"""
import datetime
import difflib
import os
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

MONTHS = {1: "январь", 2: "февраль", 3: "март", 4: "апрель", 5: "май", 6: "июнь",
          7: "июль", 8: "август", 9: "сентябрь", 10: "октябрь", 11: "ноябрь", 12: "декабрь"}


def period_sums(g):
    if any(parse_score(p)[0] is None for p in g["periods_raw"]):
        return None, None
    return (sum(parse_score(p)[0] for p in g["periods_raw"]),
            sum(parse_score(p)[1] for p in g["periods_raw"]))


def collect(path):
    """-> (games, problems, summary). problems — список словарей."""
    games, skipped = read_workbook(path)
    fname = os.path.basename(path)
    problems = []

    def add(kind, game, what, text):
        problems.append(dict(kind=kind, file=fname, game=game, what=what, text=text))

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
                    f"В блоке матча отсутствует строка «{LABEL_RU[f]}».")

        if not g["comp"]:
            add(ERROR, g, "пустое соревнование", "Не указано название соревнования.")
        if g["date"] is None:
            add(ERROR, g, "дата не разобрана",
                f"Дата не читается (записана как «{g['date_kind']}»).")
        if g["time"] is None:
            add(ERROR, g, "время не разобрано",
                f"Время не читается (записано как «{g['time_kind']}»).")
        if not g["team1"] or not g["team2"]:
            add(ERROR, g, "пустая команда",
                f"Не заполнено название команды: «{g['team1']}» — «{g['team2']}».")
        elif g["team1"] == g["team2"]:
            add(ERROR, g, "команда против себя",
                f"В обеих графах одна и та же команда: {g['team1']}.")
        if not g["roster1"] or not g["roster2"]:
            add(ERROR, g, "пустой состав",
                f"Состав не заполнен: команда А — {len(g['roster1'])} игроков, "
                f"команда В — {len(g['roster2'])}.")
        if not g["judge"]:
            add(QUESTION, g, "нет судьи", "Судья не указан.")
        if not g["secretary"]:
            add(QUESTION, g, "нет секретаря", "Секретарь не указан.")

        if not g["final_raw"]:
            add(ERROR, g, "пустой итоговый счёт", "Итоговый счёт не заполнен.")
        elif g["goals1"] is None:
            add(ERROR, g, "итог не разобран",
                f"Итоговый счёт «{g['final_raw']}» не читается.")
        if not g["periods_raw"]:
            add(ERROR, g, "пустые периоды", "Счёт по периодам не заполнен.")
        for p in g["periods_raw"]:
            if parse_score(p)[0] is None:
                add(ERROR, g, "период не разобран",
                    f"Счёт периода «{p}» не читается.")

        # расхождение периодов и итога
        sh, sa = period_sums(g)
        if g["goals1"] is not None and sh is not None and (sh, sa) != (g["goals1"], g["goals2"]):
            tie_plus_one = sh == sa and (g["goals1"] - sh) + (g["goals2"] - sa) == 1
            if g["shootout"] or g["overtime"]:
                if not tie_plus_one:
                    add(QUESTION, g, "пометка не соответствует счёту",
                        "Стоит пометка о буллитах, но ничьей после основного времени "
                        "не было.")
            elif tie_plus_one:
                add(QUESTION, g, "похоже на буллиты, но пометки нет",
                    "После трёх периодов ничья, в итоге +1 гол — как у матчей с буллитами, "
                    "но пометки «Б Буллиты» нет. Буллиты или опечатка?")
            elif (g["goals1"], g["goals2"]) == (sa, sh):
                add(ERROR, g, "итоговый счёт записан наоборот",
                    "Те же числа, что и в сумме по периодам, но переставлены местами. "
                    "Похоже, итоговый счёт записан наоборот. Так и было или перепутали?")
            else:
                add(ERROR, g, "счёт не сходится с периодами",
                    "Счёт не сходится с суммой по периодам. Ничьей после трёх периодов "
                    "не было — буллитами не объясняется.")

        for team, roster in ((g["team1"], g["roster1"]), (g["team2"], g["roster2"])):
            nums = [n for n, _ in roster]
            for n, c in Counter(nums).items():
                if c > 1 and n:
                    who_n = tuple(sorted(nm for nu, nm in roster if nu == n))
                    dup_numbers[(team, n, who_n)].append(g["row"])
            for n, nm in roster:
                if not nm:
                    add(ERROR, g, "игрок без ФИО",
                        f"В составе «{team}» есть номер №{n} без фамилии.")
                elif not n:
                    no_number[(team, nm)].append(g["row"])

    for (team, nm), rows in sorted(no_number.items()):
        add(ERROR, None, "игрок без номера",
            f"У «{nm}» («{team}») не указан игровой номер — в {len(rows)} матчах: "
            f"{rows_text(rows)}.")
    for (team, n, who_n), rows in sorted(dup_numbers.items()):
        add(ERROR, None, "повтор игрового номера",
            f"В составе «{team}» номер №{n} у двух игроков: {', '.join(who_n)} — "
            f"в {len(rows)} матчах: {rows_text(rows)}.")

    # ---------- по файлу целиком
    key = lambda g: (g["comp"], str(g["date"]), g["time"], g["team1"], g["team2"])
    seen = {}
    for g in games:
        k = key(g)
        if k in seen:
            add(ERROR, g, "матч встречается дважды",
                f"Такой же матч записан в строке {seen[k]['row']}.")
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
                    f"{d.strftime('%d.%m.%Y')} в {t} команда «{team}» указана "
                    f"сразу в {c} матчах.")

    ptm, pnum = defaultdict(Counter), defaultdict(Counter)
    for g in games:
        for team, roster in ((g["team1"], g["roster1"]), (g["team2"], g["roster2"])):
            for n, nm in roster:
                ptm[nm][team] += 1
                pnum[nm][n] += 1

    # основная команда игрока — та, за которую он выходит чаще всего
    home = {nm: teams.most_common(1)[0][0] for nm, teams in ptm.items()}

    # Если больше половины состава — игроки другой команды, это ошибка
    # заполнения одного матча, а не история про каждого игрока отдельно.
    swapped_rosters = set()
    for g in games:
        for side, (team, roster) in enumerate(
                ((g["team1"], g["roster1"]), (g["team2"], g["roster2"])), 1):
            if not roster or not team:
                continue
            alien = Counter(home[nm] for _, nm in roster if home.get(nm) not in (None, team))
            if alien and sum(alien.values()) * 2 >= len(roster):
                other, cnt = alien.most_common(1)[0]
                add(ERROR, g, "состав из игроков другой команды",
                    f"В составе команды {'А' if side == 1 else 'В'} ({team}) "
                    f"{cnt} из {len(roster)} игроков в других матчах выступают "
                    f"за «{other}». Это заявленный на матч состав или состав "
                    f"скопирован не от той команды?")
                for _, nm in roster:
                    if home.get(nm) == other:
                        swapped_rosters.add(nm)

    for nm, teams in ptm.items():
        if len(teams) > 1 and nm not in swapped_rosters:
            add(ERROR, None, "игрок в разных командах",
                f"«{nm}» выходит за разные команды: "
                f"{', '.join(f'{t} ({c})' for t, c in teams.most_common())}. "
                f"Если переходы между командами допускаются — вопрос снимается, "
                f"иначе ошибка в составе.")
    for nm, nums in pnum.items():
        if len(nums) > 1:
            add(QUESTION, None, "игрок с разными номерами",
                f"У «{nm}» в разных матчах разные номера: "
                f"{', '.join(f'№{n} ({c} раз)' for n, c in nums.items())}.")

    names = sorted({g["team1"] for g in games} | {g["team2"] for g in games})
    for i, a in enumerate(names):
        for b in names[i + 1:]:
            if difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio() > 0.85:
                add(QUESTION, None, "похожие названия команд",
                    f"Названия «{a}» и «{b}» очень похожи — возможна опечатка.")

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
    return games, problems, summary


def render(all_summaries, all_problems, path):
    """Компактный отчёт: реквизиты один раз на файл, дальше только находки.

    Разряда «вопросы» нет — любое расхождение считается ошибкой; там, где
    нужно уточнение, вопрос стоит прямо в тексте находки.
    """
    L = []
    say = L.append
    today = datetime.date.today().strftime("%d.%m.%Y")
    total_games = sum(s["games"] for s in all_summaries)

    say("=" * 78)
    say(f"ОТЧЁТ О ПРОВЕРКЕ ИСХОДНЫХ ДАННЫХ          Дата проверки: {today}")
    say("=" * 78)
    say("Данные проверены без внесения изменений.")
    say(f"Файлов: {len(all_summaries)}   матчей: {total_games}   "
        f"ошибок: {len(all_problems)}")

    by_file = defaultdict(list)
    for p in all_problems:
        by_file[p["file"]].append(p)

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

        group = by_file.get(s["file"], [])
        if not group:
            say("ошибок не найдено")
            continue
        say(f"ОШИБКИ ({len(group)}):")
        multi_comp = len(comps) > 1
        for i, p in enumerate(group, 1):
            g = p["game"]
            say("")
            if g is None:
                say(f"{i}. {p['what']}")
                say(f"   {p['text']}")
                continue
            d = g["date"].strftime("%d.%m") if g["date"] else "—"
            say(f"{i}. стр. {g['row']}   {d} {g['time'] or '—'}   "
                f"{g['team1'] or '—'} — {g['team2'] or '—'}"
                + (f"   [{g['comp']}]" if multi_comp else ""))
            if g["final_raw"]:
                mark = f" ({' '.join(g['final_notes'])})" if g["final_notes"] else ""
                line = f"   итог {g['final_raw']}{mark}"
                if g["periods_raw"]:
                    sh, sa = period_sums(g)
                    line += f", периоды {' '.join(g['periods_raw'])}"
                    if sh is not None:
                        line += f" = {sh}:{sa}"
                say(line)
            say(f"   {p['text']}")

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
    say(f"ИТОГО: матчей — {total_games}, ошибок — {len(all_problems)}")
    say("=" * 78)

    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L) + "\n")
    return len(all_problems), 0


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
        print(f"  {os.path.basename(path)}: матчей {len(games)}, ошибок {len(problems)}")

    e, _ = render(all_summaries, all_problems, OUT)
    print(f"\nвсего ошибок: {e}")
    print(f"отчёт: {OUT}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
