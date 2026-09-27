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
SRC_DIR = os.path.join(ROOT, "data", "new")
OUT = os.path.join(ROOT, "reports", "Otchet_ob_oshibkah.txt")

REQUIRED = ["comp", "date", "time", "team1", "team2", "roster1", "roster2",
            "secretary", "judge", "final", "periods"]
LABEL_RU = {"comp": "Название соревнования", "date": "Дата", "time": "Время начала матча",
            "team1": "Команда 1", "team2": "Команда 2", "roster1": "Состав команды 1",
            "roster2": "Состав команды 2", "secretary": "Секретари", "judge": "Судьи",
            "final": "Итоговый счёт матча", "periods": "Счёт по периодам"}

ERROR, QUESTION = "ОШИБКА", "ВОПРОС"


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
                f"Дату не удалось прочитать (записана как «{g['date_kind']}»).")
        if g["time"] is None:
            add(ERROR, g, "время не разобрано",
                f"Время не удалось прочитать (записано как «{g['time_kind']}»).")
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
            add(QUESTION, g, "нет судьи", "Судья не указан. Просьба дополнить.")
        if not g["secretary"]:
            add(QUESTION, g, "нет секретаря", "Секретарь не указан. Просьба дополнить.")

        if not g["final_raw"]:
            add(ERROR, g, "пустой итоговый счёт", "Итоговый счёт не заполнен.")
        elif g["goals1"] is None:
            add(ERROR, g, "итог не разобран",
                f"Итоговый счёт «{g['final_raw']}» записан так, что его не удалось прочитать.")
        if not g["periods_raw"]:
            add(ERROR, g, "пустые периоды", "Счёт по периодам не заполнен.")
        for p in g["periods_raw"]:
            if parse_score(p)[0] is None:
                add(ERROR, g, "период не разобран",
                    f"Счёт периода «{p}» записан так, что его не удалось прочитать.")

        # расхождение периодов и итога
        sh, sa = period_sums(g)
        if g["goals1"] is not None and sh is not None and (sh, sa) != (g["goals1"], g["goals2"]):
            tie_plus_one = sh == sa and (g["goals1"] - sh) + (g["goals2"] - sa) == 1
            if g["shootout"] or g["overtime"]:
                if not tie_plus_one:
                    add(QUESTION, g, "пометка не соответствует счёту",
                        f"Стоит пометка «{' '.join(g['final_notes'])}», но после основного "
                        f"времени ничьей не было: по периодам {sh}:{sa}, в итоге "
                        f"{g['goals1']}:{g['goals2']}. Просьба уточнить.")
            elif tie_plus_one:
                add(QUESTION, g, "похоже на буллиты, но пометки нет",
                    "После трёх периодов ничья, а в итоге на один гол больше. Так записаны "
                    "матчи, выигранные по буллитам, но пометки «Б Буллиты» у этого матча нет.\n"
                    "   Это победа по буллитам и пометку пропустили, или в счёте опечатка? "
                    "Просьба уточнить.")
            else:
                add(ERROR, g, "счёт не сходится с периодами",
                    "Счёт не сходится со счётом по периодам. Ничьей после трёх периодов не "
                    "было, поэтому буллитами это не объясняется: в итоге есть гол, которого "
                    "нет ни в одном периоде.\n   Просьба сверить с протоколом матча.")

        for who, roster in (("А", g["roster1"]), ("В", g["roster2"])):
            nums = [n for n, _ in roster]
            for n, c in Counter(nums).items():
                if c > 1 and n:
                    who_n = [nm for nu, nm in roster if nu == n]
                    add(ERROR, g, "повтор игрового номера",
                        f"В составе команды {who} номер №{n} у двух игроков: {', '.join(who_n)}.")
            for n, nm in roster:
                if not nm:
                    add(ERROR, g, "игрок без ФИО",
                        f"В составе команды {who} есть номер №{n} без фамилии.")
                elif not n:
                    add(QUESTION, g, "игрок без номера",
                        f"В составе команды {who} у игрока «{nm}» не указан номер.")

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
    for nm, teams in ptm.items():
        if len(teams) > 1:
            add(QUESTION, None, "игрок в разных командах",
                f"«{nm}» указан в составах разных команд: "
                f"{', '.join(f'{t} ({c} матчей)' for t, c in teams.items())}. Просьба уточнить.")
    for nm, nums in pnum.items():
        if len(nums) > 1:
            add(QUESTION, None, "игрок с разными номерами",
                f"У «{nm}» в разных матчах разные номера: "
                f"{', '.join(f'№{n} ({c} раз)' for n, c in nums.items())}. Просьба уточнить.")

    names = sorted({g["team1"] for g in games} | {g["team2"] for g in games})
    for i, a in enumerate(names):
        for b in names[i + 1:]:
            if difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio() > 0.85:
                add(QUESTION, None, "похожие названия команд",
                    f"Названия «{a}» и «{b}» очень похожи — возможна опечатка.")

    byday = Counter(g["date"] for g in games if g["date"])
    if byday:
        avg = sum(byday.values()) / len(byday)
        for d in sorted(byday):
            if byday[d] < avg / 2:
                add(QUESTION, None, "мало матчей в день",
                    f"{d.strftime('%d.%m.%Y')}: всего {byday[d]} матч(а), "
                    f"в среднем по другим дням {avg:.1f}. Возможно, данные неполные.")
        days = sorted(byday)
        for n in range((days[-1] - days[0]).days + 1):
            d = days[0] + datetime.timedelta(n)
            if d not in byday:
                add(QUESTION, None, "день без матчей",
                    f"{d.strftime('%d.%m.%Y')}: матчей нет, хотя это середина периода.")

    wb = openpyxl.load_workbook(path)
    summary = dict(
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
    L = []
    say = L.append
    today = datetime.date.today().strftime("%d.%m.%Y")

    say("=" * 78)
    say("ОТЧЁТ О ПРОВЕРКЕ ИСХОДНЫХ ДАННЫХ")
    say(f"Дата проверки: {today}")
    say("=" * 78)
    say("")
    say("Данные проверены без внесения изменений. Ниже перечислены места,")
    say("которые требуют проверки на стороне заказчика.")
    say("")

    say("-" * 78)
    say("ПРОВЕРЕННЫЕ ФАЙЛЫ")
    say("-" * 78)
    for s in all_summaries:
        say(f"\n  {s['file']}")
        say(f"     матчей: {s['games']}   команд: {s['teams']}   игроков: {s['players']}")
        say(f"     соревнования: {', '.join(s['comps']) if s['comps'] else '—'}")
        if s["dates"]:
            say(f"     период: {s['dates'][0].strftime('%d.%m.%Y')} — "
                f"{s['dates'][1].strftime('%d.%m.%Y')}, дней с матчами: {s['days']}")
        say(f"     составы (А, В): {s['rosters']}")
        say(f"     победы по буллитам: {s['shootouts']}")
        if s["skipped"]:
            say(f"     листы без матчей (пропущены): "
                f"{', '.join(f'{t} ({r} строк)' for t, r in s['skipped'])}")
        if s["hidden_rows"] or s["formulas"]:
            say(f"     скрытых строк: {s['hidden_rows']}, формул: {s['formulas']}")

    errors = [p for p in all_problems if p["kind"] == ERROR]
    questions = [p for p in all_problems if p["kind"] == QUESTION]

    say("")
    say("-" * 78)
    say("СВОДКА")
    say("-" * 78)
    say(f"\n  Проверено файлов: {len(all_summaries)}")
    say(f"  Проверено матчей: {sum(s['games'] for s in all_summaries)}")
    say(f"  Найдено ошибок: {len(errors)}")
    say(f"  Требуют уточнения (вопросы): {len(questions)}")
    if all_problems:
        say("\n  По видам:")
        for what, c in Counter(p["what"] for p in all_problems).most_common():
            say(f"     {c:4}  {what}")

    for title, items in (("ОШИБКИ", errors), ("ВОПРОСЫ", questions)):
        if not items:
            continue
        say("")
        say("=" * 78)
        say(f"{title} ({len(items)})")
        say("=" * 78)
        for i, p in enumerate(items, 1):
            g = p["game"]
            say("")
            say("-" * 78)
            say(f"{i}. {p['what'].upper()}")
            say(f"   Файл: {p['file']}")
            if g is not None:
                say(f"   Соревнование: {g['comp'] or '—'}")
                say(f"   Строка в таблице: {g['row']}")
                d = g["date"].strftime("%d.%m.%Y") if g["date"] else "—"
                say(f"   Дата: {d}   Время: {g['time'] or '—'}")
                say(f"   Матч: {g['team1'] or '—'} — {g['team2'] or '—'}")
                if g["final_raw"]:
                    mark = f"   ({' '.join(g['final_notes'])})" if g["final_notes"] else ""
                    say(f"   Итоговый счёт: {g['final_raw']}{mark}")
                if g["periods_raw"]:
                    sh, sa = period_sums(g)
                    tail = f"   (сумма {sh}:{sa})" if sh is not None else ""
                    say(f"   Счёт по периодам: {'   '.join(g['periods_raw'])}{tail}")
            say("")
            say(f"   {p['kind']}: {p['text']}")

    say("")
    say("=" * 78)
    say(f"ИТОГО: ошибок — {len(errors)}, вопросов — {len(questions)}")
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
        e = sum(1 for p in problems if p["kind"] == ERROR)
        q = len(problems) - e
        print(f"  {os.path.basename(path)}: матчей {len(games)}, ошибок {e}, вопросов {q}")

    e, q = render(all_summaries, all_problems, OUT)
    print(f"\nвсего: ошибок {e}, вопросов {q}")
    print(f"отчёт: {OUT}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
