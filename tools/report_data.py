# -*- coding: utf-8 -*-
"""Готовит данные отчёта в JSON для вёрстки Word-документа.

Ошибки группируются по дивизионам (а не по файлам): у каждого дивизиона
свой ответственный, и в документе ему отводится отдельная страница.
"""
import datetime, json, os, re, sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from error_report import collect, period_sums, SRC_DIR, MONTHS, ERROR, QUESTION

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "reports", "report_data.json")

# Как спрашивать по каждому виду вопроса
ASK = {
    "игроки выходят за разные команды":
        "Допускается ли, что игрок в течение месяца выступает за несколько команд? "
        "Если это предусмотрено правилами соревнования — вопрос снимается. "
        "Если нет — это ошибка в составах, и её нужно исправить.",
    "игроки с разными номерами":
        "Допускается ли, что у игрока в разных матчах разные игровые номера? "
        "Номер меняется вместе с командой. Если это предусмотрено правилами — "
        "вопрос снимается, если нет — это ошибка в составах.",
}
DEFAULT_ASK = "Это нормальная практика или ошибка в данных? Просим подтвердить."


def finding(p):
    g = p["game"]
    item = dict(what=p["what"], text=p["text"], row="", match="", score="")
    if g is not None:
        item["row"] = str(g["row"])
        d = g["date"].strftime("%d.%m.%Y") if g["date"] else "—"
        item["match"] = f"{d}  {g['time'] or '—'}\n{g['team1'] or '—'} — {g['team2'] or '—'}"
        if g["final_raw"]:
            mark = f" ({' '.join(g['final_notes'])})" if g["final_notes"] else ""
            line = f"итог {g['final_raw']}{mark}"
            if g["periods_raw"]:
                sh, sa = period_sums(g)
                line += f"\nпериоды {' '.join(g['periods_raw'])}"
                if sh is not None:
                    line += f" = {sh}:{sa}"
            item["score"] = line
    return item


def main():
    files = [os.path.join(SRC_DIR, f) for f in sorted(os.listdir(SRC_DIR))
             if f.lower().endswith((".xlsx", ".xlsm", ".xls"))]

    divisions = defaultdict(lambda: defaultdict(lambda: {"games": 0, "errors": []}))
    questions = defaultdict(list)
    n_files = 0

    for path in files:
        games, problems, s = collect(path)
        n_files += 1
        months = sorted(s["breakdown"], key=lambda k: (k[1], k[2]))
        for (comp, y, m), n in s["breakdown"].items():
            divisions[comp][(y, m)]["games"] += n
        for p in problems:
            g = p["game"]
            if g is not None and g["date"]:
                comp, key = g["comp"] or months[0][0], (g["date"].year, g["date"].month)
            else:
                comp, y, m = months[0]
                key = (y, m)
            if p["kind"] == ERROR:
                divisions[comp][key]["errors"].append(p)
            else:
                questions[p["what"]].append(dict(
                    comp=comp, period=f"{MONTHS[key[1]]} {key[0]}", text=p["text"]))

    div_blocks, summary = [], []
    total_games = total_errors = 0
    for comp in sorted(divisions):
        months = []
        g_sum = e_sum = 0
        for (y, m) in sorted(divisions[comp]):
            b = divisions[comp][(y, m)]
            months.append(dict(label=f"{MONTHS[m]} {y}", games=b["games"],
                               errors=len(b["errors"]),
                               findings=[finding(p) for p in b["errors"]]))
            g_sum += b["games"]
            e_sum += len(b["errors"])
            summary.append(dict(comp=comp, month=f"{MONTHS[m]} {y}",
                                games=b["games"], errors=len(b["errors"])))
        div_blocks.append(dict(comp=comp, games=g_sum, errors=e_sum, months=months))
        total_games += g_sum
        total_errors += e_sum

    q_blocks = [dict(what=what, ask=ASK.get(what, DEFAULT_ASK), items=items)
                for what, items in questions.items()]

    data = dict(
        date=datetime.date.today().strftime("%d.%m.%Y"),
        files=n_files, games=total_games, errors=total_errors,
        questions_total=sum(len(q["items"]) for q in q_blocks),
        questions=q_blocks, summary=summary, divisions=div_blocks,
    )
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    print(f"данные: файлов {data['files']}, дивизионов {len(div_blocks)}, "
          f"матчей {data['games']}, ошибок {data['errors']}, "
          f"вопросов {data['questions_total']}")


if __name__ == "__main__":
    main()
