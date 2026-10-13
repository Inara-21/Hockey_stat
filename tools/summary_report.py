# -*- coding: utf-8 -*-
"""Сводный отчёт по исправленным файлам (reports/ispravlennye).

Готовит данные в JSON для вёрстки Word: итог по каждому файлу, по каждому
дивизиону с ошибками — составы команд с лишними игроками и остальные ошибки,
в конце — количество игр для сверки.

Запуск:  python3 tools/summary_report.py
"""
import datetime
import glob
import json
import os
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from error_report import collect, ERROR, MONTHS
from report_data import finding

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "reports", "ispravlennye")
OUT = os.path.join(ROOT, "reports", "svodny_data.json")
ROSTER_KIND = "игрок за несколько команд"


def main():
    files, by_comp = [], defaultdict(list)
    for path in sorted(glob.glob(os.path.join(SRC, "*.xlsx"))):
        games, problems, s = collect(path)
        (comp, y, m), = s["breakdown"].keys()
        errors = [p for p in problems if p["kind"] == ERROR]
        item = dict(file=os.path.basename(path), comp=comp, ym=(y, m), month=f"{MONTHS[m]} {y}",
                    games=games, errors=errors, misplaced=len(s["misplaced"]))
        files.append(item)
        by_comp[comp].append(item)

    divisions = []
    for comp in sorted(by_comp):
        items = sorted(by_comp[comp], key=lambda x: x["ym"])
        if not any(x["errors"] for x in items):
            continue
        months = [x["month"] for x in items]

        # Составы команд, в которых есть игрок, записанный за несколько команд.
        players = sorted({nm for x in items for p in x["errors"] if p["what"] == ROSTER_KIND
                          for nm in p.get("players", [])})
        rosters = defaultdict(dict)                 # команда -> месяц -> {состав: матчей}
        for x in items:
            for g in x["games"]:
                for t, r in ((g["team1"], g["roster1"]), (g["team2"], g["roster2"])):
                    rosters[t].setdefault(x["month"], Counter())[tuple(r)] += 1
        teams = []
        for nm in players:
            for t in sorted(rosters):
                if t in [x["team"] for x in teams]:
                    continue
                if any(nm == pn for c in rosters[t].values() for r in c for _, pn in r):
                    teams.append(dict(team=t))
        for tm in teams:
            per_month = rosters[tm["team"]]
            variants = {r for c in per_month.values() for r in c}
            tm["same_everywhere"] = len(variants) == 1 and len(per_month) == len(months)
            shown = []
            for r in sorted(variants, key=lambda r: -sum(c.get(r, 0) for c in per_month.values())):
                where = [mo for mo in months if r in per_month.get(mo, {})]
                lines = []
                for n, nm in r:
                    other = [(t2, n2) for t2 in rosters if t2 != tm["team"]
                             for c in rosters[t2].values() for rr in c for n2, nm2 in rr if nm2 == nm]
                    other = sorted(set(other))
                    lines.append(dict(text=f"№{n} {nm}" if n else nm, flag=bool(other) and nm in players,
                                      also=", ".join(f"{t2} (№{n2})" for t2, n2 in other)))
                shown.append(dict(months=where, players=lines))
            tm["rosters"] = shown

        others = []
        for x in items:
            rows = [dict(finding(p), month=x["month"]) for p in x["errors"] if p["what"] != ROSTER_KIND]
            if rows:
                others.append(dict(month=x["month"], findings=rows))
        divisions.append(dict(comp=comp, teams=teams, others=others,
                              all_same=bool(teams) and all(t["same_everywhere"] for t in teams),
                              months_text=" и ".join(months)))

    month_labels = [f"{MONTHS[m]} {y}" for (y, m) in sorted({x["ym"] for x in files})]
    table = []
    for comp in sorted(by_comp):
        row = dict(comp=comp, cells={}, total=0)
        for x in by_comp[comp]:
            n = len(x["games"])
            row["cells"][x["month"]] = f"{n}" + (f" + {x['misplaced']} под вопросом" if x["misplaced"] else "")
            row["total"] += n
        table.append(row)

    data = dict(
        date=datetime.date.today().strftime("%d.%m.%Y"),
        files=[dict(file=x["file"], comp=x["comp"], month=x["month"], clean=not x["errors"]) for x in
               sorted(files, key=lambda x: (x["comp"], x["ym"]))],
        divisions=divisions,
        months=month_labels,
        games_table=table,
        games_total={mo: sum(len(x["games"]) for x in files if x["month"] == mo) for mo in month_labels},
        games_grand=sum(len(x["games"]) for x in files),
        misplaced_total=sum(x["misplaced"] for x in files),
    )
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    print(f"файлов {len(files)}, без ошибок {sum(1 for x in files if not x['errors'])}, "
          f"дивизионов с ошибками {len(divisions)}, матчей {data['games_grand']}")


if __name__ == "__main__":
    main()
