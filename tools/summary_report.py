# -*- coding: utf-8 -*-
"""Сводный отчёт об ошибках по всем дивизионам и месяцам.

Берёт все файлы из data/source, накладывает правки проверяющих из
data/corrections и готовит JSON для вёрстки Word:
  итог по каждому файлу; по каждому дивизиону с ошибками — составы команд,
  где игрок записан за несколько команд (с пометками: за кого ещё, нет
  номера, повтор номера), и все остальные ошибки по месяцам; в конце —
  количество игр для сверки.

Запуск:  python3 tools/summary_report.py
"""
import datetime
import json
import os
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import corrections
from error_report import collect, plural, ERROR, MONTHS, SRC_DIR
from report_data import finding

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "reports", "svodny_data.json")
ROSTER_KIND = "игрок за несколько команд"


def roster_flags(r, players, other_teams):
    """Строки состава с пометками о том, что с игроком не так."""
    nums = Counter(n for n, _ in r if n)
    out = []
    for n, nm in r:
        notes = []
        if nm in players and other_teams.get(nm):
            notes.append("записан также за " + ", ".join(f"{t} (№{x})" if x else t for t, x in other_teams[nm]))
        if nm and not n:
            notes.append("нет номера")
        if n and nums[n] > 1:
            others = [nm2 for n2, nm2 in r if n2 == n and nm2 != nm]
            notes.append(f"номер {n} также у: {', '.join(others)}")
        out.append(dict(text=f"№{n} {nm}" if n else nm, flag=bool(notes), note="; ".join(notes)))
    return out


def main():
    files, by_comp = [], defaultdict(list)
    fixed_files = {fx["file"] for fx in corrections.FIXES}
    for f in sorted(os.listdir(SRC_DIR)):
        if not f.endswith(".xlsx"):
            continue
        games, problems, s = collect(os.path.join(SRC_DIR, f))
        (comp, y, m), = s["breakdown"].keys()
        errors = [p for p in problems if p["kind"] == ERROR]
        item = dict(file=f, comp=comp, ym=(y, m), month=f"{MONTHS[m]} {y}", games=games,
                    errors=errors, misplaced=len(s["misplaced"]), answered=f in fixed_files)
        files.append(item)
        by_comp[comp].append(item)

    divisions = []
    for comp in sorted(by_comp):
        items = sorted(by_comp[comp], key=lambda x: x["ym"])
        if not any(x["errors"] for x in items):
            continue
        # Составы — только за месяцы, где есть игроки, записанные за несколько команд.
        roster_items = [x for x in items if any(p["what"] == ROSTER_KIND for p in x["errors"])]
        players = {nm for x in roster_items for p in x["errors"] if p["what"] == ROSTER_KIND
                   for nm in p.get("players", [])}
        variants = defaultdict(Counter)              # команда -> (состав, месяц) -> матчей
        where = defaultdict(set)                     # игрок -> {(команда, номер)}
        for x in roster_items:
            for g in x["games"]:
                for t, r in ((g["team1"], g["roster1"]), (g["team2"], g["roster2"])):
                    variants[t][(tuple(r), x["month"])] += 1
                    for n, nm in r:
                        if nm in players:
                            where[nm].add((t, n))
        teams = []
        for nm in sorted(players):
            for t, _ in sorted(where[nm]):
                if t not in [tm["team"] for tm in teams]:
                    teams.append(dict(team=t))
        months_r = [x["month"] for x in roster_items]
        for tm in teams:
            t = tm["team"]
            by_roster = defaultdict(dict)
            for (r, mo), n in variants[t].items():
                by_roster[r][mo] = n
            other = {nm: sorted((t2, n2) for t2, n2 in where[nm] if t2 != t) for nm in players}
            shown = []
            for r, per_month in sorted(by_roster.items(), key=lambda kv: -sum(kv[1].values())):
                if len(by_roster) == 1 and set(per_month) == set(months_r):
                    label = "Во всех матчах" + (f" ({' и '.join(months_r)})" if len(months_r) > 1 else f" ({months_r[0]})")
                else:
                    label = "; ".join(f"{mo} — {n} {plural(n, 'матч', 'матча', 'матчей')}" for mo, n in per_month.items())
                shown.append(dict(label=label, players=roster_flags(list(r), players, other)))
            tm["rosters"] = shown
            tm["same_everywhere"] = len(by_roster) == 1

        others = []
        for x in items:
            rows = [finding(p) for p in x["errors"] if p["what"] != ROSTER_KIND]
            others.append(dict(month=x["month"], findings=rows))
        divisions.append(dict(comp=comp, teams=teams, others=others,
                              answered=any(x["answered"] for x in items)))

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
        corrections_source=corrections.SOURCES,
        corrections_date=corrections.LAST_DATE,
        files=[dict(file=x["file"], comp=x["comp"], month=x["month"], clean=not x["errors"],
                    answered=x["answered"]) for x in sorted(files, key=lambda x: (x["comp"], x["ym"]))],
        divisions=divisions,
        months=month_labels,
        games_table=table,
        games_total={mo: sum(len(x["games"]) for x in files if x["month"] == mo) for mo in month_labels},
        games_grand=sum(len(x["games"]) for x in files),
        misplaced_total=sum(x["misplaced"] for x in files),
        comps=sorted(by_comp),
        divisions_all=[dict(comp=c, answered=any(x["answered"] for x in by_comp[c])) for c in sorted(by_comp)],
    )
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    print(f"файлов {len(files)}, без ошибок {sum(1 for x in files if not x['errors'])}, "
          f"дивизионов с ошибками {len(divisions)}, матчей {data['games_grand']}")


if __name__ == "__main__":
    main()
