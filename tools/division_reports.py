# -*- coding: utf-8 -*-
"""Отдельный файл с ошибками на каждый дивизион.

У каждого дивизиона свой ответственный, поэтому общий отчёт режется по
дивизионам: в файле только его ошибки, без сводок и без вопросов.

Запуск:  python3 tools/division_reports.py
"""
import datetime
import os
import re
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from error_report import collect, period_sums, SRC_DIR, MONTHS, ERROR

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "reports", "po_divizionam")


def file_name(comp):
    """«Конференция ЦЕНТР (1 Дивизион)» -> «Ошибки_ЦЕНТР_1-дивизион.txt»"""
    conf = "БЕЗ-КОНФЕРЕНЦИИ"
    m = re.search(r"Конференция\s+(\S+)", comp)
    if m:
        conf = m.group(1).upper()
    d = re.search(r"(\d+)\s*Дивизион", comp, re.IGNORECASE)
    div = f"{d.group(1)}-дивизион" if d else "дивизион"
    return f"Ошибки_{conf}_{div}.txt"


def main():
    files = [os.path.join(SRC_DIR, f) for f in sorted(os.listdir(SRC_DIR))
             if f.lower().endswith((".xlsx", ".xlsm", ".xls"))]

    # соревнование -> месяц -> (число матчей, список ошибок)
    by_comp = defaultdict(lambda: defaultdict(lambda: {"games": 0, "errors": [], "restored": []}))
    for path in files:
        games, problems, s = collect(path)
        errors = [p for p in problems if p["kind"] == ERROR]
        for (comp, y, m), n in s["breakdown"].items():
            by_comp[comp][(y, m)]["games"] += n
        # если в файле один месяц — все ошибки относятся к нему
        months = sorted(s["breakdown"], key=lambda k: (k[1], k[2]))
        for p in errors:
            g = p["game"]
            if g is not None and g["date"]:
                key = (g["date"].year, g["date"].month)
                comp = g["comp"] or months[0][0]
            else:
                comp, y, m = months[0]
                key = (y, m)
            by_comp[comp][key]["errors"].append(p)
        for r in s.get("repaired", []):
            g = r["game"]
            if g["date"]:
                key, comp = (g["date"].year, g["date"].month), g["comp"] or months[0][0]
            else:
                comp, y, m = months[0]
                key = (y, m)
            by_comp[comp][key]["restored"].append(r)

    os.makedirs(OUT, exist_ok=True)
    for old in os.listdir(OUT):
        os.remove(os.path.join(OUT, old))

    today = datetime.date.today().strftime("%d.%m.%Y")
    made = []
    for comp in sorted(by_comp):
        L = []
        say = L.append
        total = sum(v["games"] for v in by_comp[comp].values())
        total_err = sum(len(v["errors"]) for v in by_comp[comp].values())

        say("=" * 78)
        say(comp)
        say("=" * 78)
        say(f"Проверка данных по матчам.   Дата проверки: {today}")
        total_fix = sum(len(v["restored"]) for v in by_comp[comp].values())
        say(f"Проверено матчей: {total}.   Найдено ошибок: {total_err}."
            + (f"   Восстановлено по данным: {total_fix}." if total_fix else ""))
        say("")
        say("СВОДКА ПО ДИВИЗИОНУ")
        say("   Месяц            Матчей   Ошибок")
        say("   " + "-" * 34)
        for (y, m) in sorted(by_comp[comp]):
            block = by_comp[comp][(y, m)]
            label = f"{MONTHS[m]} {y}"
            say(f"   {label:<16} {block['games']:>6} {len(block['errors']):>8}")
        say("   " + "-" * 34)
        say(f"   {'всего':<16} {total:>6} {total_err:>8}")
        say("")
        say("Если количество матчей не совпадает с вашим, сообщите — значит, в")
        say("исходной таблице не все игры или есть лишние.")
        say("")
        say("Ниже — места, где данные не сходятся. Указан номер строки в исходной")
        say("таблице, дата и время матча, соперники и что именно не так.")

        for (y, m) in sorted(by_comp[comp]):
            block = by_comp[comp][(y, m)]
            say("")
            say("-" * 78)
            say(f"{MONTHS[m].upper()} {y} — матчей {block['games']}, ошибок {len(block['errors'])}")
            say("-" * 78)
            if block["restored"]:
                say("Восстановлено по данным самого файла, в ошибки не включено:")
                for r in block["restored"]:
                    g = r["game"]
                    d = g["date"].strftime("%d.%m.%Y") if g["date"] else "—"
                    say(f"   Строка {r['row']}   {d}  {g['time'] or '—'}   "
                        f"{g['team1'] or '—'} — {g['team2'] or '—'}")
                    say(f"      {r['what']}: {r['value']}. Основание: {r['why']}.")
                say("")
            if not block["errors"]:
                say("Ошибок не найдено.")
                continue
            for i, p in enumerate(block["errors"], 1):
                g = p["game"]
                say("")
                if g is None:
                    rows = p.get("rows") or []
                    where = (("Строка " if len(rows) == 1 else "Строки ")
                             + ", ".join(str(r) for r in rows[:12])
                             + (f" и ещё {len(rows) - 12}" if len(rows) > 12 else "")) if rows else ""
                    say(f"{i}. {where}" if where else f"{i}. {p['what']}")
                    say(f"   {p['text']}")
                    continue
                d = g["date"].strftime("%d.%m.%Y") if g["date"] else "—"
                say(f"{i}. Строка {g['row']}   {d}  {g['time'] or '—'}   "
                    f"{g['team1'] or '—'} — {g['team2'] or '—'}")
                if g["final_raw"]:
                    mark = f" ({' '.join(g['final_notes'])})" if g["final_notes"] else ""
                    line = f"   Итоговый счёт: {g['final_raw']}{mark}"
                    if g["periods_raw"]:
                        sh, sa = period_sums(g)
                        line += f"   По периодам: {' '.join(g['periods_raw'])}"
                        if sh is not None:
                            line += f" = {sh}:{sa}"
                    say(line)
                say(f"   {p['text']}")

        say("")
        say("=" * 78)
        name = file_name(comp)
        with open(os.path.join(OUT, name), "w", encoding="utf-8") as f:
            f.write("\n".join(L) + "\n")
        made.append((name, total, total_err))

    print(f"файлов: {len(made)}  (папка {OUT})")
    for name, g, e in made:
        print(f"   {name}: матчей {g}, ошибок {e}")


if __name__ == "__main__":
    main()
