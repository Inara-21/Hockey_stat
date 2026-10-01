# -*- coding: utf-8 -*-
"""Готовит данные отчёта в JSON для вёрстки Word-документа."""
import json, os, sys, datetime
from collections import Counter
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from error_report import collect, period_sums, SRC_DIR, MONTHS, THEMES, ERROR, QUESTION

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "reports", "report_data.json")


def main():
    files = [os.path.join(SRC_DIR, f) for f in sorted(os.listdir(SRC_DIR))
             if f.lower().endswith((".xlsx", ".xlsm", ".xls"))]
    all_problems, blocks, total = [], [], Counter()

    for path in files:
        games, problems, s = collect(path)
        all_problems += problems
        total.update(s["breakdown"])
        items = []
        for i, p in enumerate([q for q in problems if q["kind"] == ERROR], 1):
            g = p["game"]
            item = dict(n=i, what=p["what"], text=p["text"], row="", match="", score="")
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
            items.append(item)
        qs = [q for q in problems if q["kind"] == QUESTION]
        blocks.append(dict(
            questions=[dict(what=q["what"], text=q["text"]) for q in qs],
            file=s["file"],
            title=", ".join(s["comps"]) if s["comps"] else "Соревнование не указано",
            period=", ".join(f"{MONTHS[m]} {y}" for (c, y, m) in
                             sorted(s["breakdown"], key=lambda k: (k[1], k[2]))),
            games=s["games"], teams=s["teams"], players=s["players"],
            shootouts=s["shootouts"], findings=items))

    errors = [q for q in all_problems if q["kind"] == ERROR]
    questions = [q for q in all_problems if q["kind"] == QUESTION]
    kinds = Counter(p["what"] for p in errors)
    themes, used = [], set()
    for title, whats, comment in THEMES:
        n = sum(kinds.get(w, 0) for w in whats)
        if not n:
            continue
        used.update(whats)
        themes.append(dict(title=title, n=n, comment=comment,
                           detail="; ".join(f"{kinds[w]} — {w}" for w in whats if kinds.get(w))))
    themes.sort(key=lambda t: -t["n"])
    other = {w: c for w, c in kinds.items() if w not in used}
    if other:
        themes.append(dict(title="Прочее", n=sum(other.values()),
                           detail="; ".join(f"{c} — {w}" for w, c in other.items()),
                           comment=""))

    data = dict(
        date=datetime.date.today().strftime("%d.%m.%Y"),
        files=len(files),
        games=sum(b["games"] for b in blocks),
        findings=len(errors),
        questions=len(questions),
        themes=themes,
        summary=[dict(comp=c, month=f"{MONTHS[m]} {y}", games=n)
                 for (c, y, m), n in sorted(total.items(), key=lambda x: (x[0][0], x[0][1], x[0][2]))],
        blocks=blocks,
    )
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    print(f"данные: {OUT}  (файлов {data['files']}, матчей {data['games']}, "
          f"ошибок {data['findings']}, вопросов {data['questions']})")


if __name__ == "__main__":
    main()
