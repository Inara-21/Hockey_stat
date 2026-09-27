# -*- coding: utf-8 -*-
"""Протоколы игровых дней: исходная таблица -> PDF.

Данные переносятся как есть. Ничего не пересчитывается и не исправляется:
спорные места выносятся в отдельный отчёт «Вопросы и ошибки».

Запуск:  python3 tools/build_reports.py data/new/div1.xlsx
"""
import base64, html, os, subprocess, sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from hockey_data import read_workbook, parse_score, score_text, winner

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "reports")
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"
LOGO = os.path.join(ROOT, "assets", "logo_rhl.png")

LEAGUE_TITLE = "Чемпионат Регулярной хоккейной лиги 3х3"
CARDS_PER_SHEET = 4
MONTHS = {1: "январь", 2: "февраль", 3: "март", 4: "апрель", 5: "май", 6: "июнь",
          7: "июль", 8: "август", 9: "сентябрь", 10: "октябрь", 11: "ноябрь", 12: "декабрь"}

CSS = '''
* { box-sizing: border-box; }
@page { size: A4; margin: 21mm 14mm 14mm 14mm; }
body { font-family:"DejaVu Sans","Liberation Sans",sans-serif; color:#1a1f36; margin:0; font-size:9.5px; }
.sheet { page-break-after: always; }
.sheet:last-child { page-break-after: auto; }
.head { display:block; margin-bottom:7mm; }
.head .league { font-weight:800; font-size:19px; letter-spacing:1.1px; text-transform:uppercase;
                color:#2b2f77; text-align:center; line-height:1.15; margin-bottom:4.5mm; }
.head .subhead { display:flex; align-items:center; }
.head .subhead .side { flex:1 1 0; display:flex; align-items:center; }
.head .subhead .side.right { justify-content:flex-end; }
.head img.logo { height:13mm; width:auto; display:block; }
.head .title { flex:0 0 auto; font-weight:800; font-size:14px; letter-spacing:1px; }
.head .date { font-weight:800; font-size:14px; white-space:nowrap; }
.grid { display:grid; grid-template-columns:1fr 1fr; gap:7mm; }
.card { border:1.3px solid #3a3f6b; border-radius:9px; padding:4.5mm 5mm; }
.fld { display:flex; align-items:baseline; gap:5px; border-bottom:1px solid #b9bdd6;
       padding:2px 2px; min-height:14px; }
.fld.wide { margin-bottom:3mm; }
.fld.wide2 { margin-top:2.2mm; }
.row2 { display:grid; grid-template-columns:1fr 1fr; gap:5mm; margin-bottom:2.2mm; }
.lb { font-weight:800; font-size:8px; letter-spacing:.4px; white-space:nowrap; color:#2b2f4c; }
.vl { font-style:italic; color:#1a1f36; }
.mark { font-style:normal; font-weight:800; font-size:8px; letter-spacing:.5px;
        color:#2b2f77; margin-left:4px; white-space:nowrap; }
.rosters { display:grid; grid-template-columns:1fr 1fr; gap:5mm; margin:2.6mm 0; }
table.ros { width:100%; border-collapse:collapse; }
table.ros th, table.ros td { border:1px solid #b9bdd6; padding:1.8px 4px; font-size:8.3px; }
table.ros th { font-weight:800; font-size:7.5px; background:#f2f3f9; letter-spacing:.3px; }
table.ros td.num, table.ros th.num { width:22px; text-align:center; }
table.ros td.pl { font-style:italic; }
.officials { border:1.3px solid #3a3f6b; border-radius:9px; padding:5mm 6mm; margin-top:6mm;
             display:grid; grid-template-columns:1fr 1fr; gap:10mm; break-inside:avoid; }
.off-h { font-weight:800; font-size:9px; letter-spacing:.5px; margin-bottom:3.5mm; }
.off-v { font-style:italic; border-bottom:1px solid #b9bdd6; padding-bottom:3px; }
'''


def e(x):
    return html.escape(str(x))


def logo_b64():
    with open(LOGO, "rb") as f:
        return base64.b64encode(f.read()).decode()


def roster_table(roster, slots):
    rows = ""
    for i in range(slots):
        num, name = roster[i] if i < len(roster) else ("", "")
        rows += f'<tr><td class="num">{e(num)}</td><td class="pl">{e(name)}</td></tr>'
    return ('<table class="ros"><tr><th class="num">№</th><th class="pl">ИГРОК</th></tr>'
            + rows + '</table>')


def card_html(g, slots):
    mark = '<span class="mark">БУЛЛИТЫ</span>' if g["shootout"] else ""
    periods = ", ".join(score_text(p) for p in g["periods_raw"])
    return f'''<div class="card">
  <div class="fld wide"><span class="lb">СОРЕВНОВАНИЕ</span><span class="vl">{e(g["comp"])}</span></div>
  <div class="row2">
    <div class="fld"><span class="lb">ДАТА</span><span class="vl">{e(g["date"].strftime("%d.%m.%Y"))}</span></div>
    <div class="fld"><span class="lb">ВРЕМЯ</span><span class="vl">{e(g["time"])}</span></div>
  </div>
  <div class="row2">
    <div class="fld"><span class="lb">КОМАНДА&nbsp;А</span><span class="vl">{e(g["team1"])}</span></div>
    <div class="fld"><span class="lb">КОМАНДА&nbsp;В</span><span class="vl">{e(g["team2"])}</span></div>
  </div>
  <div class="rosters">{roster_table(g["roster1"], slots)}{roster_table(g["roster2"], slots)}</div>
  <div class="row2">
    <div class="fld"><span class="lb">ИТОГОВЫЙ&nbsp;СЧЁТ</span>
        <span class="vl">{e(score_text(g["final_raw"]))}</span>{mark}</div>
    <div class="fld"><span class="lb">ПОБЕДИТЕЛЬ</span><span class="vl">{e(winner(g))}</span></div>
  </div>
  <div class="fld wide2"><span class="lb">СЧЁТ&nbsp;ПО&nbsp;ПЕРИОДАМ</span><span class="vl">{e(periods)}</span></div>
</div>'''


def officials_html(judge, secretary):
    return f'''<div class="officials">
  <div><div class="off-h">СУДЬИ</div><div class="off-v">{e(judge)}</div></div>
  <div><div class="off-h">СЕКРЕТАРИ</div><div class="off-v">{e(secretary)}</div></div>
</div>'''


def day_html(day_games, slots, logo):
    date = day_games[0]["date"].strftime("%d.%m.%Y")
    judge, secretary = day_games[0]["judge"], day_games[0]["secretary"]
    chunks = [day_games[i:i + CARDS_PER_SHEET]
              for i in range(0, len(day_games), CARDS_PER_SHEET)]
    out = ""
    for i, chunk in enumerate(chunks):
        label = e(date) + (f" · лист {i+1} из {len(chunks)}" if len(chunks) > 1 else "")
        off = officials_html(judge, secretary) if i == len(chunks) - 1 else ""
        out += (f'<div class="sheet"><div class="head">'
                f'<div class="league">{e(LEAGUE_TITLE)}</div>'
                f'<div class="subhead">'
                f'<div class="side left"><img class="logo" src="data:image/png;base64,{logo}" alt="РХЛ"></div>'
                f'<div class="title">ПРОТОКОЛ ИГРОВОГО ДНЯ</div>'
                f'<div class="side right"><div class="date">{label}</div></div>'
                f'</div></div>'
                f'<div class="grid">{"".join(card_html(g, slots) for g in chunk)}</div>{off}</div>')
    return out, len(chunks)


def render_pdf(body, pdf_path, tmp_html):
    page = ('<!doctype html><html><head><meta charset="utf-8">'
            f'<style>{CSS}</style></head><body>{body}</body></html>')
    with open(tmp_html, "w", encoding="utf-8") as f:
        f.write(page)
    subprocess.run([CHROME, "--headless", "--no-sandbox", "--disable-gpu",
                    "--no-pdf-header-footer", f"--print-to-pdf={pdf_path}",
                    f"file://{tmp_html}"], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


# ------------------------------------------------------- вопросы и ошибки
def problems_report(games, path):
    lines = ["=" * 78,
             "ВОПРОСЫ И ОШИБКИ ПО ИСХОДНЫМ ДАННЫМ",
             "=" * 78,
             "",
             "Данные в протоколы перенесены ровно так, как записаны в исходной таблице.",
             "Ничего не исправлялось. Ниже — места, которые требуют вашей проверки.",
             ""]
    n = 0
    for g in games:
        if g["goals1"] is None or any(parse_score(p)[0] is None for p in g["periods_raw"]):
            continue
        sh = sum(parse_score(p)[0] for p in g["periods_raw"])
        sa = sum(parse_score(p)[1] for p in g["periods_raw"])
        if (sh, sa) == (g["goals1"], g["goals2"]) or g["shootout"] or g["overtime"]:
            continue
        n += 1
        tie_plus_one = sh == sa and (g["goals1"] - sh) + (g["goals2"] - sa) == 1
        lines += ["-" * 78,
                  f"{n}. Соревнование: {g['comp']}",
                  f"   Строка в таблице: {g['row']}",
                  f"   Дата: {g['date'].strftime('%d.%m.%Y')}   Время: {g['time']}",
                  f"   Матч: {g['team1']} — {g['team2']}",
                  f"   Итоговый счёт: {g['final_raw']}",
                  f"   Счёт по периодам: {'   '.join(g['periods_raw'])}",
                  f"   Сумма по периодам: {sh}:{sa}",
                  f"   Не сходится: по периодам {sh}:{sa}, а в итоге {g['goals1']}:{g['goals2']}"
                  f" (разница {g['goals1'] - sh}:{g['goals2'] - sa})",
                  ""]
        if tie_plus_one:
            lines += ["   ВОПРОС: после трёх периодов ничья, а в итоге на один гол больше.",
                      "   Так записаны матчи, выигранные по буллитам, но пометки «Б Буллиты»",
                      "   у этого матча нет. Это победа по буллитам и пометку пропустили,",
                      "   или в счёте опечатка? Просьба уточнить.", ""]
        else:
            lines += ["   ОШИБКА: счёт не сходится со счётом по периодам.",
                      "   Ничьей после трёх периодов не было, поэтому буллитами это",
                      "   не объясняется: в итоге есть гол, которого нет ни в одном периоде.",
                      "   Просьба сверить с протоколом матча.", ""]
    lines += ["=" * 78,
              f"Всего мест, требующих проверки: {n}",
              "=" * 78]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return n


def main(src):
    games, _ = read_workbook(src)
    games = [g for g in games if g["date"] and g["time"]]
    slots = max(max(len(g["roster1"]), len(g["roster2"])) for g in games)
    print(f"матчей: {len(games)}   строк в таблице состава: {slots}")

    os.makedirs(OUT, exist_ok=True)
    os.makedirs(os.path.join(OUT, "po_dnyam"), exist_ok=True)
    logo = logo_b64()
    tmp = os.path.join(OUT, "_tmp.html")

    byday = defaultdict(list)
    for g in games:
        byday[g["date"]].append(g)
    for d in byday:
        byday[d].sort(key=lambda x: x["time"])

    bymonth = defaultdict(str)
    order, sheets_of = [], defaultdict(int)
    total_sheets = 0
    for d in sorted(byday):
        body, n = day_html(byday[d], slots, logo)
        total_sheets += n
        key = (d.year, d.month)
        if key not in bymonth:
            order.append(key)
        bymonth[key] += body
        sheets_of[key] += n
        render_pdf(body, os.path.join(OUT, "po_dnyam", f"{d}.pdf"), tmp)

    for key in order:
        y, m = key
        name = f"Протоколы_{MONTHS[m]}_{y}.pdf"
        render_pdf(bymonth[key], os.path.join(OUT, name), tmp)
        cnt = sum(len(byday[d]) for d in byday if (d.year, d.month) == key)
        days = sum(1 for d in byday if (d.year, d.month) == key)
        print(f"   {name}: матчей {cnt}, дней {days}, листов {sheets_of[key]}")

    if os.path.exists(tmp):
        os.remove(tmp)
    n = problems_report(games, os.path.join(OUT, "Voprosy_i_oshibki.txt"))
    print(f"дней: {len(byday)}   листов всего: {total_sheets}")
    print(f"вопросов и ошибок вынесено: {n}")
    return games


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "data", "new", "div1.xlsx"))
