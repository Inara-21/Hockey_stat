# -*- coding: utf-8 -*-
"""Протоколы игровых дней: один дивизион — один месяц — один PDF.

Данные — файлы месяцев из data/source после всех правок реестра, то есть те же,
что в итоговой книге reports/Vse_matchi_s_pravkami.xlsx. Вёрстка — как у
протоколов раннего этапа (tools/build_reports.py): лист А4, 4 карточки матчей
одного игрового дня (2×2). Судья и секретарь — в карточке каждого матча: в
ВОСТОКЕ и ЗАПАДЕ за игровой день их бывает до четырёх разных.

Каждый PDF после печати сверяется со своими данными: листов столько, сколько
должно быть, на каждом листе — его матчи, и в каждой карточке найдены все
значения (дата, время, команды, все игроки с номерами, счёт, периоды, судья,
секретарь).

Запуск:  python3 tools/make_protocols.py                         все файлы
         python3 tools/make_protocols.py 2026-07_vostok_div1.xlsx  только этот
Результат: reports/protokoly/Протоколы_ВОСТОК_1_дивизион_июль_2026.pdf
"""
import html
import os
import re
import subprocess
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import corrections
from build_reports import CSS, CHROME, CARDS_PER_SHEET, LEAGUE_TITLE, MONTHS, e, logo_b64
from hockey_data import parse_score, read_workbook, score_text, winner

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "data", "source")
OUT = os.path.join(ROOT, "reports", "protokoly")
CONF = {"vostok": "ВОСТОК", "zapad": "ЗАПАД", "centr": "ЦЕНТР"}

CSS_EXTRA = """
.sheet { padding-top:1.5mm; }   /* без отступа значок над Й в заголовке печатается внизу предыдущей страницы */
/* масштаб 90%: так всё помещается по ширине; без явного масштаба браузер сам ужимает лист,
   если что-то шире страницы, и результат зависит от данных */
body { zoom:0.9; }
.grid { grid-template-columns:minmax(0,1fr) minmax(0,1fr); }
/* лист — ровно страница (A4 без полей: 297 − 21 − 14 = 262 мм, с учётом масштаба),
   номер листа — в правом нижнем углу */
.sheet { position:relative; height:290mm; }
.foot { position:absolute; right:0; bottom:0; font-size:12px; font-weight:400; color:#000; }
.head { margin-bottom:8mm; }
.head .top { display:grid; grid-template-columns:16mm 1fr 16mm; align-items:center; margin-bottom:5mm; }
.head .top img.logo { height:13mm; width:auto; display:block; }
.head .league { margin:0; text-transform:none; font-size:17px; white-space:nowrap; }
.io { font-size:7.4px; }
.head .division { text-align:center; font-weight:700; font-size:14px; color:#2b2f4c;
                  letter-spacing:.4px; margin:0 0 5mm 0; }   /* размером как ПРОТОКОЛ ИГРОВОГО ДНЯ */
.grid { gap:5mm; }
.card { padding:3.6mm 4.5mm; break-inside:avoid; }
.fld.wide { margin-bottom:2.2mm; }
.row2 { margin-bottom:1.6mm; }
.rosters { margin:2mm 0; }
.fld.wide2 { margin-top:1.4mm; }
.nowrap { white-space:nowrap; }
table.ros td { vertical-align:top; }
table.ros td.num { vertical-align:middle; text-align:right; padding-right:4px; }
.fld.off { margin-top:1.2mm; }
/* статичная информация — чёрным прямым шрифтом, заполняемая — синим курсивом */
body, .lb, table.ros th, .head .league, .head .title, .head .date { color:#000; font-style:normal; }
.vl, table.ros td, .mark, .head .division, .head .date .dt { color:#1a3a86; font-style:italic; }
.vl.team { font-weight:700; white-space:nowrap; }   /* названия команд — жирным, без переноса по дефису */
.sur { font-weight:400; }       /* фамилии игроков — обычным */
.vl.big { font-weight:700; font-size:12px; white-space:nowrap; }   /* итоговый счёт и победитель */
/* строка результата: победитель — по ширине названия, счёт — остальное; БУЛЛИТЫ при нехватке места — под счётом */
.row2.res { grid-template-columns:minmax(0,1fr) max-content; }
.fld.score { flex-wrap:wrap; }
"""


def shootout_period(g):
    """Буллиты записаны четвёртым периодом (сентябрьские файлы): после трёх периодов ничья."""
    per = [parse_score(p) for p in g["periods_raw"]]
    if len(per) != 4 or any(x[0] is None for x in per) or per[3] not in ((1, 0), (0, 1)):
        return False
    return sum(x[0] for x in per[:3]) == sum(x[1] for x in per[:3])


def periods_text(g):
    """Счёт по периодам для протокола: буллиты четвёртым периодом выводятся отдельно."""
    if shootout_period(g):
        return ", ".join(score_text(p) for p in g["periods_raw"][:3]) + "; буллиты " + score_text(g["periods_raw"][3])
    return ", ".join(score_text(p) for p in g["periods_raw"])


def player_parts(name):
    """ФИО -> (ФАМИЛИЯ, Имя Отчество)."""
    sur, _, rest = name.strip().partition(" ")
    return sur.upper(), rest.strip()


def official(name):
    """Судья, секретарь: ФАМИЛИЯ Имя Отчество."""
    return " ".join(x for x in player_parts(name) if x) if name else ""


def roster_table(roster, slots):
    rows = ""
    for i in range(slots):
        num, name = roster[i] if i < len(roster) else ("", "")
        sur, rest = player_parts(name) if name else ("", "")
        cell = f'<span class="sur">{e(sur)}</span>' + (f'<br><span class="io">{e(rest)}</span>' if rest else "")
        rows += f'<tr><td class="num">{e(num)}</td><td class="pl">{cell}</td></tr>'
    return ('<table class="ros"><tr><th class="num">№</th><th class="pl">ИГРОК</th></tr>'
            + rows + '</table>')


def card_html(g, slots):
    mark = '<span class="mark">БУЛЛИТЫ</span>' if g["shootout"] or shootout_period(g) else ""
    return f'''<div class="card">
  <div class="fld wide"><span class="lb">СОРЕВНОВАНИЕ</span><span class="vl">{e(g["comp"])}</span></div>
  <div class="row2">
    <div class="fld"><span class="lb">ДАТА</span><span class="vl">{e(g["date"].strftime("%d.%m.%Y"))}</span></div>
    <div class="fld"><span class="lb">ВРЕМЯ</span><span class="vl">{e(g["time"])}</span></div>
  </div>
  <div class="row2">
    <div class="fld"><span class="lb">КОМАНДА&nbsp;А</span><span class="vl team">{e(g["team1"].upper())}</span></div>
    <div class="fld"><span class="lb">КОМАНДА&nbsp;В</span><span class="vl team">{e(g["team2"].upper())}</span></div>
  </div>
  <div class="rosters">{roster_table(g["roster1"], slots)}{roster_table(g["roster2"], slots)}</div>
  <div class="row2 res">
    <div class="fld score"><span class="lb">ИТОГОВЫЙ&nbsp;СЧЁТ</span>
        <span class="vl big">{e(score_text(g["final_raw"]))}</span>{mark}</div>
    <div class="fld"><span class="lb">ПОБЕДИТЕЛЬ</span><span class="vl team big">{e(winner(g).upper())}</span></div>
  </div>
  <div class="fld wide2"><span class="lb">СЧЁТ&nbsp;ПО&nbsp;ПЕРИОДАМ</span><span class="vl">{e(periods_text(g))}</span></div>
  <div class="fld off nowrap"><span class="lb">СУДЬЯ</span><span class="vl">{e(official(g["judge"]))}</span></div>
  <div class="fld off nowrap"><span class="lb">СЕКРЕТАРЬ</span><span class="vl">{e(official(g["secretary"]))}</span></div>
</div>'''


def day_sheets(day_games, slots, logo, division):
    """Листы одного игрового дня: -> (html, [матчи каждого листа])."""
    date = day_games[0]["date"].strftime("%d.%m.%Y")
    chunks = [day_games[i:i + CARDS_PER_SHEET] for i in range(0, len(day_games), CARDS_PER_SHEET)]
    out = ""
    for i, chunk in enumerate(chunks):
        label = f'<span class="dt">{e(date)}</span>'
        foot = f'<div class="foot">Лист {i + 1} из {len(chunks)}</div>'
        out += (f'<div class="sheet"><div class="head">'
                f'<div class="top"><img class="logo" src="data:image/png;base64,{logo}" alt="РХЛ">'
                # заглавные сразу: при text-transform браузер уводит значок над Й вниз страницы
                f'<div class="league">{e(LEAGUE_TITLE.upper())}</div><div></div></div>'
                f'<div class="division">{e(division)}</div>'
                f'<div class="subhead">'
                f'<div class="side left"></div>'
                f'<div class="title">ПРОТОКОЛ ИГРОВОГО ДНЯ</div>'
                f'<div class="side right"><div class="date">{label}</div></div>'
                f'</div></div>'
                f'<div class="grid">{"".join(card_html(g, slots) for g in chunk)}</div>{foot}</div>')
    return out, chunks


def file_name(fname):
    m = re.match(r"(\d{4})-(\d{2})_([a-z]+)_div(\d+)", fname)
    y, mo, conf, div = int(m.group(1)), int(m.group(2)), m.group(3), m.group(4)
    return f"Протоколы_{CONF[conf]}_{div}_дивизион_{MONTHS[mo]}_{y}.pdf"


def sq(s):
    """Текст без переносов и лишних пробелов: в PDF длинные ФИО переносятся внутри ячейки."""
    return re.sub(r"\s+", " ", s or "").strip()


MEASURE = """<script>
window.addEventListener('load', () => {
  const out = [];
  document.querySelectorAll('.sheet').forEach((s, i) => {
    if (s.scrollHeight > s.clientHeight + 1) out.push('лист ' + (i + 1) + ': не помещается по высоте');
    s.querySelectorAll('.card, .card *').forEach(el => {
      if (el.scrollWidth > el.clientWidth + 1 && el.clientWidth > 0)
        out.push('лист ' + (i + 1) + ': шире карточки — ' + el.textContent.trim().replace(/\\s+/g, ' ').slice(0, 60));
    });
  });
  const pre = document.createElement('pre'); pre.id = 'measure'; pre.textContent = out.join('\\n');
  document.body.appendChild(pre);
});
</script>"""


def measure_layout(page_html, tmp):
    """Вёрстка в браузере шириной в печатную полосу (182 мм): что не помещается."""
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(page_html.replace("</body>", MEASURE + "</body>"))
    dom = subprocess.run([CHROME, "--headless", "--no-sandbox", "--disable-gpu", "--window-size=688,1200",
                          "--virtual-time-budget=5000", "--dump-dom", f"file://{tmp}"],
                         capture_output=True, text=True).stdout
    m = re.search(r'<pre id="measure">(.*?)</pre>', dom, re.S)
    if m is None:
        return ["замер вёрстки не выполнен"]
    return [x for x in html.unescape(m.group(1)).split("\n") if x.strip()]


def page_words(pdf):
    """Слова каждой страницы с координатами: [(текст, x0, y0, x1, y1), ...]."""
    out = subprocess.run(["pdftotext", "-bbox", "-enc", "UTF-8", pdf, "-"],
                         capture_output=True, text=True, check=True).stdout
    pages = []
    for page in re.findall(r"<page [^>]*>(.*?)</page>", out, re.S):
        pages.append([(html.unescape(t), float(a), float(b), float(c), float(d)) for a, b, c, d, t in
                      re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</word>', page)])
    return pages


def numbered_surnames(words):
    """Пары (номер, ФАМИЛИЯ): номер — в соседней ячейке слева, по высоте в пределах строки игрока
    (фамилия и под ней имя и отчество)."""
    nums = [w for w in words if w[0].isdigit()]
    pairs = Counter()
    for t, x0, y0, x1, y1 in words:
        if not (t.isupper() and len(t) > 1):
            continue
        for n, nx0, ny0, nx1, ny1 in nums:
            mid = (ny0 + ny1) / 2
            if x0 - 30 < nx1 <= x0 and y0 - 2 <= mid <= y1 + 12:
                pairs[(n, t)] += 1
                break
    return pairs


def check_pdf(pdf, sheets):
    """Сверка готового PDF: листы и все значения каждой карточки на своём листе."""
    def pdf_pages(*opts):
        text = subprocess.run(["pdftotext", "-enc", "UTF-8", *opts, pdf, "-"],
                              capture_output=True, text=True, check=True).stdout
        out = text.split("\f")
        return out[:-1] if out and not out[-1].strip() else out
    pages = [sq(p) for p in pdf_pages()]
    words = page_words(pdf)
    problems = []
    if len(pages) != len(sheets):
        problems.append(f"страниц {len(pages)}, а листов должно быть {len(sheets)}")
    for i, chunk in enumerate(sheets):
        page = pages[i] if i < len(pages) else ""
        if f"Лист {sum(1 for c in sheets[:i + 1] if c[0]['date'] == chunk[0]['date'])} из" not in page:
            problems.append(f"лист {i + 1}: нет номера листа внизу страницы")
        if page.count("СОРЕВНОВАНИЕ") != len(chunk):
            problems.append(f"лист {i + 1}: карточек {page.count('СОРЕВНОВАНИЕ')}, а матчей {len(chunk)}")
        for g in chunk:
            want = [g["date"].strftime("%d.%m.%Y"), g["time"], g["team1"].upper(), g["team2"].upper(),
                    official(g["judge"]), official(g["secretary"]), score_text(g["final_raw"]),
                    sq(periods_text(g)), winner(g).upper()]
            # фамилия сверяется ниже вместе с номером (одна строка таблицы), имя и отчество — здесь
            want += [player_parts(nm)[1] for _, nm in g["roster1"] + g["roster2"]]
            missing = [w for w in want if w and sq(w) not in page]
            if missing:
                problems.append(f"лист {i + 1}, стр. {g['row']} {g['team1']} — {g['team2']}: не найдено {missing[:4]}")
        # номер игрока стоит слева от его фамилии в той же строке таблицы: сколько раз пара
        # (номер, фамилия) встречается на листе — столько раз она должна стоять в карточках
        need = Counter((n, player_parts(nm)[0]) for g in chunk for n, nm in g["roster1"] + g["roster2"])
        have = numbered_surnames(words[i] if i < len(words) else [])
        short = [f"№{n} {sur}" for (n, sur), k in need.items() if have[(n, sur)] < k]
        if short:
            problems.append(f"лист {i + 1}: номер не рядом с фамилией: {short[:4]}")
    return problems


def build(fname, logo):
    games, _ = corrections.apply(fname, read_workbook(os.path.join(SRC, fname))[0])
    division = games[0]["comp"]
    slots = max(max(len(g["roster1"]), len(g["roster2"])) for g in games)
    byday = defaultdict(list)
    for g in games:
        byday[g["date"]].append(g)
    body, sheets = "", []
    for d in sorted(byday):
        html_, chunks = day_sheets(sorted(byday[d], key=lambda g: g["time"]), slots, logo, division)
        body += html_
        sheets += chunks
    os.makedirs(OUT, exist_ok=True)
    pdf = os.path.join(OUT, file_name(fname))
    tmp = os.path.join(OUT, "_tmp.html")
    page = ('<!doctype html><html><head><meta charset="utf-8">'
            f'<style>{CSS}{CSS_EXTRA}</style></head><body>{body}</body></html>')
    layout_problems = measure_layout(page, tmp)
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(page)
    subprocess.run([CHROME, "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
                    f"--print-to-pdf={pdf}", f"file://{tmp}"], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    os.remove(tmp)
    problems = layout_problems + check_pdf(pdf, sheets)
    print(f"{os.path.basename(pdf)}: матчей {len(games)}, игровых дней {len(byday)}, листов {len(sheets)} — "
          + ("сверено, всё на месте" if not problems else f"РАСХОЖДЕНИЯ: {len(problems)}"))
    for p in problems[:10]:
        print("   ", p)
    return not problems


def main(files):
    files = files or sorted(f for f in os.listdir(SRC) if f.endswith(".xlsx"))
    logo = logo_b64()
    ok = all([build(f, logo) for f in files])
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
