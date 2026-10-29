# -*- coding: utf-8 -*-
"""Сверка матча с сайтом лиги rhl-rus.ru: что там записано по этому матчу.

Запуск:  python3 tools/site_lookup.py 2026-09_centr_div1.xlsx 811 1336 1801

Для каждой строки печатает матч из нашего файла (после правок реестра) и матчи
этих же команд в этот день на сайте; тот же матч (время и порядок команд)
помечен стрелкой. Страницы сайта кешируются в data/site_cache (не в git).

Сайт — только справка. Брать с него можно счёт и время начала, когда есть
сомнения, и только по согласованию с заказчиком. Ответ проверяющего важнее сайта.
Составы на сайте не совпадают с протоколами — их не брать.
"""
import html
import os
import re
import subprocess
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import corrections
from hockey_data import read_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE = os.path.join(ROOT, "data", "site_cache")
# дивизион файла -> код лиги на сайте
CODES = {("vostok", "1"): "konferentsiya-vostok-1-divizion", ("vostok", "2"): "konferenciya-vostok-2-division",
         ("zapad", "1"): "konferentsiya-zapad-1-divizion", ("zapad", "2"): "konferentsiya-zapad-2-divizion",
         ("centr", "1"): "konferenciya-center-1-division", ("centr", "2"): "division-center-2-liga",
         ("centr", "3"): "division-center-3-liga"}


def day_page(code, d):
    """Страница расписания дивизиона за день (кешируется)."""
    os.makedirs(CACHE, exist_ok=True)
    path = os.path.join(CACHE, f"{code}_{d:%Y%m%d}.html")
    if not os.path.exists(path):
        url = f"https://rhl-rus.ru/calendar/?league={code}&Y={d.year}&M={d.month:02d}&D={d.day}"
        subprocess.run(["curl", "-sS", "--max-time", "40", "-o", path, url], check=True)
        time.sleep(0.7)
    return open(path, encoding="utf-8", errors="replace").read()


def norm_team(s):
    s = html.unescape(s).lower().replace("ё", "е")
    s = re.sub(r"\(.*?\)", "", s)
    s = re.sub(r"^хк\s+", "", s.strip())
    return re.sub(r"\s+", " ", s).strip()


def parse_day(t):
    """Матчи со страницы: время, команды, итог (с пометкой Б), периоды, буллиты."""
    out = []
    for m in re.finditer(r'<a href="(/calendar/\d+/)" class="matche"([\s\S]*?)</a>', t):
        b = m.group(2)
        tm = re.search(r'class="time">([^<]+)<', b).group(1).strip()
        names = [html.unescape(x).strip() for x in re.findall(r'class="name">([^<]+)<', b)]
        vals = [re.sub(r"\s+", "", html.unescape(x)) for x in re.findall(r'class="val[^"]*">([^<]*)<', b)]
        items = [re.sub(r"\s+", "", re.sub(r"<[^>]+>", "", x)) for x in re.findall(r'<li[^>]*>([\s\S]*?)</li>', b)]
        per = [x for x in items if re.fullmatch(r"\d+:\d+", x)]
        so = [x.strip("()") for x in items if re.fullmatch(r"\(\d+:\d+\)", x)]
        if len(per) + len(so) != len(items):
            raise ValueError(f"не разобраны периоды: {items}")
        shootout = any("Б" in v for v in vals) or bool(so)
        a, b2 = (re.sub(r"\D", "", x) for x in (vals[:2] if len(vals) >= 2 else ["", ""]))
        out.append(dict(url="https://rhl-rus.ru" + m.group(1), time=tm, t1=names[0], t2=names[1],
                        final=f"{a}:{b2}" + (" Б" if shootout else "") if a else "нет счёта",
                        periods=per + [f"буллиты {x}" for x in so]))
    return out


def main(fname, rows):
    conf, div = re.match(r"\d{4}-\d{2}_([a-z]+)_div(\d+)", fname).groups()
    games, _ = corrections.apply(fname, read_workbook(os.path.join(ROOT, "data", "source", fname))[0])
    by_row = {g["row"]: g for g in games}
    for row in rows:
        g = by_row[row]
        print(f"стр. {row}  {g['date']:%d.%m.%Y} {g['time'] or 'без времени'}  {g['team1']} — {g['team2']}  "
              f"| у нас: итог {g['final_raw']}, периоды {' '.join(g['periods_raw'])}")
        pair = {norm_team(g["team1"]), norm_team(g["team2"])}
        for m in parse_day(day_page(CODES[(conf, div)], g["date"])):
            if {norm_team(m["t1"]), norm_team(m["t2"])} != pair:
                continue
            same = m["time"] == g["time"] and norm_team(m["t1"]) == norm_team(g["team1"])
            print(f"   {'==>' if same else '   '} сайт {m['time']} {m['t1']} — {m['t2']}: итог {m['final']}, "
                  f"периоды {' '.join(m['periods'])}  {m['url']}")


if __name__ == "__main__":
    main(sys.argv[1], [int(x) for x in sys.argv[2:]])
