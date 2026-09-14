# -*- coding: utf-8 -*-
"""Автосверка: исходный Excel <-> сгенерированные PDF.

Проверяет, что каждая игра попала в отчёт и что все значения перенесены без искажений.
Ничего не исправляет — только сравнивает и сообщает.

Запуск:  python3 tools/validate.py   (после build_reports.py)
"""
import os, re, sys, collections
import pymupdf
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_reports import read_games, XLSX, OUT, ROOT, CARDS_PER_SHEET

def norm(s):
    """Схлопывает любые переносы/пробелы — в PDF имена переносятся внутри ячеек."""
    return re.sub(r"\s+", " ", s or "").strip()

def main():
    games, label_warnings = read_games(XLSX)
    report = []
    ok = True

    def say(line):
        report.append(line)
        print(line)

    say("=" * 72)
    say("ОТЧЁТ О СВЕРКЕ: Excel -> PDF")
    say("=" * 72)

    # --- 1. исходник
    by_sheet = collections.Counter(g["sheet"] for g in games)
    say(f"\n1. ИСХОДНЫЕ ДАННЫЕ")
    say(f"   Игр прочитано: {len(games)}")
    for sh, n in by_sheet.items():
        say(f"     лист '{sh}': {n}")

    # --- 2. количество карточек в PDF
    say(f"\n2. КАРТОЧЕК В ПОМЕСЯЧНЫХ PDF")
    pages_text = []
    months = []
    for sh in dict.fromkeys(g["sheet"] for g in games):
        months.append(sh)
    for sh in months:
        path = os.path.join(OUT, f"Протоколы_{sh}.pdf")
        d = pymupdf.open(path)
        texts = [norm(p.get_text()) for p in d]
        d.close()
        pages_text.extend(texts)
        cards = sum(t.count("СОРЕВНОВАНИЕ") for t in texts)
        expected = sum(1 for g in games if g["sheet"] == sh)
        mark = "OK" if cards == expected else "ОШИБКА"
        if cards != expected:
            ok = False
        say(f"   {os.path.basename(path)}: листов {len(texts)}, карточек {cards}, "
            f"игр в Excel {expected} — {mark}")
    cards_in_pdf = sum(t.count("СОРЕВНОВАНИЕ") for t in pages_text)
    say(f"   ВСЕГО: листов {len(pages_text)}, карточек {cards_in_pdf}, игр {len(games)}")
    if cards_in_pdf == len(games):
        say(f"   OK: {cards_in_pdf} = {len(games)} — ничего не потеряно и не задвоено")
    else:
        ok = False
        say(f"   ОШИБКА: в PDF {cards_in_pdf}, в Excel {len(games)}")

    # --- 3. пополевая сверка каждой игры в PDF своего дня
    say(f"\n3. ПОПОЛЕВАЯ СВЕРКА КАЖДОЙ ИГРЫ")
    byday = collections.defaultdict(list)
    for g in games:
        byday[(g["sheet"], g["date_key"])].append(g)
    checked = mismatched = 0
    problems = []
    for (sheet, date_key), day_games in byday.items():
        path = os.path.join(OUT, "by_day", f"{date_key}_{sheet}.pdf")
        d = pymupdf.open(path)
        text = norm(" ".join(p.get_text() for p in d))
        d.close()
        for g in day_games:
            missing = []
            fields = [("соревнование", g["comp"]), ("дата", g["date"]), ("время", g["time"]),
                      ("команда1", g["t1"]), ("команда2", g["t2"]),
                      ("итог", g["final"]), ("победитель", g["winner"])]
            for label, val in fields:
                if val and norm(val) not in text:
                    missing.append(f"{label}='{val}'")
            for who, roster in (("состав1", g["roster1"]), ("состав2", g["roster2"])):
                for num, name in roster:
                    if norm(name) not in text:
                        missing.append(f"{who}:'{name}'")
            for p in g["periods_raw"]:
                disp = norm(p.replace("*", " : "))
                if disp not in text:
                    missing.append(f"период='{disp}'")
            checked += 1
            if missing:
                mismatched += 1
                problems.append((sheet, g["index"], g["row"], missing))
    say(f"   Проверено игр: {checked}")
    say(f"   С расхождениями: {mismatched}")
    if mismatched:
        ok = False
        for sheet, idx, row, missing in problems[:20]:
            say(f"     лист '{sheet}' игра {idx} (стр.{row}): не найдено {missing[:4]}")
    else:
        say("   OK: все поля всех игр найдены в PDF в неизменном виде")

    # --- 4. контрольные суммы голов
    say(f"\n4. КОНТРОЛЬНЫЕ СУММЫ ГОЛОВ")
    src_h = sum(g["fh"] for g in games if g["fh"] is not None)
    src_a = sum(g["fa"] for g in games if g["fa"] is not None)
    pdf_scores = re.findall(r"ИТОГОВЫЙ СЧЁТ (\d+) : (\d+)", " ".join(pages_text))
    pdf_h = sum(int(a) for a, b in pdf_scores)
    pdf_a = sum(int(b) for a, b in pdf_scores)
    say(f"   Excel: команда А = {src_h}, команда В = {src_a}  (всего {src_h + src_a})")
    say(f"   PDF  : команда А = {pdf_h}, команда В = {pdf_a}  (всего {pdf_h + pdf_a})")
    if (src_h, src_a) == (pdf_h, pdf_a) and len(pdf_scores) == len(games):
        say(f"   OK: суммы сходятся, счётов найдено {len(pdf_scores)}")
    else:
        ok = False
        say(f"   ОШИБКА: расхождение (счётов в PDF {len(pdf_scores)})")

    # --- 5. аномалии исходных данных (НЕ исправлены, перенесены как есть)
    say(f"\n5. АНОМАЛИИ В ИСХОДНЫХ ДАННЫХ (перенесены как есть, на вашу проверку)")
    say(f"\n   5.1 Сумма по периодам != итоговый счёт:")
    n = 0
    for g in games:
        if g["fh"] is None:
            continue
        try:
            sh = sum(int(p.split("*")[0]) for p in g["periods_raw"])
            sa = sum(int(p.split("*")[1]) for p in g["periods_raw"])
        except (ValueError, IndexError):
            continue
        if (sh, sa) != (g["fh"], g["fa"]):
            n += 1
            say(f"     лист '{g['sheet']}' игра {g['index']} (стр.{g['row']}), {g['date']} "
                f"{g['time']}, {g['t1']} — {g['t2']}: итог {g['final_raw']}, "
                f"периоды {' '.join(g['periods_raw'])} (сумма {sh}*{sa})")
    say(f"     Всего: {n}")

    say(f"\n   5.2 Опечатки в подписях левого столбца (на разбор не влияют):")
    for w in label_warnings:
        say(f"     лист '{w['sheet']}' игра {w['game']}, строка {w['row']}: "
            f"'{w['got']}' вместо '{w['expected']}'")
    say(f"     Всего: {len(label_warnings)}")

    draws = [g for g in games if g["winner"] == "ничья"]
    say(f"\n   5.3 Игры, закончившиеся вничью (в поле ПОБЕДИТЕЛЬ стоит 'ничья'):")
    for g in draws:
        say(f"     лист '{g['sheet']}' игра {g['index']}, {g['date']} {g['time']}, "
            f"{g['t1']} — {g['t2']}: {g['final_raw']}")
    say(f"     Всего: {len(draws)}")

    say("\n" + "=" * 72)
    say("ИТОГ: СВЕРКА ПРОЙДЕНА — данные перенесены точно" if ok
        else "ИТОГ: ЕСТЬ ОШИБКИ ПЕРЕНОСА — см. выше")
    say("=" * 72)

    with open(os.path.join(ROOT, "reports", "otchet_o_sverke.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(report) + "\n")
    return 0 if ok else 1

if __name__ == "__main__":
    sys.exit(main())
