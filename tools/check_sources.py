# -*- coding: utf-8 -*-
"""Проверка исходных файлов на ошибки ДО формирования отчётов.

Ничего не исправляет и ничего не формирует — только читает и сообщает.
Разбор устойчивый: блоки ищутся по заголовку «Название соревнования»,
а не по фиксированному шагу строк (в файлах встречаются пустые строки
между играми и опечатки в подписях).
"""
import openpyxl, datetime, os, re, sys
from collections import Counter, defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "data", "source")

# смещения ЗНАЧЕНИЙ от строки заголовка блока
OFF = dict(comp=0, date=1, time=2, t1=3, t2=4, roster1=6, roster2=8,
           secretary=10, judge=12, final=13, periods=14)
BLOCK_ROWS = 15
# смещения ПОДПИСЕЙ (у составов/судей подпись на строку выше значений)
LABELS = {0: "Название соревнования", 1: "Дата", 2: "Время начала матча",
          3: "Команда 1", 4: "Команда2", 5: "Состав команды 1",
          7: "Состав команды 2", 9: "Секретари", 11: "Судьи",
          13: "Итоговый счёт матча", 14: "Счёт по периодам"}


def s(ws, r, c):
    v = ws.cell(row=r, column=c).value
    return "" if v is None else str(v).strip()


def to_date(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return datetime.date(v.year, v.month, v.day), "дата"
    t = "" if v is None else str(v).strip()
    for f in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%y"):
        try:
            return datetime.datetime.strptime(t, f).date(), "текст"
        except ValueError:
            pass
    return None, "НЕРАЗОБРАНО"


def to_time(v):
    if isinstance(v, (datetime.time, datetime.datetime)):
        return v.strftime("%H:%M"), "время"
    t = "" if v is None else str(v).strip()
    if re.fullmatch(r"\d{1,2}[:.]\d{2}(:\d{2})?", t):
        return t.replace(".", ":")[:5], "текст"
    return None, "НЕРАЗОБРАНО"


def parse_score(raw):
    raw = (raw or "").strip()
    for sep in ("*", ":", "-", "х", "x"):
        if sep in raw:
            a, b = raw.split(sep, 1)
            try:
                return int(a.strip()), int(b.strip())
            except ValueError:
                return None, None
    return None, None


def read_file(path):
    """-> (games, issues). issues — список (уровень, лист, строка, текст)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    games, issues = [], []
    name = os.path.basename(path)

    for ws in wb.worksheets:
        starts = [r for r in range(1, ws.max_row + 1)
                  if s(ws, r, 1).lower().startswith(("назв", "нава"))
                  and "соревн" in s(ws, r, 1).lower()]
        if not starts:
            issues.append(("инфо", ws.title, 0,
                           f"лист без игр ({ws.max_row} строк) — пропущен как не относящийся к матчам"))
            continue

        # контроль шага блоков
        for a, b in zip(starts, starts[1:]):
            if b - a not in (BLOCK_ROWS, BLOCK_ROWS + 1):
                issues.append(("ОШИБКА", ws.title, a,
                               f"нестандартный размер блока игры: {b - a} строк вместо 15"))

        for b in starts:
            g = dict(file=name, sheet=ws.title, row=b)
            g["comp"] = s(ws, b + OFF["comp"], 2)
            d, dkind = to_date(ws.cell(row=b + OFF["date"], column=2).value)
            t, tkind = to_time(ws.cell(row=b + OFF["time"], column=2).value)
            g["date"], g["time"] = d, t
            g["t1"], g["t2"] = s(ws, b + OFF["t1"], 2), s(ws, b + OFF["t2"], 2)
            g["roster1"] = [s(ws, b + OFF["roster1"], c)
                            for c in range(2, ws.max_column + 1) if s(ws, b + OFF["roster1"], c)]
            g["roster2"] = [s(ws, b + OFF["roster2"], c)
                            for c in range(2, ws.max_column + 1) if s(ws, b + OFF["roster2"], c)]
            g["secretary"] = s(ws, b + OFF["secretary"], 2)
            g["judge"] = s(ws, b + OFF["judge"], 2)
            g["final"] = s(ws, b + OFF["final"], 2)
            g["periods"] = [s(ws, b + OFF["periods"], c)
                            for c in range(2, ws.max_column + 1) if s(ws, b + OFF["periods"], c)]
            g["date_kind"], g["time_kind"] = dkind, tkind
            games.append(g)

            def add(level, text):
                issues.append((level, ws.title, b, text))

            # --- обязательные поля
            if not g["comp"]:
                add("ОШИБКА", "пустое название соревнования")
            if d is None:
                add("ОШИБКА", f"не удалось разобрать дату: {ws.cell(row=b+1, column=2).value!r}")
            if t is None:
                add("ОШИБКА", f"не удалось разобрать время: {ws.cell(row=b+2, column=2).value!r}")
            if not g["t1"] or not g["t2"]:
                add("ОШИБКА", f"пустое название команды: '{g['t1']}' — '{g['t2']}'")
            if g["t1"] and g["t1"] == g["t2"]:
                add("ОШИБКА", f"команда играет сама с собой: {g['t1']}")
            if not g["roster1"] or not g["roster2"]:
                add("ОШИБКА", f"пустой состав: к1={len(g['roster1'])}, к2={len(g['roster2'])}")
            if not g["judge"]:
                add("внимание", "не указан судья")
            if not g["secretary"]:
                add("внимание", "не указан секретарь")

            # --- счёт
            fh, fa = parse_score(g["final"])
            g["fh"], g["fa"] = fh, fa
            if not g["final"]:
                add("ОШИБКА", "пустой итоговый счёт")
            elif fh is None:
                add("ОШИБКА", f"не удалось разобрать итоговый счёт: '{g['final']}'")
            if not g["periods"]:
                add("ОШИБКА", "пустой счёт по периодам")
            bad = [p for p in g["periods"] if parse_score(p)[0] is None]
            if bad:
                add("ОШИБКА", f"не удалось разобрать счёт по периодам: {bad}")
            elif fh is not None:
                sh = sum(parse_score(p)[0] for p in g["periods"])
                sa = sum(parse_score(p)[1] for p in g["periods"])
                if (sh, sa) != (fh, fa):
                    add("внимание", f"сумма по периодам {sh}*{sa} != итог {g['final']}")

            # --- игроки
            for who, roster in (("к1", g["roster1"]), ("к2", g["roster2"])):
                for p in roster:
                    if "№" not in p:
                        add("внимание", f"у игрока нет номера ({who}): '{p}'")
                nums = [p.rsplit("№", 1)[1].strip() for p in roster if "№" in p]
                dup = [n for n, c in Counter(nums).items() if c > 1]
                if dup:
                    add("ОШИБКА", f"повторяющиеся игровые номера ({who}): {dup}")

            # --- подписи
            for off, expected in LABELS.items():
                got = s(ws, b + off, 1)
                if got != expected:
                    add("инфо", f"подпись строки {b+off}: '{got}' вместо '{expected}'")

    return games, issues


def key(g):
    return (g["comp"], str(g["date"]), g["time"], g["t1"], g["t2"])


def content(g):
    return (g["final"], tuple(g["periods"]), tuple(g["roster1"]),
            tuple(g["roster2"]), g["secretary"], g["judge"])


def main():
    files = sorted(os.listdir(SRC))
    all_games, report = {}, []

    def say(line=""):
        report.append(line)
        print(line)

    say("=" * 78)
    say("ПРОВЕРКА ИСХОДНЫХ ФАЙЛОВ")
    say("=" * 78)

    total_err = total_warn = 0
    for f in files:
        games, issues = read_file(os.path.join(SRC, f))
        all_games[f] = games
        errs = [i for i in issues if i[0] == "ОШИБКА"]
        warns = [i for i in issues if i[0] == "внимание"]
        infos = [i for i in issues if i[0] == "инфо"]
        total_err += len(errs)
        total_warn += len(warns)

        say(f"\n{'-' * 78}\nФАЙЛ: {f}")
        by_sheet = Counter(g["sheet"] for g in games)
        say(f"  игр: {len(games)} | листы: {dict(by_sheet)}")
        rost = Counter((len(g['roster1']), len(g['roster2'])) for g in games)
        say(f"  размеры составов (к1,к2): {dict(sorted(rost.items()))}")
        dk = Counter(g["date_kind"] for g in games)
        tk = Counter(g["time_kind"] for g in games)
        say(f"  формат даты: {dict(dk)} | формат времени: {dict(tk)}")

        dups = [k for k, c in Counter(key(g) for g in games).items() if c > 1]
        if dups:
            total_err += len(dups)
            say(f"  ОШИБКА: повторяющиеся игры внутри файла: {len(dups)}")
            for d in dups[:5]:
                say(f"     {d}")
        else:
            say("  повторов внутри файла нет")

        say(f"  ОШИБОК: {len(errs)} | предупреждений: {len(warns)} | замечаний к подписям: {len(infos)}")
        for lvl, sheet, row, text in errs[:30]:
            say(f"     ОШИБКА  лист '{sheet}' стр.{row}: {text}")
        for lvl, sheet, row, text in warns[:40]:
            say(f"     внимание лист '{sheet}' стр.{row}: {text}")
        for lvl, sheet, row, text in infos[:10]:
            say(f"     замечание {sheet} стр.{row}: {text}")

    # ---- сверка файлов между собой
    say(f"\n{'=' * 78}\nСВЕРКА ФАЙЛОВ МЕЖДУ СОБОЙ")
    say("=" * 78)
    index = {}
    conflicts = []
    for f in files:
        for g in all_games[f]:
            k = key(g)
            if k in index:
                prev = index[k]
                if content(prev) != content(g):
                    conflicts.append((k, prev, g))
            else:
                index[k] = g
    say(f"\nуникальных игр во всех файлах: {len(index)}")
    say(f"игр с ПРОТИВОРЕЧИВЫМ содержимым в разных файлах: {len(conflicts)}")

    kinds = Counter()
    for k, a, b in conflicts:
        if not a["final"] and b["final"]:
            kinds["в одном файле счёт пустой, в другом заполнен"] += 1
        elif a["final"] and b["final"] and a["final"] != b["final"]:
            kinds["РАЗНЫЙ итоговый счёт"] += 1
        else:
            kinds["прочие расхождения (составы/судьи/периоды)"] += 1
    for kk, v in kinds.items():
        say(f"   {kk}: {v}")

    hard = [(k, a, b) for k, a, b in conflicts
            if a["final"] and b["final"] and a["final"] != b["final"]]
    if hard:
        say(f"\n   ТРЕБУЮТ РЕШЕНИЯ — разный счёт у одной и той же игры:")
        for k, a, b in hard:
            say(f"     {k[1]} {k[2]}  {k[3]} — {k[4]}")
            say(f"        {a['file']} (стр.{a['row']}): итог {a['final']}, периоды {' '.join(a['periods'])}")
            say(f"        {b['file']} (стр.{b['row']}): итог {b['final']}, периоды {' '.join(b['periods'])}")

    say(f"\n{'=' * 78}")
    say(f"ИТОГО по всем файлам: ошибок {total_err}, предупреждений {total_warn}")
    say("=" * 78)

    with open(os.path.join(ROOT, "reports", "proverka_ishodnikov.txt"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(report) + "\n")


if __name__ == "__main__":
    main()
