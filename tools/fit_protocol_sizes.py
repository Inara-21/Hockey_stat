# -*- coding: utf-8 -*-
"""Подбор размеров шрифтов протоколов: максимальные, при которых на данных всех
дивизионов всё помещается (замер вёрстки в браузере, как перед печатью).

Размеры в карточке (подписи, данные, счёт, шапка таблицы, фамилия, имя и отчество)
увеличиваются вместе, в одной пропорции к tools/make_protocols.py:SIZES.

Запуск:  python3 tools/fit_protocol_sizes.py 1.0 1.1 1.2      проверить множители
"""
import os
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import make_protocols as mp

CARD = ("label", "data", "big", "th", "sur", "io")


def problems_for(k):
    sizes = {key: round(mp.SIZES[key] * k, 2) for key in CARD}
    logo = mp.logo_b64()
    files = sorted(f for f in os.listdir(mp.SRC) if f.endswith(".xlsx"))
    tmpdir = tempfile.mkdtemp()

    def one(f):
        page, *_ = mp.page_html(f, logo, sizes)
        return [f"{f}: {x}" for x in mp.measure_layout(page, os.path.join(tmpdir, f + ".html"))]

    with ThreadPoolExecutor(4) as ex:
        out = [x for r in ex.map(one, files) for x in r]
    return sizes, out


if __name__ == "__main__":
    for k in map(float, sys.argv[1:] or ["1.0"]):
        sizes, out = problems_for(k)
        print(f"множитель {k}: {sizes} — " + ("всё помещается" if not out else f"не помещается: {len(out)}"))
        for x in out[:8]:
            print("   ", x)
