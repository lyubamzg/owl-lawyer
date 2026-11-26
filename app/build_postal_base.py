#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сборка локальной базы почтовых индексов РФ для «Совы-юриста».

Источник — эталонный справочник почтовых индексов ОПС АО «Почта России»
(выгрузка anpalcactus/russian-postal-codes, CC BY 4.0). В нём все почтовые
индексы страны: 61 208 индексов, 30 774 населённых пункта, справочник улиц
для городов, где индексов несколько.

Что получается в папке data/:
  postal_base.json.gz    — весь справочник индексов + карта «населённый
                           пункт → индекс» (для поиска по адресу офлайн);
  postal_streets.txt.gz  — улицы индексов (только там, где индексов много).

Запуск:
    python3 build_postal_base.py            # скачать источник и собрать базу
    python3 build_postal_base.py --force    # пересобрать, не скачивая заново
"""
import csv
import gzip
import json
import re
import sys
import urllib.request
from collections import defaultdict
from pathlib import Path

BASE = Path(__file__).resolve().parent
DATA = BASE / 'data'
CACHE = Path.home() / '.cache' / 'owl-lawyer' / 'postal'

SRC = 'https://raw.githubusercontent.com/anpalcactus/russian-postal-codes/main/data'
FILES = {'indexes.csv': 'indexes.csv', 'streets.csv.gz': 'streets.csv.gz'}

OUT_BASE = DATA / 'postal_base.json.gz'
OUT_STREETS = DATA / 'postal_streets.txt.gz'

# «г.», «пос.» и прочие — служебные, в название населённого пункта не входят
TYPE_PREFIX_RE = re.compile(
    r'^(?:г|гор|город|пос|поселок|село|с|сп|ст|ст-ца|станица|д|дер|деревня|'
    r'хут|хуттор|с-ц|мкр|мк-н|микрорайон|пер|пер-д|зимник|курорт)',
    re.IGNORECASE)


def norm(text):
    """«посёлок им. Кирова» → «поселок им кирова» (для сопоставления)."""
    s = str(text or '').strip().lower().replace('ё', 'е')
    s = re.sub(r'[^0-9a-zа-я]+', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def download(name):
    """Скачать исходный файл (в кэш) — если его там ещё нет."""
    CACHE.mkdir(parents=True, exist_ok=True)
    path = CACHE / name
    if path.is_file() and path.stat().st_size:
        return path
    url = SRC + '/' + name
    print('качаю %s …' % url, flush=True)
    req = urllib.request.Request(url, headers={'User-Agent': 'OwlLawyer/1.0'})
    with urllib.request.urlopen(req, timeout=120) as resp:
        data = resp.read()
    path.write_bytes(data)
    print('  %s (%.1f МБ)' % (path, len(data) / 1048576.0))
    return path


def build(force=False):
    DATA.mkdir(parents=True, exist_ok=True)
    if force:
        idx_path = CACHE / 'indexes.csv'
        st_path = CACHE / 'streets.csv.gz'
    else:
        idx_path = download('indexes.csv')
        st_path = download('streets.csv.gz')

    with open(idx_path, encoding='utf-8') as f:
        rows = list(csv.DictReader(f))

    release = ''
    # география → индексы
    geo = defaultdict(list)          # (регион, район, населённый пункт) → [(имя, индекс)]
    indexes = {}                      # индекс → название объекта
    regions = {}                      # регион → основной индекс
    for r in rows:
        idx = (r.get('index') or '').strip()
        name = (r.get('name') or '').strip()
        if idx:
            indexes[idx] = name
        if (r.get('addressable') or '').strip() != '1':
            continue
        region = (r.get('region') or '').strip()
        area = (r.get('area') or '').strip()
        city = (r.get('city') or '').strip()
        settle = (r.get('settlement') or '').strip()
        place = settle or city or region          # Москва: регион и есть город
        if not place:
            continue
        geo[(region, area, place)].append((name, idx))
        if region and region not in regions:
            regions[region] = idx

    # карта «населённый пункт → варианты»
    settlements = {}
    multi = set()
    for (region, area, place), items in geo.items():
        items.sort(key=lambda x: (0 if x[0].lower() == place.lower() else 1, x[1]))
        key = norm(place)
        entry = [region, area, items[0][1], len(items)]
        settlements.setdefault(key, []).append(entry)
        if len(items) > 1:
            multi.add(items[0][1])
            multi.update(i for _, i in items)

    # регион → индекс (основной: индекс с названием, совпадающим с регионом)
    reg_index = {}
    for (region, area, place), items in geo.items():
        if area:
            continue
        for name, idx in items:
            if name.lower() == region.lower():
                reg_index[region] = idx
                break
        else:
            reg_index.setdefault(region, items[0][1])

    release = _release(rows)
    base = {
        'release': release,
        'source': 'АО «Почта России», эталонный справочник почтовых индексов ОПС',
        'source_url': 'https://github.com/anpalcactus/russian-postal-codes',
        'indexes': indexes,
        'settlements': settlements,
        'regions': {norm(k): v for k, v in reg_index.items()},
    }
    with gzip.open(OUT_BASE, 'wt', encoding='utf-8') as f:
        json.dump(base, f, ensure_ascii=False, separators=(',', ':'))
    print('  %s (%.1f МБ): индексов %d, населённых пунктов %d'
          % (OUT_BASE.name, OUT_BASE.stat().st_size / 1048576.0,
             len(indexes), len(settlements)))

    # улицы — только для населённых пунктов с несколькими индексами
    n = 0
    with gzip.open(st_path, 'rt', encoding='utf-8') as src, \
            gzip.open(OUT_STREETS, 'wt', encoding='utf-8', compresslevel=9) as out:
        for row in csv.DictReader(src):
            idx = (row.get('index') or '').strip()
            if idx not in multi:
                continue
            street = (row.get('address') or '').strip()
            if not street:
                continue
            out.write('%s\t%s\n' % (idx, street))
            n += 1
    print('  %s (%.1f МБ): улиц %d'
          % (OUT_STREETS.name, OUT_STREETS.stat().st_size / 1048576.0, n))
    return 0


def _release(rows):
    """Дата актуальности справочника — по самой свежей дате ввода индекса."""
    dates = [r.get('act_date') for r in rows if r.get('act_date')]
    return max(dates) if dates else ''


def main(argv):
    force = '--force' in argv
    try:
        return build(force=force)
    except Exception as ex:  # noqa: BLE001
        print('Не удалось собрать базу индексов: %s' % ex)
        return 1


if __name__ == '__main__':
    sys.exit(main(sys.argv))
