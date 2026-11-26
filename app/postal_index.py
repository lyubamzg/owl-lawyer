#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Почтовый индекс по адресу — локальная база индексов РФ и ФИАС.

Порядок поиска:
  1. индекс уже стоит в адресе — возвращаем его;
  2. локальная база (data/postal_base.json.gz) — эталонный справочник
     почтовых индексов ОПС: все индексы РФ, ищем по населённому пункту,
     а в городах, где индексов несколько, — по улице;
  3. ГАС «Правосудие» (sudrf_lookup) — индекс по дому из ФИАС;
  4. сайт «Почты России» (postrf_index) — запасной источник.

Локальная база работает без сети — поэтому приложение может подтянуть индекс
даже тогда, когда сайты не отвечают. Собрать или обновить базу:

    python3 build_postal_base.py

Проверка вручную:
    python3 postal_index.py "Примерная область, г. Пример, ул. Спортивная, д. 1"
"""
from __future__ import annotations

import gzip
import json
import re
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
DATA = BASE / 'data'
BASE_FILE = DATA / 'postal_base.json.gz'
STREETS_FILE = DATA / 'postal_streets.txt.gz'

# «123456», но не «2026» и не часть номера телефона
INDEX_RE = re.compile(r'(?<![\d-])(\d{6})(?![\d-])')
WORD_RE = re.compile(r'[0-9A-Za-zА-Яа-яЁё]+')

# служебные слова в начале названия населённого пункта
TYPE_WORDS = {
    'г', 'гор', 'город', 'пос', 'поселок', 'посёлок', 'село', 'с', 'сп', 'ст',
    'стца', 'станица', 'д', 'дер', 'деревня', 'хут', 'хуттор', 'сц', 'мкр',
    'мкн', 'микрорайон', 'зимник', 'курорт', 'пгт', 'п', 'рабп', 'имени', 'им',
}
# слова, после которых идёт название улицы
STREET_WORDS = {
    'ул', 'улица', 'просп', 'проспект', 'прт', 'пр', 'пер', 'перул', 'бул',
    'бульвар', 'ш', 'шоссе', 'наб', 'набережная', 'пл', 'площадь', 'тер',
    'тракт', 'проезд', 'спуск', 'тупик', 'линия', 'аллея', 'мкр', 'мкн',
    'мк', 'кв', 'квартал',
}
# сокращения типа улицы в справочнике: «б-р. Покровский» → «покровский»
STREET_TYPE_TOKENS = STREET_WORDS | {
    'б', 'р', 'бул', 'прт', 'пр', 'наб', 'ш', 'пер', 'пл', 'мк', 'мкр',
    'туп', 'сп', 'тракт', 'тер', 'лин', 'ал', 'проезд', 'аллея', 'набережная',
    'площадь', 'улица', 'проспект', 'шоссе', 'переулок', 'бульвар',
    'микрорайон', 'квартал', 'линия', 'спуск', 'тупик', 'ж', 'тер',
}
# только сокращения (полные слова «улица», «проспект» оставляем)
STREET_ABBR = {
    'б', 'р', 'бул', 'прт', 'пр', 'наб', 'ш', 'пер', 'пл', 'мк', 'мкр',
    'туп', 'сп', 'тракт', 'тер', 'лин', 'ал', 'ул', 'ж', 'вл', 'мкн',
}
# сокращения регионов и городов
PLACE_ALIASES = {
    'спб': 'санкт петербург',
    'санкт петербург': 'санкт петербург',
    'мск': 'москва',
    'н новгород': 'нижний новгород',
    'нижний новгород': 'нижний новгород',
    'екатеринбург': 'екатеринбург',
    'свердловск': 'екатеринбург',
    'нижний новгород': 'нижний новгород',
    'петербург': 'санкт петербург',
    'ленинград': 'санкт петербург',
}

_BASE = None          # локальная база (грузится при первом обращении)
_STREETS = None       # карта «улица → индекс» (грузится по требованию)

# сколько индексов у населённого пункта — уже «много»,
# индекс населённого пункта не говорит о доме и требует уточнения
MANY_INDEXES = 10


def norm(text):
    """«посёлок им. Кирова» → «поселок им кирова»."""
    s = str(text or '').strip().lower().replace('ё', 'е')
    s = re.sub(r'[^0-9a-zа-я]+', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def _base():
    """Локальная база индексов (None — если её нет)."""
    global _BASE
    if _BASE is None:
        _BASE = {}
        if BASE_FILE.is_file():
            try:
                with gzip.open(BASE_FILE, 'rt', encoding='utf-8') as f:
                    _BASE = json.load(f)
            except (OSError, ValueError):
                _BASE = {}
    return _BASE or None


def _streets():
    """Карта «индекс → {улица: индекс}» (None — если файла нет)."""
    global _STREETS
    if _STREETS is None:
        _STREETS = {}
        if STREETS_FILE.is_file():
            try:
                with gzip.open(STREETS_FILE, 'rt', encoding='utf-8') as f:
                    for line in f:
                        idx, _, street = line.rstrip('\n').partition('\t')
                        if not idx or not street:
                            continue
                        _STREETS.setdefault(idx, {})[_street_norm(street)] = idx
            except OSError:
                _STREETS = {}
    return _STREETS or None


def has_index(addr):
    """Есть ли в адресе шестизначный почтовый индекс."""
    return bool(INDEX_RE.search(str(addr or '')))


def add_index(addr, index):
    """Дописать индекс в начало адреса (как в почтовом адресе)."""
    s = str(addr or '').strip()
    if not index or has_index(s):
        return s
    return '%s, %s' % (index, s)


def _tokens(text):
    """[(слово, начало, конец)]: слова адреса в нормализованном виде."""
    out = []
    for m in WORD_RE.finditer(text):
        w = norm(m.group(0))
        if w:
            out.append((w, m.start(), m.end()))
    return out


def _region_first(base):
    """Первое слово названия региона → ключи регионов (для «Ленинградская обл»)."""
    cache = getattr(_region_first, 'cache', None)
    if cache is not None:
        return cache
    first = {}
    for entries in (base.get('settlements') or {}).values():
        for entry in entries:
            key = norm(entry[0])
            if not key:
                continue
            head = key.split(' ')[0]
            first.setdefault(head, set()).add(key)
    cache = {k: sorted(v) for k, v in first.items()}
    _region_first.cache = cache
    return cache


def _region_of(text, base):
    """Регион адреса: (ключ, индекс, (начало, конец) в токенах)."""
    reg = base.get('regions') or {}
    toks = _tokens(text)
    for n in (3, 2, 1):
        for i in range(len(toks) - n + 1):
            key = ' '.join(t[0] for t in toks[i:i + n])
            if key in reg:
                return key, reg[key], (i, i + n)
    # «Ленинградская обл» — регион по первому слову
    first = _region_first(base)
    for i, tok in enumerate(toks):
        keys = first.get(tok[0])
        if keys and len(keys) == 1:
            key = keys[0]
            return key, reg.get(key, ''), (i, i + 1)
    return '', '', (0, 0)


def _settlement_of(text, base, reg_key='', skip=()):
    """Населённый пункт адреса: (ключ, варианты, позиция токенов).

    skip — диапазон токенов, занятый регионом: названия региона мы за
    населённый пункт не принимаем (кроме случая «Москва», где одно и то же
    слово — и регион, и город).
    """
    settle = base.get('settlements') or {}
    reg_keys = set((base.get('regions') or {}).keys())
    toks = _tokens(text)
    best = None
    for n in (4, 3, 2, 1):
        for i in range(len(toks) - n + 1):
            key = ' '.join(t[0] for t in toks[i:i + n])
            if key not in settle:
                alias = PLACE_ALIASES.get(key)
                if not alias or alias not in settle:
                    continue
                key = alias
            if key in reg_keys and reg_key and key != reg_key:
                continue           # это название региона, а не посёлка
            if i < skip[1] and skip[0] < i + n and key != reg_key:
                continue           # слово региона — не населённый пункт
            # «г. Пример» — населённый пункт идёт сразу за типом
            marked = i > 0 and toks[i - 1][0] in TYPE_WORDS
            score = (2 if marked else 0, n, -i)
            if best is None or score > best[0]:
                best = (score, key, settle[key], (i, i + n))
        if best:
            break
    if not best:
        return '', [], (0, 0)
    return best[1], best[2], best[3]


def _street_of(text):
    """Улица адреса (после «ул.», «просп.» и т. п.)."""
    toks = _tokens(text)
    for i, tok in enumerate(toks):
        if tok[0] in STREET_WORDS:
            part = [t[0] for t in toks[i + 1:i + 4]]
            part = [p for p in part if p not in TYPE_WORDS]
            if part:
                return ' '.join(part)
    return ''


def _street_norm(text):
    """«б-р. Покровский» → «покровский»: убираем сокращения типа улицы."""
    parts = [p for p in norm(text).split(' ') if p]
    parts = [p for p in parts if p not in STREET_ABBR]
    return ' '.join(parts)


def _pick_region(entries, region_key, region_index):
    """Выбрать вариант населённого пункта по региону адреса."""
    if region_key:
        for entry in entries:
            if norm(entry[0]) == region_key:
                return entry[0], entry[1], entry[2]
    return entries[0][0], entries[0][1], entries[0][2]


def from_base(addr):
    """Индекс по локальной базе: (индекс, уровень, неоднозначно)."""
    base = _base()
    if not base:
        return '', '', False
    text = str(addr or '')
    region_key, region_index, reg_span = _region_of(text, base)
    key, entries, _span = _settlement_of(text, base, region_key, skip=reg_span)
    if entries:
        region, area, idx = _pick_region(entries, region_key, region_index)
        count = _count_of(entries, idx)
        if count > 1:
            street = _street_of(text)
            if street:
                found = _street_index(idx, entries, street)
                if found:
                    return found, 'улица', False
        # в большом городе индекс населённого пункта неточен — уточняем по дому
        return idx, 'населённый пункт', count > MANY_INDEXES
    if region_index:
        return region_index, 'регион', False
    return '', '', False


def _count_of(entries, idx):
    """Сколько индексов у населённого пункта."""
    for region, area, index, *rest in entries:
        if index == idx:
            return rest[0] if rest else 1
    return 1


def _has_many_indexes(idx):
    streets = _streets()
    return bool(streets and idx in streets)


def _street_index(idx, entries, street):
    """Индекс по улице среди индексов населённого пункта."""
    streets = _streets()
    if not streets:
        return ''
    candidates = [idx] + [e[2] for e in entries]
    key = _street_norm(street)
    for cand in candidates:
        table = streets.get(cand)
        if not table:
            continue
        if key in table:
            return table[key]
        for s, val in table.items():
            sk = _street_norm(s)
            if sk == key or sk.endswith(' ' + key) or sk.startswith(key + ' '):
                return val
    return ''


def postal_code(addr, online=True):
    """Почтовый индекс по адресу.

    Возвращает (индекс, уровень, источник). Уровень: «в адресе», «дом»,
    «улица», «населённый пункт», «регион».
    """
    s = str(addr or '').strip()
    if not s:
        return '', '', 'адрес пустой'
    m = INDEX_RE.search(s)
    if m:
        return m.group(1), 'в адресе', 'индекс уже стоит в адресе'

    idx, level, vague = from_base(s)
    base_note = ''
    if idx:
        base = _base() or {}
        base_note = 'локальная база индексов ОПС (%s)' % base.get('release', '?')
        if not vague:
            return idx, level, base_note

    # в большом городе индекс населённого пункта неточен — уточняем по дому
    if online:
        try:
            import sudrf_lookup
            res = sudrf_lookup.lookup(s)
            if res.get('postal_code'):
                return (res['postal_code'], res.get('index_level') or 'дом',
                        'по ФИАС (ГАС «Правосудие»)')
        except Exception:  # noqa: BLE001 — сеть недоступна, идём к запасному
            pass
        try:
            import postrf_index
            idx2, level2, note2 = postrf_index.postal_code(s)
            if idx2:
                return idx2, level2, note2
        except Exception:  # noqa: BLE001
            pass
    if idx:
        return idx, level, base_note + ' — уточнить по дому не удалось'
    return '', '', 'индекс не найден'


def ensure_index(addr, online=True):
    """Адрес с почтовым индексом: (адрес, индекс, источник).

    Если индекс в адресе уже есть — адрес не меняется.
    """
    s = str(addr or '').strip()
    if not s:
        return s, '', 'адрес пустой'
    if has_index(s):
        m = INDEX_RE.search(s)
        return s, m.group(1), 'индекс уже стоит в адресе'
    idx, level, note = postal_code(s, online=online)
    if not idx:
        return s, '', note
    return add_index(s, idx), idx, '%s (%s)' % (note, level)


def index_info(idx):
    """Что за индекс: название объекта по справочнику ОПС."""
    base = _base()
    if not base:
        return ''
    return (base.get('indexes') or {}).get(str(idx or ''), '')


def _cli(argv):
    for a in argv[1:]:
        idx, level, note = postal_code(a)
        print('%s\n  индекс: %s\n  уровень: %s\n  источник: %s\n  название: %s'
              % (a, idx or '—', level or '—', note, index_info(idx) or '—'))
    return 0


if __name__ == '__main__':
    sys.exit(_cli(sys.argv))
