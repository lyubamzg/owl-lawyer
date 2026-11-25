"""Почтовый индекс по адресу — сайт «Почты России» (pochta.ru).

Официальный сайт Почты России отдаёт адресные подсказки (тот же справочник
ФИАС, что и у ГАС «Правосудие») вместе с почтовым индексом дома.
Регистрация и ключи не нужны — используется тот же публичный адрес подсказок,
что работает на сайте.

Запасной источник нужен потому, что:
  * www.postrf.ru с этой машины недоступен — сеть перехватывает TLS и отдаёт
    чужой сертификат (а не сертификат Почты России);
  * ГАС «Правосудие» (sudrf.ru) периодически отвечает «503 Service Unavailable»
    — тогда индекс не подставлялся вообще.

Запрос (тот же, что делает сам сайт):
  POST https://www.pochta.ru/api/suggestions/api/v2/suggestion.find-addresses
  {"query": "...", "limit": 5, "fromBound": "REGION",
   "language": "RUSSIAN", "mailDirection": "ALL"}

Ответ — список подсказок, у каждой «postalCode» (индекс) и элементы адреса
(область, район, населённый пункт, улица, дом, квартира).

Ответы кэшируются в ~/.cache/owl-lawyer/postrf.json (ключ — нормализованный
адрес), поэтому при повторной сборке сайт не опрашивается заново.

Проверка вручную:
  python3 postrf_index.py "Примерная область, п. Пример, ул. Спортивная, д. 13"
"""

from __future__ import annotations

import json
import os
import re
import time
import urllib.request
from pathlib import Path

SUGGEST_URL = 'https://www.pochta.ru/api/suggestions/api/v2/suggestion.find-addresses'
REFERER = 'https://www.pochta.ru/'
USER_AGENT = ('Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) '
              'Chrome/124.0.0.0 Safari/537.36 OwlLawyer/1.0')
TIMEOUT = 25
CACHE_FILE = Path(os.environ.get('OWL_LAWYER_CACHE_POSTRF') or
                  (Path.home() / '.cache' / 'owl-lawyer' / 'postrf.json'))

MIN_PAUSE = 0.4       # пауза между обращениями к сайту
_LAST_CALL = [0.0]

# индекс уже стоит в адресе
INDEX_RE = re.compile(r'(?<![\d-])(\d{6})(?![\d-])')
HOUSE_RE = re.compile(r'\bд\.?\s*([0-9]+[а-яА-Я]?(?:/[0-9]+)?)', re.IGNORECASE)
KV_RE = re.compile(r'\bкв\.?\s*[^,;]*', re.IGNORECASE)
# «Примерское поселение, …» — Почта России такого МО не знает, отбрасываем
MO_PREFIX_RE = re.compile(r'^\s*[^,]*поселение\s*,\s*', re.IGNORECASE)
# тип улицы в начале куска: «ул. Спортивная», «просп. Комсомольский»
STREET_TYPE_RE = re.compile(
    r'^\s*(?:ул|улица|просп|проспект|пр-т|прт|пер|перул|бул|бульвар|ш|шоссе|наб|'
    r'набережная|пл|площадь|мкр|мк-н|тер|тракт|проезд|спуск|тупик)\.?\s+',
    re.IGNORECASE)
# населённый пункт: «п. Пример», «г. Пример», «пос. Клястино»
NP_TYPE_RE = re.compile(r'^\s*(?:п|пг|г|пос|дер|с|сп|х|к)\.?\s+', re.IGNORECASE)


def _pause():
    delta = time.time() - _LAST_CALL[0]
    if delta < MIN_PAUSE:
        time.sleep(MIN_PAUSE - delta)
    _LAST_CALL[0] = time.time()


def has_index(addr):
    """В адресе уже есть шестизначный индекс."""
    return bool(INDEX_RE.search(str(addr or '')))


def house_no(addr):
    """Номер дома из адреса («ул. Ленина, д. 33» → «33»)."""
    m = HOUSE_RE.search(str(addr or ''))
    return m.group(1) if m else ''


def _parts(addr):
    """Адрес → (населённый пункт, улица, дом) — по тому, что видно в строке."""
    a = str(addr or '')
    a = MO_PREFIX_RE.sub('', a)
    a = KV_RE.sub(' ', a)
    a = re.sub(r'\s+', ' ', a).strip(' ,;')
    house = house_no(a)
    a = HOUSE_RE.sub(' ', a)
    chunks = [c.strip(' ,;.') for c in a.split(',') if c.strip(' ,;')]
    settlement = ''
    street = ''
    rest = []
    for c in chunks:
        if re.fullmatch(r'\d+[а-яА-Я]?(?:/\d+)?', c):    # номер дома
            continue
        if NP_TYPE_RE.match(c) and not settlement:
            settlement = NP_TYPE_RE.sub('', c).strip(' ,;.')
            continue
        if STREET_TYPE_RE.match(c) and not street:
            street = STREET_TYPE_RE.sub('', c).strip(' ,;.')
            continue
        rest.append(c)
    if not settlement and rest:
        settlement = rest.pop(0)
    if not street and rest:
        street = rest.pop(0)
    return settlement, street, house


def query_variants(addr):
    """Варианты запроса — от самого точного к самому общему."""
    settlement, street, house = _parts(addr)
    out = []
    if settlement and street and house:
        out.append(f'{settlement}, {street}, {house}')
        out.append(f'{street}, {house}, {settlement}')
        out.append(f'{settlement}, {street}')
    if settlement and street:
        out.append(f'{settlement}, {street}')
    if street and house:
        out.append(f'{street}, {house}')
    if settlement:
        out.append(settlement)
    seen, res = set(), []
    for v in out:
        if v and v.lower() not in seen:
            seen.add(v.lower())
            res.append(v)
    return res


def _suggest(query, limit=6):
    """Подсказки «Почты России» по адресу (список)."""
    payload = {'query': query, 'limit': limit, 'fromBound': 'REGION',
               'language': 'RUSSIAN', 'mailDirection': 'ALL'}
    data = json.dumps(payload, ensure_ascii=False).encode('utf-8')
    headers = {'User-Agent': USER_AGENT, 'Referer': REFERER,
               'Accept': 'application/json', 'Content-Type': 'application/json'}
    req = urllib.request.Request(SUGGEST_URL, data=data, headers=headers)
    _pause()
    with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
        body = resp.read().decode('utf-8', 'replace')
    try:
        out = json.loads(body)
    except Exception:  # noqa: BLE001
        return []
    return out if isinstance(out, list) else []


def _el(item, code):
    """Значение элемента подсказки по коду ('S' — улица, 'N' — дом, …)."""
    for e in item.get('elements') or []:
        if e.get('content') == code:
            return str(e.get('value') or '')
    return ''


def pick(items, house='', street='', settlement=''):
    """Лучшая подсказка: нужный дом > нужная улица > населённый пункт."""
    def score(item):
        sc = 0
        if house and str(item.get('house') or '').lower() == house.lower():
            sc += 100
        if street and _el(item, 'S').lower() == street.lower():
            sc += 30
        if settlement and _el(item, 'P').lower() == settlement.lower():
            sc += 10
        if item.get('postalCode'):
            sc += 5
        if _el(item, 'F'):          # квартира — подсказка до уровня квартиры
            sc += 2
        return sc
    best = None
    for item in items or []:
        if not item.get('postalCode'):
            continue
        if best is None or score(item) > score(best):
            best = item
    return best


def _cache_read():
    try:
        with open(CACHE_FILE, encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:  # noqa: BLE001
        return {}


def _cache_write(cache):
    try:
        CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        tmp = CACHE_FILE.with_suffix('.tmp')
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(cache, f, ensure_ascii=False)
        tmp.replace(CACHE_FILE)
    except Exception:  # noqa: BLE001
        pass


def postal_code(addr):
    """Адрес → почтовый индекс с сайта «Почты России».

    Возвращает (индекс, уровень, пояснение); индекс '' — если не нашлось.
    """
    if not addr or has_index(addr):
        return '', '', 'в адресе уже есть индекс'
    cache = _cache_read()
    settlement, street, house = _parts(addr)
    changed = False
    for q in query_variants(addr):
        key = f'{settlement}|{street}|{house}|{q}'.lower()
        if key in cache:
            got = cache[key]
            if got:
                return got.get('postal_code', ''), got.get('level', ''), 'из кэша'
            continue
        try:
            items = _suggest(q)
        except Exception as ex:  # noqa: BLE001 — сеть может отвалиться
            return '', '', f'«Почта России» не ответила ({ex.__class__.__name__})'
        best = pick(items, house=house, street=street, settlement=settlement)
        if best:
            got = {'postal_code': str(best.get('postalCode') or ''),
                   'level': 'дом' if best.get('house') else
                            ('улица' if _el(best, 'S') else 'населённый пункт'),
                   'address': str(best.get('address') or best.get('addressGuid') or '')}
            cache[key] = got
            changed = True
            return got['postal_code'], got['level'], 'с сайта «Почта России»'
        cache[key] = {}
        changed = True
    if changed:
        _cache_write(cache)
    return '', '', '«Почта России» не нашла адрес'


def _cli(argv):
    addr = ' '.join(argv).strip()
    if not addr:
        print('Укажите адрес: python3 postrf_index.py "Примерная область, п. Пример, '
              'ул. Спортивная, д. 13"')
        return 1
    idx, level, note = postal_code(addr)
    print(f'Запрос:  {addr}')
    print(f'Индекс:  {idx or "не найден"} ({level or "—"})')
    print(f'Источник: {note}')
    return 0 if idx else 1


if __name__ == '__main__':
    import sys
    sys.exit(_cli(sys.argv[1:]))
