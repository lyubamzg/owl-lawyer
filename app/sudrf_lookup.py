"""Суд и почтовый индекс по адресу — ГАС РФ «Правосудие» (sudrf.ru).

Открытые источники, без регистрации и ключей:

  * адресные подсказки (ФИАС) — POST /suggestions/api/4_1/rs/suggest/address
  * территориальная подсудность  — GET  /api.php?opt=court_list&action=getCourtListByFiasCode

Что возвращает lookup():

  {'ok': True, 'postal_code': '123456', 'index_level': 'дом',
   'fias_id': '…', 'address': 'Примерная обл, Примерский р-н, …',
   'region': 'Московская обл', 'settlement': 'поселок Примерный',
   'courts': [{'kind': 'мировой', 'name': 'Судебный участок № 45 …',
               'address': '123456, …', 'code': '47MS0049'}, …],
   'ms': {...}, 'rs': {...}, 'notes': [...]}

Почтовый индекс берётся из ФИАС (тот же справочник, на котором работает
«Почта России»). Ответы кэшируются в ~/.cache/owl-lawyer/sudrf.json: при
повторной сборке сайт не опрашивается заново, а при отсутствии сети берётся
кэш. Если нет ни сети, ни кэша — документ собирается с тем судом, который
юрист указал вручную, а в отчёт падает предупреждение.

Проверка вручную:  python3 sudrf_lookup.py "Московская область, г. Пример, ул. Тестовая, д. 1"
"""

from __future__ import annotations

import json
import os
import re
import time
import urllib.request
from pathlib import Path

BASE_URL = 'https://sudrf.ru'
SUGGEST_URL = BASE_URL + '/suggestions/api/4_1/rs/suggest/address'
COURTS_URL = BASE_URL + '/api.php?opt=court_list&action=getCourtListByFiasCode'
REFERER = BASE_URL + '/index.php?id=300'
USER_AGENT = ('Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) '
              'Chrome/124.0.0.0 Safari/537.36 OwlLawyer/1.0')
TIMEOUT = 25
CACHE_FILE = Path(os.environ.get('OWL_LAWYER_CACHE') or
                  (Path.home() / '.cache' / 'owl-lawyer' / 'sudrf.json'))

# вид суда по коду ГАС «Правосудие»: 47MS0049 — мировой судья (судебный участок),
# 47RS0009 — районный, 47OS0000 — областной, 78AJ0002 — апелляционный,
# 78KJ0003 — кассационный, 63GV0001 — гарнизонный военный
KIND_BY_CODE = {
    'MS': 'мировой',
    'RS': 'районный',
    'OS': 'областной',
    'AJ': 'апелляционный',
    'KJ': 'кассационный',
    'GV': 'военный',
    'VO': 'военный',
    'VG': 'военный',
}

# «квартира» в подсказку адреса не принимается — из-за неё поиск не находит дом
KV_RE = re.compile(r'\bкв\.?\s*[^\s,;]+', re.IGNORECASE)
HOUSE_RE = re.compile(r'\bд\.?\s*([0-9]+[а-яА-Я]?(?:/[0-9]+)?)', re.IGNORECASE)
INDEX_RE = re.compile(r'(?<![\d-])(\d{6})(?![\d-])')
REGION_RE = re.compile(r'[А-ЯЁ][а-яё]+(?:[ -][А-ЯЁ][а-яё]+)*\s+'
                       r'(?:область|обл\.?|край|республика|респ\.?|округ)', re.IGNORECASE)


# --------------------------------------------------------------- HTTP ----
_LAST_CALL = [0.0]
MIN_PAUSE = 0.5      # пауза между обращениями к сайту — он не любит частых запросов
_FAILS = [0, 0.0]    # подряд неудачных запросов и время последней неудачи
FAIL_LIMIT = 3       # после стольких подряд неудач запросы не повторяем (сервис лежит)
FAIL_COOLDOWN = 120  # через сколько секунд попробовать снова


class ServiceUnavailable(RuntimeError):
    """Сайт ГАС «Правосудие» временно не отвечает."""


def _pause():
    delta = time.time() - _LAST_CALL[0]
    if delta < MIN_PAUSE:
        time.sleep(MIN_PAUSE - delta)
    _LAST_CALL[0] = time.time()


def _http_json(url, payload=None, retries=2, timeout=TIMEOUT):
    """GET/POST → JSON. Повторяет при сбое сети и при ответе «сервис временно недоступен»."""
    if _FAILS[0] >= FAIL_LIMIT and (time.time() - _FAILS[1]) < FAIL_COOLDOWN:
        raise ServiceUnavailable('ГАС «Правосудие» временно не отвечает — ждём и пробуем позже')
    data = json.dumps(payload, ensure_ascii=False).encode('utf-8') if payload is not None else None
    headers = {'User-Agent': USER_AGENT, 'Referer': REFERER, 'Accept': 'application/json'}
    if data is not None:
        headers['Content-Type'] = 'application/json'
    last = None
    for attempt in range(retries + 1):
        req = urllib.request.Request(url, data=data, headers=headers)
        try:
            _pause()
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read().decode('utf-8', 'replace')
            _FAILS[0] = 0
            return json.loads(raw)
        except Exception as ex:  # noqa: BLE001 — сеть может отвалиться в любой момент
            last = ex
            code = getattr(ex, 'code', None)
            if code in (429, 500, 502, 503, 504):
                _FAILS[0] += 1
                _FAILS[1] = time.time()
            if attempt < retries:
                time.sleep(1.0 * (attempt + 1))
    raise last if last else RuntimeError('неизвестная ошибка запроса')


# --------------------------------------------------------------- кэш ----
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


# --------------------------------------------------------- подготовка ----
def clean_query(addr):
    """Адрес → строка для поиска: без квартиры, без уже стоящего индекса."""
    a = KV_RE.sub(' ', str(addr or ''))
    a = re.sub(r'^\s*\d{6}\s*,\s*', '', a)
    a = re.sub(r'\s+', ' ', a).strip(' ,;.-')
    return a


def house_no(addr):
    """Номер дома из адреса («ул. Ленина, д. 33» → «33»)."""
    m = HOUSE_RE.search(str(addr or ''))
    return m.group(1) if m else ''


def region_hint(addr):
    """Область/край из самого адреса («Примерная область, …»)."""
    m = REGION_RE.search(str(addr or ''))
    return m.group(0).strip() if m else ''


def has_index(addr):
    """В адресе уже есть шестизначный индекс."""
    return bool(INDEX_RE.search(str(addr or '')))


STREET_RE = re.compile(r',\s*(?:ул|улица|просп|проспект|пер|переулок|бул|бульвар|ш|шоссе|наб|'
                        r'набережная|пр-т|прт|проезд|пл|площадь|мкр|мк-н|с|ст|тер|тракт)\.?\s+[^,]+',
                        re.IGNORECASE)
HOUSE_PART_RE = re.compile(r',\s*д\.?\s*[0-9]+[а-яА-Я]?(?:/[0-9]+)?', re.IGNORECASE)


def _strip_house(q):
    return HOUSE_PART_RE.sub('', q).strip(' ,;')


def _strip_street(q):
    return STREET_RE.sub('', q).strip(' ,;')


def _query_variants(query):
    """Варианты запроса для подсказок ФИАС.

    Сначала адрес как есть, затем — без «г. …»: подсказки ищутся по разным
    сочетаниям, и название города в запросе иногда мешает найти адрес.
    """
    out = [query]
    q2 = re.sub(r',\s*г\.\s*[^,]+', '', query).strip(' ,;')
    if q2 and q2 not in out:
        out.append(q2)
    return out


def locality_candidates(query):
    """Запросы для поиска индекса населённого пункта — от точного к общему."""
    q_nohouse = _strip_house(query)
    q_nostreet = _strip_street(q_nohouse)
    out = []
    for bound, q in (('settlement', q_nohouse), ('settlement', q_nostreet),
                     ('city', q_nostreet), ('city', q_nohouse)):
        if q and q not in [x[1] for x in out]:
            out.append((bound, q))
    return out


# ------------------------------------------------------------- поиск ----
def suggest(query, to_bound='house', count=8, region_guid=None):
    """Адресные подсказки ФИАС. to_bound: house | street | settlement | city."""
    payload = {'query': query,
               'from_bound': {'value': 'region'},
               'to_bound': {'value': to_bound},
               'count': count, 'lang': 'ru_RU'}
    if region_guid:
        payload['locations'] = [{'region_fias_id': region_guid}]
    data = _http_json(SUGGEST_URL, payload)
    return data.get('suggestions') or []


def _cached_suggest(cache, query, to_bound):
    key = f'sug|{to_bound}|{query}'
    if key in cache:
        return cache[key]
    items = suggest(query, to_bound=to_bound)
    cache[key] = items
    return items


def pick_suggestion(items, house=None):
    """Лучшее совпадение: нужный номер дома > точное совпадение (qc_geo=0) > дом."""
    if not items:
        return None

    def score(s):
        d = s.get('data') or {}
        sc = 0
        if house and str(d.get('house') or '').lower() == house.lower():
            sc += 100
        if str(d.get('qc_geo')) == '0':
            sc += 10
        if str(d.get('fias_level')) == '8':
            sc += 5
        if d.get('postal_code'):
            sc += 1
        return sc

    return max(items, key=score)


def courts_by_fias(fias_id):
    """Суды, подсудные адресу (ФИАС-GUID)."""
    if not fias_id:
        return []
    url = f'{COURTS_URL}&fiasId={fias_id}&page=1'
    data = _http_json(url)
    return data.get('list') or []


def _cached_courts(cache, fias_id):
    key = f'court|{fias_id}'
    if key in cache:
        return cache[key]
    items = courts_by_fias(fias_id)
    cache[key] = items
    return items


def court_kind(item):
    """Вид суда: мировой / районный / областной / …."""
    code = str(item.get('VNKOD') or '')
    kind = KIND_BY_CODE.get(code[2:4].upper(), '') if len(code) >= 6 else ''
    if kind:
        return kind
    name = str(item.get('ZNACHATR') or '').lower()
    if 'судебный участок' in name:
        return 'мировой'
    if 'районный суд' in name or 'городской суд' in name:
        return 'районный'
    if 'областной суд' in name or 'краевой суд' in name:
        return 'областной'
    return ''


def _norm_courts(items):
    out = []
    for it in items or []:
        name = str(it.get('ZNACHATR') or '').strip()
        if not name:
            continue
        out.append({'kind': court_kind(it),
                    'name': name,
                    'address': str(it.get('ADDRESS') or '').strip(),
                    'code': str(it.get('VNKOD') or '').strip()})
    return out


def court_header(court):
    """Суд → строка шапки («Мировому судье судебного участка № 45 …»)."""
    if not court:
        return ''
    name = court['name'].strip()
    if court.get('kind') == 'мировой':
        n = re.sub(r'^Судебный участок', 'судебного участка', name).strip()
        if n == name:
            n = f'судебного участка {name}'
        return f'Мировому судье {n}'
    return f'В {name}'


def pick_court(courts, kind):
    for c in courts:
        if c.get('kind') == kind:
            return c
    return None


# ----------------------------------------------------------- основной ----
def lookup(addr, region='', offline=False):
    """Адрес → почтовый индекс, ФИАС-GUID и список подсудных судов.

    region — подсказка области (если известна), помогает выбрать нужный
    населённый пункт при одноимённых улицах.
    """
    out = {'ok': False, 'query': '', 'postal_code': '', 'index_level': '',
           'fias_id': '', 'address': '', 'region': region or '',
           'settlement': '', 'city': '', 'courts': [], 'ms': None, 'rs': None,
           'notes': [], 'error': '', 'service_down': False}
    query = clean_query(addr)
    out['query'] = query
    if not query:
        out['error'] = 'пустой адрес'
        return out

    cache = _cache_read()
    changed = False
    house = house_no(addr)

    def _sug(bound, q=None):
        nonlocal changed
        q = q or query
        if offline:
            return cache.get(f'sug|{bound}|{q}', [])
        try:
            items = _cached_suggest(cache, q, bound)
            changed = True
            return items
        except Exception as ex:  # noqa: BLE001
            out['service_down'] = True
            out['notes'].append(f'поиск адреса не удался ({ex.__class__.__name__}: {ex}) — '
                                'взят кэш или ничего; суд и индекс в этом заявлении '
                                'оставлены как есть')
            return cache.get(f'sug|{bound}|{q}', [])

    # 1. дом — из него и индекс, и ФИАС-GUID для подсудности
    best = None
    for q in _query_variants(query):
        cand = pick_suggestion(_sug('house', q), house)
        if cand is None:
            continue
        if best is None:
            best = cand
        if (cand.get('data') or {}).get('fias_id'):
            best = cand
            break
    data = (best or {}).get('data') or {}
    if best:
        out['address'] = str(best.get('value') or '')
        out['fias_id'] = str(data.get('fias_id') or '')
        if data.get('postal_code'):
            out['postal_code'] = str(data['postal_code'])
            out['index_level'] = 'дом'
        if data.get('region_with_type'):
            out['region'] = str(data.get('region_with_type'))
        if data.get('settlement_with_type'):
            out['settlement'] = str(data.get('settlement_with_type'))
        elif data.get('city_with_type'):
            out['settlement'] = str(data.get('city_with_type'))
        out['city'] = str(data.get('city') or data.get('settlement') or '')
        if house and str(data.get('house') or '') != house:
            out['notes'].append(f'в ФИАС нет дома {house} по этому адресу — взята ближайшая '
                                f'запись ({data.get("house") or "?"}), проверьте адрес')
        if not out['fias_id'] and data.get('street_fias_id'):
            out['fias_id'] = str(data.get('street_fias_id'))

    # 1б. дома в ФИАС нет (или он не распознался) — ищем по улице:
    # по ней известны подсудные суды. Суд определяется по территории улицы,
    # поэтому в документе он помечается как требующий проверки.
    if not out['fias_id']:
        q_street = _strip_house(query)
        if q_street and q_street != query:
            pick = pick_suggestion(_sug('street', q_street))
            d = (pick or {}).get('data') or {}
            if d.get('street_fias_id'):
                out['fias_id'] = str(d['street_fias_id'])
                out['notes'].append('в ФИАС нет дома — подсудность определена по улице, '
                                    'проверьте участок суда')
                if d.get('postal_code') and not out['postal_code']:
                    out['postal_code'] = str(d['postal_code'])
                    out['index_level'] = 'улица'

    # 2. индекс населённого пункта — если по дому индекса нет    if not out['postal_code']:
        for bound, q2 in locality_candidates(query):
            items = _sug(bound, q2)
            pick = pick_suggestion(items)
            d = (pick or {}).get('data') or {}
            if d.get('postal_code'):
                out['postal_code'] = str(d['postal_code'])
                out['index_level'] = ('населённый пункт' if bound == 'settlement' else 'город')
                if not out['region'] and d.get('region_with_type'):
                    out['region'] = str(d.get('region_with_type'))
                if not out['settlement']:
                    out['settlement'] = str(d.get('settlement_with_type')
                                            or d.get('city_with_type') or '')
                break
        if out['postal_code']:
            out['notes'].append('индекс взят по населённому пункту (по дому в ФИАС его нет)')

    # 3. суды по ФИАС-GUID
    if out['fias_id']:
        try:
            items = cache.get(f'court|{out["fias_id"]}', []) if offline else _cached_courts(cache, out['fias_id'])
            if not offline:
                changed = True
            out['courts'] = _norm_courts(items)
        except Exception as ex:  # noqa: BLE001
            out['notes'].append(f'список судов не получен ({ex.__class__.__name__})')
    else:
        if not out['service_down']:
            out['notes'].append('адрес не найден в ФИАС — подсудность определить не удалось')

    out['ms'] = pick_court(out['courts'], 'мировой')
    out['rs'] = pick_court(out['courts'], 'районный')
    out['ok'] = bool(out['fias_id'] or out['postal_code'])
    if changed:
        _cache_write(cache)
    return out


def court_for_addr(addr, typ='СП', cena=None, threshold=50000, region='', offline=False,
                   res=None):
    """Адрес → строки шапки суда по виду документа.

    СП (судебный приказ) — всегда мировой судья (ст. 121 ГПК РФ).
    ИСК — районный суд; если цена иска не больше порога (по умолчанию 50 000 ₽,
    ст. 23 ГПК РФ для иных имущественных споров) — мировой судья.

    Возвращает (строки шапки, пояснение, результат lookup).
    """
    if res is None:
        res = lookup(addr, region=region, offline=offline)
    lines, note, court = [], '', None
    if typ == 'СП':
        court = res.get('ms') or res.get('rs')
        if court:
            lines = [court_header(court)]
            note = f'мировой судья ({court["name"]})'
        else:
            note = 'подсудный суд по адресу не найден'
    else:
        price = float(cena or 0)
        if price <= float(threshold or 0) and res.get('ms'):
            court = res['ms']
            lines = [court_header(court)]
            note = f'мировой судья ({court["name"]}), цена иска ≤ {threshold:.0f} ₽'
        elif res.get('rs'):
            court = res['rs']
            lines = [court_header(court)]
            note = f'районный суд ({court["name"]})'
        elif res.get('ms'):
            court = res['ms']
            lines = [court_header(court)]
            note = f'мировой судья ({court["name"]}) — районный не найден'
        else:
            note = 'подсудный суд по адресу не найден'
    if lines and court and court.get('address'):
        lines.append(f'Адрес суда: {court["address"]}')
    return lines, note, res


def add_index(addr, index):
    """Дописать почтовый индекс в адрес (в начало, как в почтовом адресе)."""
    a = str(addr or '').strip()
    if not a or not index:
        return a
    if has_index(a):
        return a
    return f'{index}, {a}'


def _cli(argv):
    addr = ' '.join(argv).strip()
    if not addr:
        print('Укажите адрес: python3 sudrf_lookup.py "Московская область, г. Пример, ул. Тестовая, д. 1"')
        return 1
    res = lookup(addr)
    print(f'Запрос:  {res["query"]}')
    print(f'Индекс:  {res["postal_code"] or "не найден"} ({res["index_level"] or "—"})')
    print(f'ФИАС:    {res["address"] or "не найден"}')
    for n in res['notes']:
        print(f'  ⚠ {n}')
    print('Суды по адресу:')
    for c in res['courts']:
        print(f'  [{c["kind"] or "—":<12}] {c["name"]}')
        if c['address']:
            print(f'                 {c["address"]}')
    if not res['courts']:
        print('  (не найдены)')
    return 0


if __name__ == '__main__':
    import sys
    raise SystemExit(_cli(sys.argv[1:]))
