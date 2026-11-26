# -*- coding: utf-8 -*-
"""Локальная база судебных участков — резервный источник подсудности.

Живая ГАС «Правосудие» (sudrf.ru) — основной источник, но её сервис подсказок
периодически не отвечает, и тогда в шапке документа остаётся суд, указанный
руками (часто это один и тот же участок для всего района — а он неверен).

Здесь лежит список всех мировых судебных участков России (код ВНКО, официальное
название, районный суд, описание территории) — data/sudebnye_uchastki.json.
По адресу ответчика подбирается участок, которому этот адрес подсуден:

  «Примерная область, Примерский район, МО Примерское городское поселение,
   г. Пример, ул. Примерная, д. 6»
      → судебный участок № 72 Примерского муниципального района

Как ищется: по региону отбираются участки этого субъекта, затем считаются
совпадения «населённый пункт / поселение / микрорайон» и «улица» с описанием
территории участка. Поселения весят в 10 раз больше улиц: иначе участки
городских микрорайонов (где в описании перечислены улицы) перебивают
поселения.

Если участок не найден или найдено несколько равновероятных — суд не
подставляется: лучше пустое поле и предупреждение, чем неверный участок.

Файлы базы:
  data/sudebnye_uchastki.json — база (собирается скриптом
  update_sudebnye_uchastki.py). Формат: {"participants": [ {...} ]}
"""

from __future__ import annotations

import difflib
import json
import re
from pathlib import Path

BASE_FILE = 'sudebnye_uchastki.json'
SEARCH_DIRS = [
    Path(__file__).resolve().parent / 'data',
    Path.home() / 'Desktop' / 'OwlLawyer' / 'data',
    Path.home() / '.local' / 'share' / 'owl-lawyer' / 'data',
]

REGION_RE = re.compile(r'[А-ЯЁ][а-яё]+(?:[ -][А-ЯЁ][а-яё]+)*\s+'
                       r'(?:область|обл\.?|край|республика|респ\.?|автономный округ|округ)', re.I)
STREET_RE = re.compile(r'(?:^|,)\s*(?:ул|улица|просп|проспект|пр-т|прт|пер|переулок|бул|бульвар|'
                       r'ш|шоссе|наб|набережная|проезд|пр-д|пл|площадь|мкр|мк-н|тер|тракт)\.?\s+',
                       re.I)
TOKEN_RE = re.compile(r'[а-яёa-z]{4,}')

# слова, которые ничего не различают
IGNORE = {
    'область', 'области', 'край', 'края', 'республика', 'республике', 'округ', 'автономный',
    'поселение', 'поселения', 'поселению', 'городское', 'городское', 'сельское', 'сельское',
    'сельские', 'город', 'города', 'деревня', 'деревни', 'поселок', 'поселки', 'село', 'села',
    'улица', 'улицы', 'переулок', 'переулки', 'проспект', 'шоссе', 'проезд', 'проезды',
    'бульвар', 'площадь', 'набережная', 'микрорайон', 'микрорайоны', 'территория', 'территории',
    'лесничество', 'заповедник', 'дом', 'дома', 'квартира', 'квартиры', 'строение', 'корпус',
    'жилой', 'жилая', 'помещение', 'здание', 'центр', 'станция', 'школа', 'сад', 'сквер',
    'квартал', 'кварталы', 'часть', 'сторона', 'входящих', 'принадлежащем', 'областей',
    'районы', 'муниципальный', 'муниципальные', 'посёлок', 'сельсовет', 'сельсовета',
}

_cache = {}
LOCAL_FILE = 'sudebnye_uchastki.local.json'


# ------------------------------------------------- запомненные адреса ----
def local_path():
    for d in SEARCH_DIRS:
        if (d / BASE_FILE).exists():
            return d / LOCAL_FILE
    return None


def load_local():
    """Запомненные сопоставления «населённый пункт → участок» (из живой ГАС)."""
    if 'local' in _cache:
        return _cache['local']
    data = []
    p = local_path()
    if p and p.exists():
        try:
            raw = json.loads(p.read_text(encoding='utf-8'))
            data = raw.get('remembered') or []
        except Exception:  # noqa: BLE001
            data = []
    _cache['local'] = data
    return data


def remember(addr, court):
    """Запомнить: адрес ответчика → участок (из живой ГАС «Правосудие»).

    Так база растёт по мере работы: если ГАС ответила, участок для этого
    населённого пункта известен точно и сохраняется локально.
    """
    if not addr or not court:
        return
    settle, _ = addr_tokens(addr)
    if not settle:
        return
    items = load_local()
    key = ' '.join(sorted(set(settle)))
    for it in items:
        if it.get('key') == key:
            return
    items.append({'key': key, 'tokens': sorted(set(settle)), 'court': court})
    p = local_path()
    if not p:
        return
    try:
        p.write_text(json.dumps({'remembered': items}, ensure_ascii=False, indent=1),
                     encoding='utf-8')
        _cache['local'] = items
    except Exception:  # noqa: BLE001
        pass


def _remembered(addr):
    """Участок из запомненных сопоставлений (точное совпадение по нп/поселению)."""
    settle, _ = addr_tokens(addr)
    if not settle:
        return None
    want = set(settle)
    best = None
    for it in load_local():
        toks = set(it.get('tokens') or [])
        common = want & toks
        if not common:
            continue
        if best is None or len(common) > len(best[1]):
            best = (it, common)
    return best[0] if best else None


# ------------------------------------------------------------- загрузка ----
def base_path():
    for d in SEARCH_DIRS:
        p = d / BASE_FILE
        if p.exists():
            return p
    return None


def load():
    """База участков (с кэшем в памяти). Пустой словарь, если базы нет."""
    if 'data' in _cache:
        return _cache['data']
    data = {'participants': [], 'regions': {}, 'by_region': {}}
    p = base_path()
    if p:
        try:
            raw = json.loads(p.read_text(encoding='utf-8'))
            parts = raw.get('participants') or []
            data['participants'] = parts
            for part in parts:
                rc = part.get('region_code') or ''
                rn = part.get('region') or ''
                if rc:
                    data['regions'].setdefault(rc, rn)
                data['by_region'].setdefault(rc, []).append(part)
        except Exception:  # noqa: BLE001 — база не должна ломать сбор документов
            data = {'participants': [], 'regions': {}, 'by_region': {}}
    _cache['data'] = data
    return data


def norm(s):
    s = re.sub(r'\(.*?\)', ' ', str(s or ''))
    s = s.replace('ё', 'е')
    s = re.sub(r'[^а-яa-z0-9]+', ' ', s.lower())
    return re.sub(r'\s+', ' ', s).strip()


def region_key(name):
    n = norm(name)
    n = re.sub(r'\b(республика|республике|область|области|край|края|округ|автономный)\b', ' ', n)
    n = re.sub(r'\s+-\s+.*$', ' ', n)
    return re.sub(r'\s+', ' ', n).strip()


def region_code(addr, base=None):
    """Код региона (ВНКО) по названию субъекта в адресе."""
    base = base or load()
    text = str(addr or '')
    m = REGION_RE.search(text)
    name = m.group(0).strip() if m else ''
    key = region_key(name)
    if key:
        for rc, rn in base['regions'].items():
            if region_key(rn) == key:
                return rc, rn
        best = difflib.get_close_matches(key, [region_key(x) for x in base['regions'].values()],
                                         n=1, cutoff=0.8)
        if best:
            for rc, rn in base['regions'].items():
                if region_key(rn) == best[0]:
                    return rc, rn
    # «Республика Татарстан», «г. Москва», «г. Санкт-Петербург» — регион
    # стоит в начале или назван по городу: ищем название региона в адресе
    atext = norm(text)
    cands = []
    for rc, rn in base['regions'].items():
        rk = region_key(rn)
        for variant in (rk, rk.replace('город ', '')):
            if len(variant) < 5:
                continue
            if variant in atext:
                cands.append((len(variant), rc, rn))
                break
    if cands:
        cands.sort(reverse=True)
        return cands[0][1], cands[0][2]
    return '', name


# ------------------------------------------------------------- разбор ------
def _stem(word):
    """Основа слова для сравнения: «Рябово» ↔ «Рябовское», «Ульяновка» ↔ «Ульяновское»."""
    w = norm(word)
    if len(w) > 6:
        return w[:6]
    if len(w) > 4:
        return w[:5]
    return w


def addr_tokens(addr):
    """Адрес → (населённые пункты/поселения, улицы)."""
    a = re.sub(r'^\s*\d{6}\s*,\s*', '', str(addr or ''))
    a = re.sub(r'\bкв\.?\s*[^,;]*', ' ', a, flags=re.I)
    parts = [p.strip(' ,;.') for p in a.split(',') if p.strip(' ,;')]
    settle, streets = [], []
    for p in parts:
        if STREET_RE.search(' ' + p + ' ') or STREET_RE.match(p):
            streets.append(p)
        else:
            settle.append(p)
    def words(seq):
        out = []
        for chunk in seq:
            for w in TOKEN_RE.findall(norm(chunk)):
                if w in IGNORE:
                    continue
                out.append(w)
        return out
    return words(settle), words(streets)


def _territory_parts(desc):
    """Описание территории → (часть про поселения, часть про улицы)."""
    d = str(desc or '')
    m = re.search(r'(улиц\w*|проезд\w*|переулк\w*|проспект\w*|шоссе|бульвар\w*|набережн\w*|'
                  r'микрорайон\w*|деревн\w*|поселк\w*)\s*[:]', d, re.I)
    if m:
        return d[:m.start()], d[m.start():]
    return d, ''


def _score(part, tokens_):
    """Сколько токенов адреса нашлось в тексте территории."""
    t = norm(part)
    hit = 0
    for w in tokens_:
        st = _stem(w)
        if len(st) < 4:
            continue
        if st in t:
            hit += 1
    return hit


def find(addr, region=''):
    """Адрес → (участок, пояснение).

    Участок — словарь из базы (code, name, region, court, num) либо None,
    если подсудность не определяется однозначно.
    """
    base = load()
    if not base['participants']:
        return None, 'локальная база судебных участков не загружена'
    # 1) запомненные сопоставления (живая ГАС отвечала раньше) — самые точные
    rem = _remembered(addr)
    if rem:
        return ({'name': rem.get('court') or '', 'num': '', 'code': '',
                 'court': rem.get('court') or '', 'region': '', 'desc': ''},
                'участок из сохранённых сопоставлений ГАС «Правосудие»')
    rc, rname = region_code(addr, base)
    if not rc and region:
        for code, rn in base['regions'].items():
            if region_key(rn) == region_key(region):
                rc = code
                break
    cands = base['by_region'].get(rc) or []
    if not cands:
        return None, (f'в базе нет участков для региона «{rname or region or "?"}»'
                      if rc or region else 'регион в адресе не определён')
    settle, streets = addr_tokens(addr)
    if not settle and not streets:
        return None, 'в адресе нет ни населённого пункта, ни улицы'
    scored = []
    for p in cands:
        s_part, st_part = _territory_parts(p.get('desc') or '')
        s = 10 * _score(s_part, settle) + _score(st_part, streets)
        s += 10 * _score(p.get('name') or '', settle)
        if s:
            scored.append((s, p))
    if not scored:
        return None, 'по описанию территории участков ничего не подошло'
    scored.sort(key=lambda x: -x[0])
    top = scored[0]
    if len(scored) > 1 and scored[1][0] == top[0]:
        return None, (f'по адресу подходят несколько участков с одинаковым совпадением '
                      f'({top[0]}) — суд определите вручную')
    p = top[1]
    note = (f'судебный участок № {p["num"]} {p.get("court") or ""}'.strip()
            + (f' (по базе участков, совпадений: {top[0]})' if top[0] else ''))
    return p, note


# ------------------------------------------------------------- шапка ------
def court_header(part):
    """Участок → строка шапки «Мировому судье судебного участка № 73 …»."""
    if not part:
        return ''
    name = (part.get('name') or '').strip()
    if not name:
        name = f'Судебный участок № {part.get("num", "?")}'
    n = re.sub(r'^Судебный участок', 'судебного участка', name).strip()
    if n == name:
        n = f'судебного участка {name}'
    return f'Мировому судье {n}'


def court_lines(part):
    """Строки шапки суда по участку (основная + адрес суда, если есть)."""
    head = court_header(part)
    if not head:
        return []
    out = [head]
    if part.get('site'):
        out.append(f'Официальный сайт суда: {part["site"]}')
    return out


if __name__ == '__main__':
    import sys
    q = ' '.join(sys.argv[1:])
    p, n = find(q)
    print('участок:', (p or {}).get('name'), '| код:', (p or {}).get('code'))
    print('пояснение:', n)
    if p:
        print('шапка:', court_header(p))
