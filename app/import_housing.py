#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Импорт данных в книгу ввода из:
  - «Портянки» (Excel расчёта: тип, ЛС, ФИО, адрес, период, долг, пени, госпошлина);
  - скана(-ов) справки формы 9 (OCR: ФИО, дата рождения, паспорт).

Использование:
    python3 import_housing.py [портянка.xlsx] [форм9.pdf ...]

Без аргументов ищет «Портянка*.xlsx» и «Ф-9*.pdf» в ../incoming/
(после установки — в ../filing/incoming/).
Один скан может содержать МНОГО справок (по одному на каждый лицевой счёт):
№ ЛС на бланке формы 9 стоит в ПРАВОМ ВЕРХНЕМ УГЛУ страницы — он распознаётся
автоматически по цифрам в первых строках OCR каждой страницы (затем по словам
«лицевой счёт/л/с» в тексте) и справки привязываются к своим ЛС
(лист «Ф-9 справки» в книге). Запасной способ: ЛС в имени файла
(«Ф-9 ЛС065005800.pdf») или вручную в «f9-index.xlsx».

Результат: «housing-input.xlsx»:
  листы Заявления и Должники заполнены, лист «Сверка импорта» — что проверить.

Правило (по указанию юриста): в документ включаются ТОЛЬКО основной наниматель
и/или собственники из справки формы 9; «Портянка» по ЛС показывает лишь одного
из ответчиков. Данные не двоятся: если по ЛС есть Ф-9, полный состав должников
берётся из неё, портянка даёт суммы/адрес/период/тип.
"""
import re
import shutil
import subprocess
import sys
import calendar
import difflib
import tempfile
from datetime import date, datetime
from pathlib import Path

import openpyxl

import platform_utils

BASE = Path(__file__).resolve().parent
IN_GUESS = BASE.parent / 'incoming'
if not IN_GUESS.is_dir():
    IN_GUESS = BASE.parent / 'filing' / 'incoming'
OUT_BOOK = BASE / 'housing-input.xlsx'

DATE_RE = re.compile(r'\b(\d{2}[.,]\d{2}[.,](?:19|20)\d{2})\b')
PASPORT_RE = re.compile(r'(?<![\d.])((?:4|8)\d{3})\s?,?\s?(\d{6})(?!\d)')
# лицевой счёт, написанный ВНУТРИ справки: «лицевой счёт № 065005800», «л/с 065005800»
# допустимы пробелы/дефисы/точки внутри номера: «065 005 800», «065-005-800»
LS_TEXT_RE = re.compile(
    r'(?:лицев[а-яё]+\s*сч[еёт]+[а-яё]*|л\s*/\s*с\s*сч[еёт]+[а-яё]*|л\s*/\s*с|лс|кл\.\s*сч[еёт]+[а-яё]*)'
    r'\s*(?:№|#|:)?\s*((?:[0-9][\s\-._]{0,2}){5,11}[0-9])',
    re.IGNORECASE)

# фамилия (возможно с maiden в скобках) + «Имя Отчество(ович и т.п.)» или «И.Г.»
PERSON_RE = re.compile(
    r'([А-ЯЁ][а-яё\-]+(?:\s*\([^()\n]{1,40}\))?)\s+'
    r'(?:([А-ЯЁ][а-яё\-]{2,})\s+([А-ЯЁ][а-яё\-]{2,}(?:ович|евич|овна|евна|ична|инична))'
    r'|([А-ЯЁ])\s?\.[\s,]?([А-ЯЁ])\.?)')
MINOR_RE = re.compile(r'в\s+инт\.?\s*(.*)$', re.IGNORECASE)
# для графы «собственник(и) являются / и/или наниматель»: фамилия в ЛЮБОМ регистре
# (OCR часто пишет со строчной: «у басильева Ж.П»), инициалы или полное ФИО
# граф: фамилия в любом регистре (OCR часто строчную) + инициалы ОБЯЗАТЕЛЬНО
# заглавными («Васильев А.А.») или полное «Имя Отчество»: это отсекает служебные
# строки бланка («России по», «право на», «состоящую из») и короткий шум OCR
GRAPH_RE = re.compile(
    r'([А-ЯЁа-яё][а-яё\-]{2,})\s+'
    r'(?:([А-ЯЁ])\.\s*([А-ЯЁ])\.?'
    r'|([А-ЯЁ][а-яё\-]+)\s+([А-ЯЁ][а-яё\-]+(?:ович|евич|овна|евна|ична|инична)))')
JUNK_GRAPH_RE = re.compile(
    r'выбывш|члены\s+семьи|инт\.\s|умерл|умерщ|умер\b', re.IGNORECASE)

# ------------------------------------------------------------- портянка ----
def fio_candidates(raw):
    """Поле ФИО портянки → [(фамилия, инициалы, полное ФИО)]; минусы несовершеннолетних."""
    raw = str(raw or '').strip()
    minors_ctx = None
    m = MINOR_RE.search(raw)
    if m:
        minors_ctx = 'в инт. ' + m.group(1).strip(' ,;.')
        raw = raw[:m.start()].strip(' ,;.,')
    persons = []
    for m in PERSON_RE.finditer(raw):
        surname = re.sub(r'\s+', ' ', m.group(1))
        if m.group(2):
            persons.append({'surname': surname,
                            'init': m.group(2)[0] + m.group(3)[0], 'full': f'{surname} {m.group(2)} {m.group(3)}'})
        else:
            persons.append({'surname': surname,
                            'init': m.group(4) + m.group(5), 'full': None})
    return persons, minors_ctx


def parse_period(raw):
    s = str(raw or '').replace('пот', 'по')
    m = re.search(r'(\d{2}[.,]\d{2}[.,]\d{4})\s*по\s*(\d{2}[.,]\d{2}[.,]\d{4})', s, re.IGNORECASE)
    if not m:
        return None, None, f'не удалось разобрать период: «{str(raw).strip()}»'
    d1, d2 = parse_date(m.group(1)), parse_date(m.group(2))
    err = None
    if d1 is None:
        d1, w1 = last_day_of_month(m.group(1))
        err = w1
    if d2 is None:
        d2, w2 = last_day_of_month(m.group(2))
        err = err or w2
    return d1, d2, err


def last_day_of_month(raw):
    """Несуществующая дата («31.11.2025») → (последний день месяца, предупреждение).

    В портянке встречаются опечатки в датах; из-за них заявление не собиралось
    (пустое «По состоянию на»). Принимаем ближайший существующий день и говорим об этом.
    """
    s = str(raw or '').replace(',', '.')
    m = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', s)
    if not m:
        return None, f'несуществующая дата: {raw} (в строке источника)'
    day, mon, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if not (1 <= mon <= 12):
        return None, f'несуществующая дата: {raw} (в строке источника)'
    last = calendar.monthrange(year, mon)[1]
    if day <= last:
        return None, f'несуществующая дата: {raw} (в строке источника)'
    fixed = date(year, mon, last)
    return fixed, (f'в строке источника дата {raw} не существует — принято '
                   f'{fixed:%d.%m.%Y}, проверьте период')


def parse_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, str):
        for f in ('%d.%m.%Y', '%d,%m,%Y'):
            try:
                return datetime.strptime(v.strip().strip(' .г'), f).date()
            except ValueError:
                pass
    return None


def load_portyanka(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    header_i = next((i for i, r in enumerate(rows)
                     if r and str(r[0]).strip() == '№ п/п'), None)
    if header_i is None:
        sys.exit('В портянке не найдена строка заголовков (№ п/п).')
    heads = [str(v).strip() if v is not None else '' for v in rows[header_i]]
    idx = {h: j for j, h in enumerate(heads) if h}

    items, warn = [], []
    for i, row in enumerate(rows[header_i + 1:], header_i + 2):
        if str(row[0] or '').strip().upper() == 'ИТОГО':
            break

        def val(name):
            if name not in idx:
                return None
            return row[idx[name]] if idx[name] < len(row) else None

        typ_raw = str(val('Тип') or '').strip()
        ls = str(val('Лицевой счет') or '').strip()
        if not ls:
            warn.append(f'строка {i}: нет лицевого счёта — строка пропущена')
            continue
        if 'приказ' in typ_raw.lower():
            typ = 'СП'
        elif 'иск' in typ_raw.lower():
            typ = 'ИСК'
        else:
            warn.append(f'строка {i} (ЛС {ls}): неизвестный тип «{typ_raw}»')
            typ = ''
        if not str(val('ФИО') or '').strip():
            warn.append(f'строка {i} (ЛС {ls}): пустое ФИО в портянке')

        def num(name):
            v = val(name)
            if v in (None, ''):
                return None
            try:
                return float(str(v).replace(' ', '').replace(',', '.'))
            except ValueError:
                warn.append(f'строка {i} (ЛС {ls}): не распознана сумма «{name}»: {v}')
                return None

        d1, d2, perr = parse_period(val('Период'))
        if perr:
            warn.append(f'строка {i} (ЛС {ls}): {perr}')
        dolg, peni, gp = num('Сумма долга'), num('Сумма долга пени'), num('Общая ГП')
        if dolg is not None and peni is not None:
            try:
                itogo = num('Сумма долга итого')
                if itogo is not None and abs(itogo - (dolg + peni)) > 0.011:
                    warn.append(f'строка {i} (ЛС {ls}): сумма итого {itogo:.2f} ≠ долг+пени {dolg + peni:.2f}')
            except (TypeError, ValueError):
                pass
        addr = re.sub(r'\s+', ' ', str(val('Адрес') or '').strip())
        persons, minors = fio_candidates(val('ФИО'))
        if not persons:
            warn.append(f'строка {i} (ЛС {ls}): из поля «{val("ФИО")}» не извлечено ни одного ФИО')
        items.append({'row': i, 'typ': typ, 'ls': ls, 'persons': persons,
                      'minors_text': minors, 'addr': addr, 'd1': d1, 'd2': d2,
                      'dolg': dolg, 'peni': peni, 'gp': gp})
    return items, warn


def gost_sum(cena, gp):
    """Заглушка: сверка пошлины появится после таблицы её расчёта."""
    return 0.0


def collect_groups(items, warn):
    """Объединяем строки портянки по (Тип, ЛС) в одно заявление."""
    groups = {}
    warns = []
    for it in items:
        key = (it['typ'], it['ls'])
        g = groups.setdefault(key, {'rows': [], 'it': it})
        g['rows'].append(it)
    out = []
    for (typ, ls), g in groups.items():
        rows = g['rows']
        base = rows[0]
        seen, merged = {}, []
        for r in rows:
            for p in r['persons']:
                key = (re.sub(r'\s+', ' ', p['surname']).lower(), p['init'])
                if key in seen:
                    seen[key]['rows_seen'].add(r['row'])
                    continue
                seen[key] = {**p, 'rows_seen': {r['row']}}
                merged.append(seen[key])
        # похожие фамилии при ОДИНАКОВЫХ инициалах — варианты одного лица:
        # «Романов (Сурмава) И.Г.» и «Романова (Сурмава) И.Г.» склеиваются,
        # а «Крылов Г.Н.» и «Крылов А.Г.» (разные инициалы) остаются раздельно
        # похожие фамилии при ОДИНАКОВЫХ инициалах — варианты одного лица:
        # «Романов (Сурмава) И.Г.» и «Романова (Сурмава) И.Г.» склеиваются,
        # а «Крылов Г.Н.» и «Крылов А.Г.» (разные инициалы) остаются раздельно
        core = lambda s: re.sub(r'\s*\([^)]*\)', '', s).strip().lower()
        for i, a in enumerate(merged):
            for b in merged[i + 1:]:
                sa, sb = core(a['surname']), core(b['surname'])
                similar = sa != sb and (sb.startswith(sa) or sa.startswith(sb))
                if similar and a['init'] == b['init']:
                    keep, drop = (a, b) if a['full'] else (b, a)
                    keep['rows_seen'] |= drop['rows_seen']
                    drop['drop'] = True
                elif similar and (a['init'] == '?' or b['init'] == '?'):
                    warns.append(f'ЛС {ls}: похожие фамилии с неполными ФИО («{a["surname"]}» / '
                                 f'«{b["surname"]}») — возможно, одно лицо, проверьте.')
        merged = [m for m in merged if not m.get('drop')]
        # лица из строк про несовершеннолетних («в инт. ...») — помечаются
        for m in merged:
            for r in rows:
                if r['row'] in m['rows_seen'] and r['minors_text']:
                    m['minor_flag'] = r['minors_text']
        base_dolg = {r['dolg'] for r in rows if r['dolg'] is not None}
        base_peni = {r['peni'] for r in rows if r['peni'] is not None}
        base_gp = {r['gp'] for r in rows if r['gp'] is not None}
        if len(base_dolg) > 1:
            warn.append(f'ЛС {ls}: разные суммы долга в строках портянки '
                        f'({", ".join(str(x) for x in sorted(base_dolg))}) — проверьте!')
        d2_all = {r['d2'] for r in rows if r['d2'] is not None}
        minors_all = [r['minors_text'] for r in rows if r['minors_text']]
        out.append({'typ': typ, 'ls': ls, 'persons': merged,
                    'minors_text': '; '.join(minors_all),
                    'addr': base['addr'], 'd1': base['d1'],
                    'd2': base['d2'], 'dolg': sorted(base_dolg)[0] if base_dolg else None,
                    'peni': sorted(base_peni)[0] if base_peni else None,
                    'gp': sorted(base_gp)[0] if base_gp else None,
                    'nrows': len(rows)})
        if not base_dolg and not base_peni:
            warn.append(f'ЛС {ls}: суммы долга и пени пустые/нулевые — заявление, '
                        'возможно, не нужно. Проверьте.')
        if len(rows) > 1:
            warn.append(f'ЛС {ls}: в портянке {len(rows)} строк с одним ЛС — '
                        'объединены в одно заявление (двоение исключено).')
    warn.extend(warns)
    return out


# ------------------------------------------------------------------ Ф-9 ----
def norm_ls(v):
    """Ключ сравнения ЛС: только цифры без ведущих нулей (Excel их теряет)."""
    digits = re.sub(r'\D', '', str(v))
    return digits.lstrip('0') or digits


DATE_SHORT_RE = re.compile(r'\d{2}[.,/\-]\d{2}[.,/\-]\d{4}')


# Строки-мусор, которые OCR выдаёт вместо номера ЛС в углу страницы.
# Слова нейтральные: названий реальных регионов здесь быть не должно —
# файл публикуется, и любая география в нём должна быть вымышленной.
LS_JUNK = ('российская', 'область', 'справка', 'акционерное', 'о регистрации')


def ls_digit_variants(line):
    """Строка угла страницы → набор строк-цифр: прямые цифры + расшифровка букв,
    в которые OCR превратил цифры («ДС Об Бо 9590» → «065009590», «007-690» → «007690»)."""
    line = DATE_SHORT_RE.sub(' ', line)
    low = line.strip().lower()
    if any(k in low for k in LS_JUNK):
        return set()
    variants = ['']
    for ch in line:
        if ch.isdigit():
            add = [ch]
        elif ch in 'ОоАаСсд':
            add = ['0']
        elif ch in 'бД':
            add = ['6']
        elif ch in 'БЕ':
            add = ['6', '5']
        elif ch in 'ЗзЭэ':
            add = ['3']
        elif ch in 'лЛ':
            add = ['4']
        elif ch in 'Т':
            add = ['1', '7']
        elif ch == 'е':
            add = ['0', '6']
        elif ch in ' .-–—<>()^\\/:;|,':
            add = ['']  # пунктуация/пробел могут быть разрывом или мусором
        else:
            break  # прочая буква — дальше только прямой набор цифр
        variants = [v + a for v in variants for a in add]
        if len(variants) > 256:
            break
    out = set()
    for v in variants:
        out.update(x for x in re.findall(r'\d{5,12}', v))
    out.update(re.findall(r'\d{5,12}', re.sub(r'\s+', '', line)))
    return out


def page_ls_candidates(text):
    """Кандидаты ЛС страницы, СПЕЦИАЛЬНО по справке формы 9:
    № ЛС стоит в ПРАВОМ ВЕРХНЕМ УГЛУ — просто цифры в первых строках OCR
    (без слов «лицевой счёт»). Поэтому сперва берём числа из первой дюжины строк,
    отсекая даты и мелкие номера страниц, затем — числа рядом со словами
    «лицевой счёт / л/с / кл. счёт» по всему тексту."""
    out = []
    head = [l.strip() for l in text.splitlines() if l.strip()][:12]
    for line in head:
        for m in re.finditer(r'(?<![\d.])(\d[\d\s.\-\u2013\u2014]{4,16}\d)(?![\d.])',
                             DATE_SHORT_RE.sub(' ', line)):
            cand = re.sub(r'\D', '', m.group(1))
            if 6 <= len(cand) <= 12 and cand not in out:
                out.append(cand)
    for m in LS_TEXT_RE.finditer(text):
        cand = re.sub(r'[\s\-._]', '', m.group(1))
        if 6 <= len(cand) <= 12 and cand not in out:
            out.append(cand)
    return out


def ocr_pdf(path: Path, tmpdir: Path):
    """OCR файла → [(номер страницы, текст)].

    Каждая страница распознаётся отдельно — так один скан со многими
    справками можно привязать по индексу «страница → ЛС».
    Страница без шапки «Выдана гр.» считается продолжением предыдущей
    (двухстраничная справка), и её текст приклеивается к предпоследней.

    """
    subprocess.run(['pdftoppm', '-r', '300', '-gray', '-png',
                    str(path), str(tmpdir / 'p')], check=True)
    pages = []
    imgs = sorted(tmpdir.glob('p-*.png'), key=lambda p: int(re.search(r'p-(\d+)', p.stem).group(1)))
    for img in imgs:
        n = int(re.search(r'p-(\d+)', img.stem).group(1))
        subprocess.run(['tesseract', str(img), str(tmpdir / f't{n}'), '-l', 'rus',
                        '--psm', '4'], check=True, capture_output=True)
        pages.append((n, (tmpdir / f't{n}.txt').read_text(encoding='utf-8')))
        print(f'· страница {n} распознана', flush=True)
    merged = []
    for n, t in pages:
        if merged and not re.search(r'Выдана\s+гр', t):
            m = merged[-1]
            merged[-1] = (m[0], m[1] + '\n' + t)
        else:
            merged.append((n, t))
    return merged


# ФИО в шапке справки: фамилия + имя + отчество (или «И.О.»)
F9_FIO_RE = re.compile(
    r'([А-ЯЁ][а-яё\-]+)\s+'
    r'(?:([А-ЯЁ][а-яё\-]+)\s+([А-ЯЁ][а-яё\-]+(?:ович|евич|овна|евна|ична|инична|ини[а-яё]*))'
    r'|([А-ЯЁ])\.\s?([А-ЯЁ])\.?)')
# дата рождения в шапке: между ФИО и ней может быть место рождения (город, посёлок)
F9_DR_RE = re.compile(
    r'Выдана\s+гр[а-яё.]*\s*([^\n]{0,140}?)'
    r'(\d{2}[.,]\d{2}[.,]\d{4})[^\n]{0,20}?рожден', re.IGNORECASE)
# адрес в справке: «в доме № 38, кв.№ 16, корпус № , по улице Московское ш., г. Пример»
F9_ADDR_RE = re.compile(
    r'в\s*д\s*оме\s*№?\s*\.?\s*(\d{1,4})'          # дом
    r'[^|\n]{0,45}?к\.?\s*в?\.?\s*№?\s*\.?\s*(\d{1,4})'  # квартира
    r'[^|\n]{0,45}?(?:по|то|во)\s+улице\s*([^|\n]{2,45})',  # улица
    re.IGNORECASE)
# запасной вариант адреса — без улицы (если OCR потерял слово «улице»)
F9_ADDR_NOUL_RE = re.compile(
    r'в\s*д\s*оме\s*№?\s*\.?\s*(\d{1,4})'
    r'[^|\n]{0,45}?к\.?\s*в?\.?\s*№?\s*\.?\s*(\d{1,4})',
    re.IGNORECASE)


def parse_f9_addr(text):
    """Строка адреса из справки Ф-9 → (дом, квартира, улица).

    Дом и квартира — без ведущих нулей («038» → «38»), улица — без «, г. Пример».
    Пусто, если адрес в тексте не найден.
    """
    m = F9_ADDR_RE.search(text) or F9_ADDR_NOUL_RE.search(text)
    if not m:
        return '', '', ''
    dom = (m.group(1) or '').lstrip('0') or m.group(1)
    kv = (m.group(2) or '').lstrip('0') or m.group(2)
    ul = ''
    if m.lastindex and m.lastindex >= 3 and m.group(3):
        ul = re.sub(r'\s*,?\s*г\.\s*Пример.*$', '', m.group(3), flags=re.IGNORECASE)
        ul = re.sub(r'\s+', ' ', ul).strip(' ,.')
    return dom, kv, ul


def parse_f9_birth(text):
    """Шапка справки Ф-9 → (ФИО полностью, дата рождения 'ДД.ММ.ГГГГ').

    Между ФИО и датой может стоять место рождения (город, посёлок),
    поэтому ФИО ищем в тексте ДО даты, а не по жёсткому шаблону.
    """
    m = F9_DR_RE.search(text)
    if not m:
        return '', ''
    dr = m.group(2).replace(',', '.')
    head = m.group(1)
    fm = F9_FIO_RE.search(head)
    if not fm:
        return '', dr
    if fm.group(2):   # фамилия + имя + отчество
        fio = f'{fm.group(1)} {fm.group(2)} {fm.group(3)}'
    else:             # фамилия + инициалы
        fio = f'{fm.group(1)} {fm.group(4)}.{fm.group(5)}.'
    return re.sub(r'\s+', ' ', fio).strip(), dr


def parse_f9(text):
    """Текст справки ф.9 → список лиц: {fio, init, dr, place, pasport, kem}."""
    persons = {}

    def add(surname, init, full, dr='', place='', pp='', kem=''):
        key = (re.sub(r'\s*\([^)]*\)', '', surname).lower(), init)
        e = persons.setdefault(key, {'surname': surname, 'init': init,
                                     'fio': full or (f'{surname} {init[0]}.{init[1]}.' if init and len(init) >= 2 else surname),
                                     'dr': '', 'place': '', 'pasport': '', 'kem': ''})
        for k, v in (('dr', dr), ('place', place), ('pasport', pp), ('kem', kem)):
            if v and not e.get(k):
                e[k] = v

    # основной наниматель: «Выдана гр. ФИО [место рождения] [дата] года рождения»
    fio, dr = parse_f9_birth(text)
    if fio:
        w = fio.split()
        add(w[0], w[1][0] + w[2][0] if len(w) >= 3 else '?', fio, dr=dr)
    # таблица зарегистрированных — до раздела о выбывших;
    # шапку «Выдана гр. …» не разбираем как таблицу (там ФИО в дательном падеже,
    # не зарегистрированное лицо — иначе мусорная строка попадает в «Должники»)
    seg = text
    head = re.search(r'Выдана\s+гр[а-яё.]*[^\n]*', text, re.IGNORECASE)
    if head:
        seg = text[head.end():]
    cut = re.search(r'выбывш[а-яё]*|Члены семьи', seg, re.IGNORECASE)
    if cut:
        seg = seg[:cut.start()]
    spans = [(m.start(), m.end(), m) for m in PERSON_RE.finditer(seg)]
    for j, (s0, e0, m) in enumerate(spans):
        nxt = spans[j + 1][0] if j + 1 < len(spans) else len(seg)
        after = seg[e0:nxt]
        surname = re.sub(r'\s+', ' ', m.group(1))
        pp, pasport_pos = '', len(after)
        pm = PASPORT_RE.search(after)
        if pm:
            pp = f'{pm.group(1)} {pm.group(2)}'
            pasport_pos = pm.start()
        kem = ''
        if pm:
            rest = re.sub(r'\s+', ' ', after[pm.end():]).strip(' ,;')
            # «кем выдан» — только до конца строки таблицы: дальше начинается
            # следующая запись (следующее ФИО/дата), и без обрезки она
            # приклеивается к паспорту («4012 345678, выдан Иванов Пётр…»)
            rest = re.split(r'\s(?=[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+|'
                            r'\d{2}[.,]\d{2}[.,]\d{4}\b)', rest)[0]
            rest = re.sub(r'^(выдан|выдано|кем выдан)\s+', '', rest, flags=re.IGNORECASE)
            kem = re.sub(r'^[.,;\s]+', '', rest)[:120]
        # место рождения — текст между ФИО и первой датой/паспортом
        first_date = DATE_RE.search(after)
        place_to = min(x for x in (pasport_pos,
                                   first_date.start() if first_date else len(after)) if x is not None)
        place = re.sub(r'\s+', ' ', after[:place_to]).strip(' ,;.')
        if re.search(r'регистрац|Год|Фамилия|постоянно|временно', place, re.IGNORECASE):
            place = ''
        # дата рождения — первая дата, НО если паспорт идёт раньше неё, это дата выдачи
        dates = [d[0].replace(',', '.') for d in DATE_RE.finditer(after)]
        dr = dates[0] if dates and (not pm or pm.start() > DATE_RE.search(after).start()) else ''
        if m.group(2):
            add(surname, m.group(2)[0] + m.group(3)[0], f'{surname} {m.group(2)} {m.group(3)}',
                dr, place, pp, kem)
        else:
            add(surname, m.group(4) + m.group(5), None, dr, place, pp, kem)
    return list(persons.values())


GRAPH_MARKER_RE = re.compile(
    r'(?:собственник[а-яё]*\s*(?:и\s*/?\s*или\s*нанимател[а-яё]*|явл[а-яё]+)'
    r'|наниматель[а-яё]*\s*явл[а-яё]+)', re.IGNORECASE)


def parse_f9_graph(text, page_keys=()):
    """Нижняя графа справки «собственником является / собственник и/или наниматель»
    → [«Фамилия И.О.»].

    По указанию юриста: для судебных приказов значимы ТОЛЬКО лица из этой графы
    (кол-во приказов = кол-во лиц), зарегистрированные игнорируются;
    для иска графовые лица + совершеннолетние зарегистрированные — одним документом.

    page_keys — «ключи лиц» зарегистрированных на ЭТОЙ странице (fio_key-подобные);
    имя из графы оставляем, только если оно подтверждается таблицей регистраций той
    же страницы (защита от имён подписантов и шума OCR: «Паспортист Е.М.»).
    """
    m = GRAPH_MARKER_RE.search(text)
    if not m:
        return []
    seg = text[m.start(): m.start() + 700]
    persons = []
    for line in seg.splitlines():
        ln = line.strip()
        if not ln or JUNK_GRAPH_RE.search(ln):
            continue
        for pm in GRAPH_RE.finditer(ln):
            if re.search(r'умер', line, re.IGNORECASE):
                continue
            fam_core = re.sub(r'\s*\([^)]*\)', '', pm.group(1)).strip().lower()
            fam = fam_cap(pm.group(1))
            if pm.group(4) and pm.group(5):
                persons.append((fam, f'{pm.group(4)[0]}{pm.group(5)[0]}'))
            elif pm.group(2) and pm.group(3):
                persons.append((fam, pm.group(2) + pm.group(3)))
    # оставить только подтверждённое таблицей регистраций страницы
    if page_keys:
        kept = []
        for fam, init in persons:
            fam_low = fam.lower()
            if (fam_low, init) in page_keys:
                kept.append(f'{fam} {init[0]}.{init[1]}.')
                continue
            for pfam, pinit in page_keys:
                if pinit == init and difflib.SequenceMatcher(None, fam_low, pfam).ratio() >= 0.72:
                    kept.append(f'{fam} {init[0]}.{init[1]}.')
                    break
        persons = kept
    else:
        persons = [f'{fam} {i[0]}.{i[1]}.' for fam, i in persons]
    seen, out = set(), []
    for x in persons:
        k = x.lower()
        if k not in seen:
            seen.add(k)
            out.append(x)
    return out[:8]


def fam_cap(x):
    return re.sub(r'\s+', ' ', x).strip().capitalize()


def load_f9_index():
    """«f9-index.xlsx» (Файл / Страница / ЛС) → ({имя файла: {страница: ЛС}}, [замечания]).

    Ничего не отбрасываем молча: ЛС, который Excel превратил в число («65005800.0»),
    страница «стр. 2», имя файла без расширения или с лишними пробелами — всё читается;
    нераспознанное попадает в замечания с номером строки.
    """
    idx = BASE / 'f9-index.xlsx'
    if not idx.exists():
        return {}, []
    wb = openpyxl.load_workbook(idx)
    ws = wb.active
    heads = [str(c.value).strip() if c.value else '' for c in ws[1]]
    if 'Файл' not in heads or 'Страница' not in heads or 'ЛС' not in heads:
        return None, []   # колонок нет — сообщит вызывающий
    i_file, i_page, i_ls = heads.index('Файл'), heads.index('Страница'), heads.index('ЛС')
    out, notes, n_rows, n_bound = {}, [], 0, 0
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r is None or all(v is None or str(v).strip() == '' for v in r):
            continue
        n_rows += 1
        fname = str(r[i_file] or '').strip()
        pm = re.search(r'\d+', str(r[i_page] or ''))
        page = int(pm.group()) if pm else None
        digits = re.sub(r'\D', '', str(r[i_ls] or ''))
        ls = digits.lstrip('0') or digits
        if not fname:
            notes.append(f'строка {n_rows}: не указано имя файла скана')
        elif page is None:
            notes.append(f'строка: «{fname}»: не распознан номер страницы («{r[i_page]}»)')
        elif not ls:
            notes.append(f'строка: «{fname}», стр. {page}: не распознан ЛС («{r[i_ls]}»)')
        else:
            d = out.setdefault(fname.lower(), {})
            d[page] = ls
            out.setdefault(fname.lower().rsplit('.', 1)[0], d)   # и то же — без расширения
            n_bound += 1
    notes.append(f'Индекс принят: строк {n_rows}, привязок {n_bound}.')
    return out, notes


def person_keys(persons):
    """Список лиц страницы/портянки → {(фамилия_без_скобок_в_нижнем_регистре, инициалы)}."""
    keys = set()
    for pr in persons:
        fam = re.sub(r'\s*\([^)]*\)', '', str(pr.get('surname') or '')).strip().lower()
        init = str(pr.get('init') or '')
        if len(fam) >= 2 and len(init) >= 2:
            keys.add((fam, init))
    return keys


def best_ls_match(blob, known):
    """Строка-блоб цифр угла страницы → (ЛС, точное?, оценка) среди известных ЛС.

    Сначала точное совпадение цифр; затем нечёткое (difflib) — OCR портит и цифры
    (буквами), и теряет/добавляет разряды («Об Бо 9590» ← 0650099?90). Выбираем
    лучшего и принимаем его, только если он ОДНОЗНАЧНО лучше второго (зазор ≥ 0.06)
    и оценка ≥ 0.80 — иначе None: ошибочная привязка хуже непривязанной страницы.

    Ключи known — нормализованные (Excel теряет ведущие нули: «000008256» → «8256»),
    поэтому искомый блоб сравниваем и с СЫРЫМ ЛС, и с нормализованным: без этого
    точное совпадение не срабатывало никогда, а счёта с ведущими нулями не набирали
    порог нечёткого совпадения.
    """
    if not blob:
        return None, False, 0.0
    cands = {}
    for k, v in known.items():
        cands.setdefault(k, v)                     # нормализованный
        raw = str(v)
        cands.setdefault(raw, v)                  # сырой (с ведущими нулями)
        cands.setdefault(raw.lstrip('0') or raw, v)
    if blob in cands:
        return cands[blob], True, 1.0
    best, best_r, second = None, 0.0, 0.0
    for k, v in cands.items():
        r = difflib.SequenceMatcher(None, blob, k).ratio()
        if r > best_r:
            best, best_r, second = v, r, best_r
        elif r > second:
            second = r
    if best_r >= 0.80 and best_r - second >= 0.06:
        return best, False, best_r
    return None, False, best_r



def fio_norm(s):
    """ФИО → (фамилия, инициалы) в сравнимом виде: нижний регистр, ё→е, без скобок."""
    s = str(s or '').lower().replace('ё', 'е')
    s = re.sub(r'\s*\([^)]*\)', ' ', s)
    parts = [p for p in re.split(r'[\s,]+', s) if p]
    parts = [p for p in parts
             if p not in ('в', 'инт', 'н', 'за', 'и', 'или', 'все', 'их', 'проживающий',
                          'совместно', 'с', 'гражданином', 'гражданиномкой', 'проживающими')]
    parts = [re.sub(r'^(н|м|ж)[\-.]', '', p) for p in parts]   # «н-го», «м-ей»
    if not parts:
        return '', ''
    fam = parts[0]
    inits = ''
    for tok in parts[1:4]:
        t = tok.replace('.', '')
        if not t:
            continue
        inits += t if (len(t) <= 3 and t.isalpha() and t.islower()) else t[0]
    return fam, inits


def name_score(fio_a, fio_b):
    """0..1: насколько похожи два ФИО — 0.6 за фамилию, 0.4 за инициалы.

    Фамилии сравниваем точно, а при несовпадении — нечётко (difflib, порог 0.75):
    в портянке и в справке фамилия может быть с опечаткой («Васильчиков»/«Васильчиков»).
    """
    fa, ia = fio_norm(fio_a)
    fb, ib = fio_norm(fio_b)
    if not fa or not fb:
        return 0.0
    s = 0.0
    if fa == fb:
        s += 0.6
    else:
        r = difflib.SequenceMatcher(None, fa, fb).ratio()
        if r >= 0.75:
            s += 0.6 * r
    if ia and ib:
        if ia == ib:
            s += 0.4
        elif len(ia) == len(ib):
            same = sum(1 for x, y in zip(ia, ib) if x == y)
            s += 0.4 * same / len(ia)
    return min(s, 1.0)


def port_addr_parts(raw):
    """Адрес портянки («Пример 2,Тестовая,д.27кв.63») → (дом, кв) без ведущих нулей."""
    s = str(raw or '')
    dom, kv = '', ''
    m = re.search(r'д[.\s№]*(\d{1,4})', s, re.IGNORECASE)
    if m:
        dom = m.group(1)
    m = re.search(r'кв[.\s№]*(\d{1,4})', s, re.IGNORECASE)
    if m:
        kv = m.group(1)
    return (dom.lstrip('0') or dom), (kv.lstrip('0') or kv)


def build_port_index(groups):
    """Заявления портянки → {ЛС: {'dom','kv','fios': [ФИО], 'addr': адрес}}.

    Нужно для привязки справок Ф-9 по фамилии и адресу (по указанию юриста):
    рукописный ЛС в углу справки OCR читает ненадёжно, а фамилия с адресом — надёжно.
    """
    idx = {}
    for g in groups:
        dom, kv = port_addr_parts(g.get('addr'))
        fios = []
        for p in g.get('persons') or []:
            nm = p.get('full') or f"{p.get('surname', '')} {p.get('init', '')}".strip()
            if nm:
                fios.append(nm)
        idx[str(g.get('ls'))] = {'dom': dom, 'kv': kv, 'fios': fios,
                                 'addr': g.get('addr') or ''}
    return idx


def match_page_ls(t, prs, known, known_persons, port_index=None):
    """Страница Ф-9 → (ЛС, как привязано, причина непривязки).

    Порядок (по указанию юриста — «привязывать по фамилиям и адресам»):
      1. рукописный ЛС в углу страницы — только ТОЧНОЕ совпадение: OCR рукописи
         даёт сплошные ошибки («46060464445» вместо «065006700»), нечёткое
         совпадение тут даёт ложные привязки;
      2. фамилия И адрес совпали одновременно — основной путь;
      3. ЛС, написанный в тексте справки рядом со «л/с» (точное совпадение);
      4. запасная привязка по ФИО зарегистрированных.
    """
    port_index = port_index or {}
    # 1) угол страницы — только точное совпадение
    head = [l.strip() for l in t.splitlines() if l.strip()][:5]
    for line in head:
        line = DATE_SHORT_RE.sub(' ', line)
        for blob in sorted(ls_digit_variants(line)):
            m_ls, is_exact, _score = best_ls_match(blob, known)
            if m_ls and is_exact:
                return m_ls, 'распознан в углу страницы', ''

    # 2) фамилия И адрес (основной путь)
    if port_index:
        f9_fio, _dr = parse_f9_birth(t)
        if not f9_fio and prs:
            f9_fio = prs[0].get('fio') or ''
        dom, kv, _ul = parse_f9_addr(t)
        addr_ok = {}
        for ls, info in port_index.items():
            if dom and kv and info['dom'] == dom and info['kv'] == kv:
                addr_ok[ls] = info
        scored = []
        for ls, info in addr_ok.items():
            best = max((name_score(f9_fio, f) for f in info['fios']), default=0.0)
            scored.append((best, ls, info))
        # совпали и фамилия, и адрес → привязываем
        good = [x for x in scored if x[0] >= 0.55]
        if len(good) == 1:
            s, ls, info = good[0]
            return ls, f'по фамилии и адресу (совпадение {s:.0%})', ''
        if len(good) > 1:
            good.sort(reverse=True)
            if good[0][0] > good[1][0]:
                s, ls, info = good[0]
                return ls, f'по фамилии и адресу (совпадение {s:.0%})', ''
            return None, '', ('похожие адреса в портянке: '
                              + ', '.join(x[1] for x in good) + ' — уточните вручную')
        # не подошло ни одно — разбираем причину
        if not f9_fio:
            reason = 'в справке не распознано ФИО владельца'
        elif not dom or not kv:
            reason = 'в справке не распознан адрес (дом/квартира)'
        elif not addr_ok:
            reason = (f'в портянке нет адреса «д. {dom}, кв. {kv}»')
        else:
            ls = max(scored, key=lambda x: x[0])
            names = ', '.join(ls[2]['fios'][:3]) or 'ФИО не заполнено'
            reason = (f'адрес совпал с ЛС {ls[1]}, но в портянке «{names}», '
                      f'а в справке «{f9_fio}» — проверьте вручную')
        return None, '', reason

    # 3) ЛС, написанный в тексте справки рядом со «л/с»
    for cand in page_ls_candidates(t):
        m_ls, is_exact, _score = best_ls_match(cand, known)
        if m_ls and is_exact:
            return m_ls, 'распознан в тексте справки (рядом с «л/с»)', ''

    # 4) запасная привязка по ФИО зарегистрированных (портянка ↔ таблица справки)
    if known_persons:
        pkeys = person_keys(prs)
        if pkeys:
            votes = {}
            for kls, pk_set in known_persons.items():
                hit = len(pkeys & pk_set)
                if hit:
                    votes[kls] = hit
            if votes:
                top = max(votes.values())
                winners = [kls for kls, v in votes.items() if v == top]
                if len(winners) == 1:
                    return winners[0], 'по ФИО зарегистрированных (совпадение с портянкой)', ''
    return None, '', 'не удалось сопоставить ни по фамилии с адресом, ни по углу страницы'



def load_f9_files(paths, report, warn, known_ls=(), known_persons=None, port_index=None):
    tmp = Path(tempfile.mkdtemp(prefix='f9ocr-'))
    out_dir = BASE / 'output' / 'f9-text'
    out_dir.mkdir(parents=True, exist_ok=True)
    by_ls, unattached, texts_by_ls = {}, [], {}
    known_persons = known_persons or {}
    known = {norm_ls(x): str(x) for x in known_ls if norm_ls(x)}
    index, idx_notes = load_f9_index()
    if index is None:
        warn.append('В файле «f9-index.xlsx» нет колонок «Файл», «Страница», «ЛС» — '
                    'поскипал привязку по нему.')
    elif idx_notes:
        report.append('Ф-9 индекс: ' + idx_notes[-1])
        for w in idx_notes[:-1]:
            warn.append('Ф-9 индекс: ' + w)
    total_paths = len(paths)
    for pi, path in enumerate(paths, 1):
        try:
            pages = ocr_pdf(path, tmp)
        except Exception as ex:  # noqa: BLE001
            warn.append(f'Ф-9 «{path.name}»: не удалось распознать ({ex})')
            continue
        print(f'[{pi}/{total_paths}] {path.name}: распознано страниц: {len(pages)}', flush=True)
        (out_dir / platform_utils.ascii_filename(path.stem + '.txt')).write_text(
            '\n\n=== страница ' + ' ===\n\n=== '.join(
                str(n) for n, t in pages) + ' ===\n' + '\n\n=== конец страницы ===\n\n'.join(
                t for n, t in pages), encoding='utf-8')
        lm = re.search(r'ЛС\s*[:№]?\s*([0-9]{6,12})', path.stem, re.IGNORECASE)
        ls_name = lm.group(1) if lm else ''
        if ls_name:  # ЛС в имени файла — все страницы к нему
            n_faces = 0
            for n, t in pages:
                prs = parse_f9(t)
                n_faces += len(prs)
                by_ls.setdefault(ls_name, []).extend(prs)
                texts_by_ls.setdefault(ls_name, []).append(
                    (path.name, n, t, parse_f9_graph(t, person_keys(prs)), 'по имени файла'))
            report.append(f'Ф-9 «{path.name}»: ЛС {ls_name} (из имени) → страниц: {len(pages)}, лиц: {n_faces}')
            continue
        mp = (index or {}).get(path.name.lower())
        bound, from_text = [], 0
        graphs_log = []
        for n, t in pages:
            prs = parse_f9(t)
            ls = (mp or {}).get(n, '')
            reason = ''
            if ls:
                ls = known.get(norm_ls(ls), ls)  # индекс мог съесть ведущий ноль
                how = ' (по индексу)'
            if not ls:
                ls, how, reason = match_page_ls(t, prs, known, known_persons, port_index)
                if how:
                    how = f' ({how})'
                else:
                    how = ''
            graph = parse_f9_graph(t, person_keys(prs))
            graphs_log.append((path.name, n, bool(graph)))
            dead_note = ''
            if GRAPH_MARKER_RE.search(t):
                gseg = t[GRAPH_MARKER_RE.search(t).start():]
                if re.search(r'умер', gseg[:400], re.IGNORECASE) and not graph:
                    dead_note = '; в графе «собственник» отмечен умерший'
            if ls:
                by_ls.setdefault(ls, []).extend(prs)
                texts_by_ls.setdefault(ls, []).append((path.name, n, t, graph, how.strip(' ()')))
                bound.append(f'{n} → ЛС {ls}{how}' + dead_note)
            elif prs or graph:
                unattached.append({'file': path.name, 'page': n, 'persons': prs,
                                   'reason': reason})
                if dead_note:
                    warn.append(f'Ф-9 «{path.name}», стр. {n}: {dead_note.strip("; ")}')
        if bound:
            report.append(f'Ф-9 «{path.name}»: страниц: {len(pages)}, привязано: {len(bound)}: '
                          + '; '.join(bound))
            # страницы без графы «собственник/наниматель» — юристу вписать вручную
            pages_nog = []
            for b in bound:
                n = int(b.split(' ')[0])
                gl = [gr for _, pg_n, gr in graphs_log if pg_n == n]
                if not gl or not gl[0]:
                    pages_nog.append(str(n))
            if pages_nog:
                warn.append(f'Ф-9 «{path.name}»: на страницах {", ".join(pages_nog)} не распознана графа '
                            '«собственник/наниматель» — СП/иск соберутся по зарегистрированным; '
                            'впишите лица вручную на лист «Ф-9 справки» (через «;») и соберите заново.')
        if not mp and not bound:
            report.append(f'Ф-9 «{path.name}»: ЛС не распознан (ни по фамилии с адресом, ни в углу, '
                          'ни по ФИО) → страницы ждут привязки. Что распозналось, смотрите в '
                          f'«output/f9-text/{platform_utils.ascii_filename(path.stem + ".txt")}»; дальше — ЛС в имени файла '
                          'или вручную в «f9-index.xlsx» (Файл/Страница/ЛС) и перезапуск импорта.')
        bound_pages = {int(b.split(" ")[0]) for b in bound}
        miss = [n for n, _ in pages if n not in bound_pages]
        if miss and (mp or bound):
            warn.append(f'Ф-9 «{path.name}»: страницы {miss} без ЛС — '
                        'лица ждут на листе «Ф-9 не привязаны».')
    return by_ls, unattached, texts_by_ls


# --------------------------------------------------------------- сборка ----
GUIDE_SHEET = 'Справочник ЛС'
# суд по типу документа — подставляется, только если юрист не задал свой
DEFAULT_SUD = {
    'СП': 'Мировому судье судебного участка № 75 / Примерного района Примерной области',
    'ИСК': 'В Примерный городской суд / Примерного района Примерной области',
}
# поля, которые переносим из прежней книги при пересборке
KEEP_COLS = ('Категория', 'Суд (шапка)', 'Тип жилья (текст)', 'Цена иска, руб.', 'Отмена СП (да/нет)')


def default_sud(typ):
    return DEFAULT_SUD.get(str(typ or '').strip().upper(), '')


def read_guide(wb):
    """Лист «Справочник ЛС» → {ЛС: {'Категория':…, 'Суд (шапка)':…, 'Тип жилья (текст)':…}}.

    Это долговременная память юриста: один раз вписал «наниматель/собственник» —
    и значение подставляется при каждом импорте и не теряется.
    """
    out = {}
    if GUIDE_SHEET not in wb.sheetnames:
        return out
    ws = wb[GUIDE_SHEET]
    heads = [str(c.value).strip() if c.value else '' for c in ws[1]]
    if 'Лицевой счёт' not in heads:
        return out
    i_ls = heads.index('Лицевой счёт')
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or i_ls >= len(row):
            continue
        ls = str(row[i_ls] or '').strip()
        if not ls:
            continue
        rec = {}
        for col in ('Категория', 'Суд (шапка)', 'Тип жилья (текст)'):
            if col in heads:
                v = row[heads.index(col)] if heads.index(col) < len(row) else None
                if v is not None and str(v).strip():
                    rec[col] = str(v).strip()
        out[ls] = rec
    return out


def read_prev_book(path):
    """Прежняя книга → {(тип, ЛС): {колонка: значение}} — ручные правки юриста.

    Импорт пересобирает книгу с нуля, и без этого ручной труд (категория, суд,
    примечания) пропадал бы (инцидент 20.10.2026).
    """
    out = {}
    if not path or not Path(path).exists():
        return out
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception:  # noqa: BLE001
        return out
    if 'Заявления' not in wb.sheetnames:
        return out
    ws = wb['Заявления']
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return out
    heads = [str(v).strip() if v is not None else '' for v in rows[0]]
    try:
        i_typ, i_ls = heads.index('Тип документа'), heads.index('Лицевой счёт')
    except ValueError:
        return out
    for row in rows[1:]:
        if not row or all(v is None or str(v).strip() == '' for v in row):
            continue
        ls = str(row[i_ls] or '').strip()
        typ = str(row[i_typ] or '').strip().upper()
        if not ls:
            continue
        rec = {}
        for col in KEEP_COLS:
            if col in heads:
                j = heads.index(col)
                v = row[j] if j < len(row) else None
                if v is not None and str(v).strip():
                    rec[col] = v
        if rec:
            out[(typ, ls)] = rec
    return out


def scan_saved_categories(base_dir=None):
    """Реестры прошлых сборов и бэкапы книги → {ЛС: категория}.

    Второй контур памяти: если категория не нашлась ни в «Справочнике ЛС»,
    ни в прежней книге (инцидент 21.10.2026 — 25 заявлений не собрались), её
    всё ещё можно достать из готовых Реестров предыдущих запусков generate.py
    или из бэкапов книги. Возвращаются только значения «наниматель/собственник».
    """
    base_dir = Path(base_dir) if base_dir else BASE
    out = {}
    sources = sorted(base_dir.glob('output/*/registry.xlsx'), key=lambda p: p.stat().st_mtime)
    sources += sorted(base_dir.glob('housing-input-backup*.xlsx'),
                      key=lambda p: p.stat().st_mtime)
    for p in sources:
        try:
            wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
        except Exception:  # noqa: BLE001
            continue
        ws = wb['Заявления'] if 'Заявления' in wb.sheetnames else wb.active
        try:
            rows = list(ws.iter_rows(values_only=True))
        except Exception:  # noqa: BLE001
            continue
        if not rows:
            continue
        heads = [str(v).strip() if v is not None else '' for v in rows[0]]
        if 'Категория' not in heads or 'Лицевой счёт' not in heads:
            continue
        i_ls, i_cat = heads.index('Лицевой счёт'), heads.index('Категория')
        for row in rows[1:]:
            if not row or i_ls >= len(row) or i_cat >= len(row):
                continue
            cat = row[i_cat]
            if (cat not in (None, '') and str(cat).strip().lower()
                    in ('наниматель', 'собственник')):
                out[norm_ls(row[i_ls])] = str(cat).strip()
        wb.close()
    return out


def fmt_d(d):
    return f'{d:%d.%m.%Y}' if d else ''


def sost_on(d2):
    """По состоянию на = первый день месяца после конца периода (как в их шаблоне)."""
    if not d2:
        return None
    y, m = (d2.year + 1, 1) if d2.month == 12 else (d2.year, d2.month + 1)
    return date(y, m, 1)


def addr_full(raw):
    """'Пример,Тестовая,д.55кв.164' → 'Примерная область, Примерный район, г. Пример, ул. Тестовая, д. 55, кв. 164'.

    Если адрес уже развёрнут (есть «д. N» и «кв. N» или упомянут регион) —
    возвращаем как есть: разбирать заново нельзя, иначе «Примерная область,
    Примерный район, г. Пример, ул. Тестовая, д. 1, кв. 1» склеится в мусор.
    """
    raw = str(raw or '').strip()
    if not raw:
        return ''
    if re.search(r'\bд[.\s№]*\d', raw, re.IGNORECASE) and \
            re.search(r'\bкв[.\s№]*\d', raw, re.IGNORECASE):
        return re.sub(r'\s+', ' ', raw)
    parts = [p.strip() for p in raw.split(',') if p.strip()]
    city, street, house = (parts + ['', '', ''])[:3]
    hm = re.match(r'д[.№\s]*(\d+)(?:кв[.№\s]*(\d+))?$', house)
    dom, kv = (hm.group(1), hm.group(2)) if hm else (house, '')
    city_txt = city if re.search(r'\b(г|пгт|с|п|д)[.\s]', city) else f'г. {city}'
    street_txt = street
    if street and not re.search(r'(ул|пр|шоссе|ш[.\s]|пер|бул)', street, re.IGNORECASE):
        street_txt = f'ул. {street}'
    out = f'Примерная область, Примерный район, {city_txt}, {street_txt}, д. {dom}'
    if kv:
        out += f', кв. {kv}'
    return re.sub(r'\s+', ' ', out)


# ---------------------------------------------------------------- главный ----
def main():
    report = ['ИМПОРТ «Портянки» + Ф-9 → книга ввода', '']
    args = sys.argv[1:]
    # Папка от диалога «Откуда загрузить документы?»: ищем файлы в ней.
    source = IN_GUESS
    for a in args:
        if Path(a).is_dir():
            source = Path(a)
            break
    portyanka = next((Path(a) for a in args if a.lower().endswith('.xlsx')), None)
    f9_files = [Path(a) for a in args if Path(a).suffix.lower() in ('.pdf', '.png', '.jpg', '.jpeg', '.tif', '.tiff')]
    if not portyanka:
        cand = sorted({*source.glob('ledger*.xlsx'), *source.glob('Портянка*.xlsx')}) if source.exists() else []
        portyanka = cand[0] if cand else None
    if not f9_files and source.exists():
        f9_files = sorted({*source.glob('f9*.pdf'), *source.glob('Ф-9*.pdf')})
    # портянку ищем и рядом со выбранными сканами
    if portyanka is None and f9_files:
        parent = f9_files[0].parent
        cand = sorted({*parent.glob('ledger*.xlsx'), *parent.glob('Портянка*.xlsx')})
        if cand:
            portyanka = cand[0]
    if portyanka is None and not f9_files:
        sys.exit('Не выбрано ничего для загрузки: ни портянки (.xlsx), ни сканов Ф-9. '
                 'В диалоге выделите файлы — например, «Ф-9 … .pdf» — и откройте.')

    import importlib.util
    spec = importlib.util.spec_from_file_location('kniga_sozdat', BASE / 'create_workbook.py')
    kmod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(kmod)

    warn = []
    if portyanka:
        items, warns = load_portyanka(portyanka)
        warn += warns
        groups = collect_groups(items, warn)
        report.append(f'Портянка: «{portyanka.name}» → строк: {len(items)}, заявлений: {len(groups)}')
    else:
        items, warns, groups = [], [], []
        report.append('Портянка не выбрана — принимаю только сканы Ф-9: '
                      'справки привяжутся к своим ЛС, суммы и заявления не трогаются.')

    # индекс ФИО по ЛС: запасная привязка страниц справок, когда угол OCR размазал
    known_persons = {}
    for g in groups:
        ks = {k for k in person_keys(g['persons'])}
        known_persons.setdefault(g['ls'], set()).update(ks)
    f9_by_ls, unattached, texts_by_ls = load_f9_files(
        f9_files, report, warn, known_ls={g['ls'] for g in groups},
        known_persons=known_persons, port_index=build_port_index(groups))

    # Импорт пересобирает книгу с нуля, ручные правки в ней пропадут:
    # заранее сохраняем копию с датой, чтобы всегда было что вернуть.
    rebuild = portyanka is not None
    prior_wb = None
    if OUT_BOOK.exists():   # бэкап и при дописывании Ф-9 в готовую книгу — всегда есть откат
        stamp = datetime.now().strftime('%Y-%m-%d %H-%M-%S')
        bak = OUT_BOOK.with_name(f'housing-input-backup {stamp}.xlsx')
        shutil.copy2(OUT_BOOK, bak)
        report.append(f'Бэкап прежней книги: {bak.name}')
        # прежнюю книгу читаем ДО пересборки: из неё берём справочник ЛС
        # и ручные правки юриста (иначе они пропали бы вместе с книгой)
        try:
            prior_wb = openpyxl.load_workbook(OUT_BOOK, data_only=True)
        except Exception as ex:  # noqa: BLE001
            warn.append(f'не удалось прочитать прежнюю книгу ({ex}) — ручные правки не перенесены')

    if not rebuild and OUT_BOOK.exists():
        wb = openpyxl.load_workbook(OUT_BOOK)
    else:
        wb = kmod.build_book(include_tests=False)
        if 'Инструкция' in wb.sheetnames:
            wb['Инструкция']['A1'].value = 'КНИГА ВВОДА (сформирована импортом из «Портянки» и Ф-9)'
    ws_z, ws_d = wb['Заявления'], wb['Должники']
    n_claim, n_person = 0, 0

    # Справочник ЛС и ручные правки прежней книги переживают переимпорт.
    # Справочник читаем из ПРЕЖНЕЙ книги (её уже скопировали в бэкап выше) и
    # переносим в новую: иначе он стирался бы вместе с книгой.
    guide = read_guide(prior_wb) if prior_wb else {}
    prev = read_prev_book(OUT_BOOK) if rebuild else {}
    saved_cats = scan_saved_categories() if rebuild else {}   # третий запасной источник
    # правки юриста в самой справочнике (если он правил книгу вручную) — приоритетнее
    guide_now = read_guide(wb) if rebuild else {}
    for ls, rec in guide_now.items():
        merged = dict(guide.get(ls, {}))
        merged.update({k: v for k, v in rec.items() if v})
        guide[ls] = merged
    # и переносим справочник в новую книгу
    if guide and GUIDE_SHEET in wb.sheetnames:
        ws_g = wb[GUIDE_SHEET]
        if ws_g.max_row > 1:
            ws_g.delete_rows(2, ws_g.max_row)
        for ls in sorted(guide):
            rec = guide[ls]
            ws_g.append([ls, rec.get('Категория', ''), rec.get('Суд (шапка)', ''),
                         rec.get('Тип жилья (текст)', ''), rec.get('Заметки', '')])

    for g in groups:
        n = ws_z.max_row
        ls = g['ls']
        # категория: справочник → ручная правка прежней книги → реестры/бэкапы
        # прошлых сборов → пусто
        cat = (guide.get(ls, {}).get('Категория')
               or prev.get((g['typ'], ls), {}).get('Категория')
               or saved_cats.get(norm_ls(ls), ''))
        sud = guide.get(ls, {}).get('Суд (шапка)') or prev.get((g['typ'], ls), {}).get('Суд (шапка)') or ''
        if not cat:
            warn.append(f'ЛС {ls}: не указана категория (наниматель/собственник) — впишите в лист '
                        '«Справочник ЛС», заявление без неё не соберётся.')
        if not sud:
            sud = default_sud(g['typ'])
            if sud:
                warn.append(f'ЛС {ls}: суд подставлен по типу документа ({g["typ"]}) — проверьте участок.')
        housing = (guide.get(ls, {}).get('Тип жилья (текст)')
                   or prev.get((g['typ'], ls), {}).get('Тип жилья (текст)') or '')
        cena = prev.get((g['typ'], ls), {}).get('Цена иска, руб.') or ''
        otmena = prev.get((g['typ'], ls), {}).get('Отмена СП (да/нет)') or 'нет'
        ws_z.append([n, g['typ'], cat, ls, sud, addr_full(g['addr']), housing,
                     fmt_d(sost_on(g['d2'])), fmt_d(g['d1']), fmt_d(g['d2']), g['dolg'],
                     fmt_d(g['d1']), fmt_d(g['d2']), g['peni'], g['gp'], cena, otmena])
        n_claim += 1

        f9 = f9_by_ls.get(g['ls'], [])
        rows = []
        clean = lambda s: re.sub(r'\s*\([^)]*\)', '', s).lower().split()[0]
        f9_keys = {(clean(q['fio']), q['init']) for q in f9}
        if f9:
            # полные ФИО портянки достраивают инициалы из Ф-9
            for p in f9:
                nm = p['fio']
                if not re.match(r'^[А-ЯЁ][а-яё\-]+ [А-ЯЁ][а-яё\-]+ [А-ЯЁ][а-яё\-]', nm):
                    full = next((c['full'] for c in g['persons']
                                 if clean(c['surname']) == clean(nm) and
                                 (c['init'] == p['init'] or c['init'] == '?')), None)
                    if full:
                        nm = full
                pp = f'{p["pasport"]}, выдан {p["kem"]}' if p['pasport'] and p['kem'] else p['pasport']
                rows.append([g['ls'], nm, p['dr'], p['place'], pp, ''])  # Пол — пусто
                n_person += 1
            for p in g['persons']:
                key = (re.sub(r'\s*\([^)]*\)', '', p['surname']).lower(), p['init'])
                if key not in f9_keys and clean(p['surname']) not in {k[0] for k in f9_keys}:
                    warn.append(f'ЛС {g["ls"]}: из портянки «{p["full"] or p["surname"] + " " + p["init"]}.» '
                                'не найден в Ф-9 — проверьте, входит ли в документ.')
        else:
            if g['persons']:
                for p in g['persons']:
                    nm = p['full'] or f'{p["surname"]} {p["init"][0]}.{p["init"][1]}.'
                    rows.append([g['ls'], nm, '', '', '', ''])
                    n_person += 1
                warn.append(f'ЛС {g["ls"]}: нет Ф-9 — ФИО из портянки; дата рождения и паспорт НЕ заполнены.')
            else:
                warn.append(f'ЛС {g["ls"]}: В ПОРТЯНКЕ ПУСТАЯ КОЛОНКА «ФИО» — должников нет! '
                            'Добавьте их вручную на лист «Должники» (или это ЛС без жилого помещения?).')
        for r in rows:
            ws_d.append(r)

    if unattached:
        ws_un = wb['Ф-9 не привязаны'] if 'Ф-9 не привязаны' in wb.sheetnames else wb.create_sheet('Ф-9 не привязаны')
        if ws_un.max_row > 1:   # лист перезаписывается целиком — иначе копятся дубли заголовков
            ws_un.delete_rows(1, ws_un.max_row)
        ws_un.append(['Файл', 'Страница', 'ФИО полностью', 'Дата рождения',
                      'Место рождения', 'Паспорт (кем и когда выдан)', 'Почему не привязано'])
        for c in ws_un[ws_un.max_row]:
            c.font, c.fill = openpyxl.styles.Font(bold=True), openpyxl.styles.PatternFill('solid', fgColor='DDEBF7')
        pages_un = sorted({u['page'] for u in unattached})
        for u in unattached:
            reason = u.get('reason') or 'причина не определена'
            if not u['persons']:
                ws_un.append([u['file'], u['page'], '', '', '', '', reason])
                continue
            for p in u['persons']:
                pp = f'{p["pasport"]}, выдан {p["kem"]}' if p['pasport'] and p['kem'] else p['pasport']
                ws_un.append([u['file'], u['page'], p['fio'], p['dr'], p['place'], pp, reason])
                n_person += 1
        for i, wdt in enumerate((24, 9, 30, 14, 28, 44, 60), 1):
            ws_un.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wdt
        for u in unattached:
            if u.get('reason'):
                warn.append(f'Ф-9 «{u["file"]}», стр. {u["page"]}: {u["reason"]}.')
        warn.append(f'Ф-9 «{unattached[0]["file"]}»: страницы {pages_un} без ЛС. '
                    'Два способа привязать: 1) скан, где каждой справке соответствует своя '
                    'страница — впишите ЛС в «f9-index.xlsx» и перезапустите импорт; '
                    '2) сканами по одному файлу на справку с именем «Ф-9 ЛС<номер>.pdf».')
        if 'f9-index.xlsx' not in {p.name for p in BASE.iterdir()}:
            idx = openpyxl.Workbook()
            ws_i = idx.active
            ws_i.title = 'индекс'
            ws_i.append(['Файл', 'Страница', 'ЛС'])
            for u in unattached:
                ws_i.append([u['file'], u['page'], ''])
            for i, wdt in enumerate((40, 10, 16), 1):
                ws_i.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wdt
            idx.save(BASE / 'f9-index.xlsx')
            warn.append('Создан «f9-index.xlsx» — впишите ЛС напротив каждой страницы '
                        'и перезапустите импорт (python3 import_housing.py).')

    if texts_by_ls:
        ws_f = wb['Ф-9 справки'] if 'Ф-9 справки' in wb.sheetnames else wb.create_sheet('Ф-9 справки')
        # ручная правка графы юристом — сохраняем по ключу (ЛС, файл, стр.)
        keep_graph = {}
        if ws_f.max_row > 1:
            h_old = [str(c.value).strip() if c.value else '' for c in ws_f[1]]
            g_head = 'Графа «собственник и/или наниматель» (через «;»; для СП/ИСК)'
            try:
                i_ls_o, i_f_o, i_p_o = (h_old.index('Лицевой счёт'), h_old.index('Файл'),
                                        h_old.index('Стр.'))
            except ValueError:
                i_ls_o = i_f_o = i_p_o = -1
            i_g_o = h_old.index(g_head) if g_head in h_old else -1
            if i_ls_o >= 0 and i_g_o >= 0:
                for row in ws_f.iter_rows(min_row=2, values_only=True):
                    if not row:
                        continue
                    gv = row[i_g_o] if i_g_o < len(row) else None
                    if gv is None or not str(gv).strip():
                        continue
                    keep_graph[(str(row[i_ls_o] or ''), str(row[i_f_o] or ''),
                                str(row[i_p_o] or ''))] = str(gv).strip()
            ws_f.delete_rows(2, ws_f.max_row - 1)   # лист генерируется импортом целиком
        heads_f = ['Лицевой счёт', 'Файл', 'Стр.',
                   'Текст справки Ф-9 (распознано автоматически — сверить с оригиналом)',
                   'Графа «собственник и/или наниматель» (через «;»; для СП/ИСК)',
                   'Как привязано к лицевому счёту']
        for j, h in enumerate(heads_f, 1):
            ws_f.cell(row=1, column=j, value=h)
        row_f = 1
        n_graph = 0
        for ls in sorted(texts_by_ls, key=str):
            for fname, n, t, graph, how in texts_by_ls[ls]:
                row_f += 1
                ws_f.cell(row=row_f, column=1, value=ls)
                ws_f.cell(row=row_f, column=2, value=fname)
                ws_f.cell(row=row_f, column=3, value=n)
                ws_f.cell(row=row_f, column=4, value=t.strip())
                gv = '; '.join(graph) if graph else ''
                if not gv:
                    gv = keep_graph.get((str(ls), str(fname), str(n)), '')
                if gv:
                    ws_f.cell(row=row_f, column=5, value=gv)
                    n_graph += 1
                if how:
                    ws_f.cell(row=row_f, column=6, value=how)
        for c in ws_f[1]:
            c.font, c.fill = openpyxl.styles.Font(bold=True), \
                openpyxl.styles.PatternFill('solid', fgColor='DDEBF7')
        for i, wdt in enumerate((14, 26, 6, 110, 40, 34), 1):
            ws_f.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wdt
        if n_graph:
            report.append(f'Графа «собственник/наниматель» распознана на {n_graph} справк(ах) — '
                          'по ней: СП отдельно на каждое лицо, в иcк — собственник + совершеннолетние.')
        report.append(f'Справок Ф-9 привязано: {sum(len(v) for v in texts_by_ls.values())} '
                      f'по {len(texts_by_ls)} ЛС (лист «Ф-9 справки»)')

    # справки, привязанные к ЛС, которых нет в портянке, в документы не попадут — сообщаем
    orphan_ls = sorted((ls for ls in f9_by_ls
                        if str(ls) not in {str(g['ls']) for g in groups}), key=str)
    if orphan_ls:
        warn.append('ЛС со справками Ф-9, которых НЕТ в портянке — привязаны, но в документы '
                    'не войдут: ' + ', '.join(orphan_ls) + '. Проверьте портянку (или это чужие справки).')

    ws_s = wb['Сверка импорта'] if 'Сверка импорта' in wb.sheetnames else wb.create_sheet('Сверка импорта')
    if ws_s.max_row > 1:   # лист перезаписывается целиком — иначе копятся дубли заголовков
        ws_s.delete_rows(1, ws_s.max_row)
    ws_s.append(['НУЖНО ЗАПОЛНИТЬ РУКАМИ:'])
    ws_s.append(['• Категория (наниматель/собственник) — вписывается один раз в лист «Справочник ЛС»,'])
    ws_s.append(['  оттуда подставляется при каждом импорте (в портянке и в Ф-9 её нет).'])
    ws_s.append(['• Суд (шапка) — подставляется по типу документа; свой указывается в «Справочник ЛС».'])
    ws_s.append(['• Справки Ф-9, не привязанные к ЛС, — лист «Ф-9 не привязаны» (с причиной).'])
    ws_s.append([])
    ws_s.append(['ПРЕДУПРЕЖДЕНИЯ СВЕРКИ:'])
    for w in sorted(set(warn)):
        ws_s.append(['• ' + w])
    ws_s.column_dimensions['A'].width = 130
    for c in ('A1',):
        ws_s[c].font = openpyxl.styles.Font(bold=True, size=13)
    ws_d.column_dimensions['F'].width = 28

    wb.save(OUT_BOOK)
    report += ['', f'Записано заявлений: {n_claim}; строк должников: {n_person}',
               f'Предупреждений сверки: {len(set(warn))}',
               f'Книга сохранена: {OUT_BOOK}']
    print('\n'.join('\n'.join(r) if isinstance(r, list) else r for r in report))


if __name__ == '__main__':
    platform_utils.fix_console()
    not_found = platform_utils.missing_tools()
    if not_found:
        print('Для импорта не хватает программ:\n  • '
              + '\n  • '.join(not_found)
              + '\n\nУстановите их по инструкции пакета Сова-юрист и повторите.')
        sys.exit(2)
    main()
