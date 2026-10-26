#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор заявлений о выдаче судебных приказов и исковых заявлений
о взыскании задолженности за ЖКУ.

Использование:
    python3 generate.py [input-workbook.xlsx]

Читает книгу ввода (листы «Настройки», «Заявления», «Должники»),
проверяет данные на очевидные ошибки (сверка) и формирует:
  - готовые .docx заявления (по одному файлу на лицевой счёт);
  - registry.xlsx со всеми документами и итогами;
  - validation-report.txt с найденными ошибками и предупреждениями.
"""
import re
import sys
from datetime import date, datetime
from pathlib import Path

import openpyxl
import platform_utils
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

BASE = Path(__file__).resolve().parent

# ---------------------------------------------------------------- числа ----
UNITS_M = 'один два три четыре пять шесть семь восемь девять'.split()
UNITS_F = 'одна две три четыре пять шесть семь восемь девять'.split()
TEENS = ('десять одиннадцать двенадцать тринадцать четырнадцать '
         'пятнадцать шестнадцать семнадцать восемнадцать девятнадцать').split()
TENS = ['', '', 'двадцать', 'тридцать', 'сорок', 'пятьдесят',
        'шестьдесят', 'семьдесят', 'восемьдесят', 'девяносто']
HUNDREDS = ['', 'сто', 'двести', 'триста', 'четыреста', 'пятьсот',
            'шестьсот', 'семьсот', 'восемьсот', 'девятьсот']


def plural(n, one, few, many):
    if n % 100 in range(11, 20) or n % 10 in (0, 5, 6, 7, 8, 9):
        return many
    if n % 10 == 1:
        return one
    return few


def _words_upto_1000(n, fem=False):
    out = []
    h, t = divmod(n, 100)
    if h:
        out.append(HUNDREDS[h])
    if t // 10 == 1:
        out.append(TEENS[t % 10])
    else:
        if t // 10 >= 2:
            out.append(TENS[t // 10])
        if t % 10:
            out.append((UNITS_F if fem else UNITS_M)[t % 10 - 1])
    return out


def number_words(n):
    """Целое число словами по-русски."""
    if n == 0:
        return 'ноль'
    out = []
    millions, n = divmod(n, 1_000_000)
    thousands, n = divmod(n, 1_000)
    if millions:
        out += _words_upto_1000(millions)
        out.append(plural(millions, 'миллион', 'миллиона', 'миллионов'))
    if thousands:
        out += _words_upto_1000(thousands, fem=True)
        out.append(plural(thousands, 'тысяча', 'тысячи', 'тысяч'))
    if n:
        out += _words_upto_1000(n)
    return ' '.join(out)


def sum_words(x):
    """Сумма в стиле шаблонов: 7237 (Семь тысяч двести тридцать семь) рублей 25 копеек."""
    cents = int(round(float(x) * 100))
    r, k = divmod(cents, 100)
    w = number_words(r).capitalize()
    return (f'{r} ({w}) {plural(r, "рубль", "рубля", "рублей")} '
            f'{k:02d} {plural(k, "копейка", "копейки", "копеек")}')


def fmt_num(x):
    """7237.25 → '7237,25'"""
    cents = int(round(float(x) * 100))
    return f'{cents // 100},{cents % 100:02d}'


def fmt_price(x):
    """7237.25 → '7237,25 рублей'"""
    cents = int(round(float(x) * 100))
    r = cents // 100
    return f'{fmt_num(x)} {plural(r, "рубль", "рубля", "рублей")}'


def surname_genitive(surname, fem):
    """Родительный падеж фамилии: Иванов→Иванова, Петрова→Петровой."""
    s = surname.strip()
    low = s.lower()
    if fem:
        if low.endswith(('ая',)):
            return s[:-2] + 'ой'
        if low.endswith(('яя', 'ь')):
            return s[:-2] + 'ей' if low.endswith('яя') else s
        if low.endswith(('а', 'я')):
            return s[:-1] + 'ой'
        return s  # -их/-ых и прочие не склоняются
    else:
        if low.endswith(('ов', 'ев', 'ёв', 'ин', 'ын')):
            return s + 'а'
        if low.endswith(('ский', 'цкий')):
            return s[:-2] + 'ого'
        if low.endswith('ой'):
            return s[:-2] + 'ого'
        return s  # не склоняются (заканчиваются на согласную без -ов/-ин и т.п.)


PAREN_RE = re.compile(r'\s*\([^()]*\)')


def clean_fio(fio):
    """Убирает девичью фамилию в скобках: «Романова (Сурмава) И.Г.» → «Романова И.Г.»."""
    return PAREN_RE.sub('', fio).strip()


def fio_short(fio, fem):
    """Иванов Иван Иванович → Иванова И.И. (в родительном падеже для ПРОШУ)."""
    parts = clean_fio(fio).split()
    if len(parts) < 3:   # лицо графы Ф-9 без полных данных в «Должники» — оставляем как есть
        return fio
    fam = surname_genitive(parts[0], fem)
    return f'{fam} {parts[1][0]}.{parts[2][0]}.'


def fio_nom_short(fio):
    """Иванов Иван Иванович → Иванов И.И. (им. падеж, для имён файлов/реестра)."""
    parts = clean_fio(fio).split()
    if len(parts) >= 3:
        return f'{parts[0]} {parts[1][0]}.{parts[2][0]}.'
    return ' '.join(parts)


# ---------------------------------------------------------------- чтение ----
def parse_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if v is None or str(v).strip() == '':
        return None
    s = str(v).strip().strip(' .г')
    for f in ('%d.%m.%Y', '%d.%m.%y', '%Y-%m-%d'):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def fmt_date(d):
    return f'{d:%d.%m.%Y}' if d else ''


def parse_sum(v):
    if v is None or str(v).strip() == '':
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(' ', ' ').replace(' ', '').replace(',', '.')
    s = re.sub(r'[^0-9.]', '', s)
    try:
        return float(s) if s else None
    except ValueError:
        return None


def cell_str(v):
    if v is None:
        return ''
    return str(v).strip()


CYR = re.compile('[а-яёА-ЯЁ]')
LAT = re.compile('[a-zA-Z]')
FIO_RE = re.compile(r'^[А-ЯЁ][а-яё\-]+(?:\s*\([^()]{1,40}\))? [А-ЯЁ][а-яё\-]+ [А-ЯЁ][а-яё\-]+$')


def load_workbook_data(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    for name in ('Настройки', 'Заявления', 'Должники'):
        if name not in wb.sheetnames:
            sys.exit(f'В книге нет листа «{name}». Пересоздайте её: python3 create_workbook.py')

    cfg = {}
    for row in wb['Настройки'].iter_rows(min_row=2, values_only=True):
        if row and cell_str(row[0]):
            cfg[cell_str(row[0]).lower().replace(' ', '_')] = cell_str(row[1] if len(row) > 1 else None)

    claims, debtors = [], []
    ws = wb['Заявления']
    headers = [cell_str(c.value) for c in ws[1]]
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or all(v is None or str(v).strip() == '' for v in row):
            continue
        d = dict(zip(headers, row))
        d['_row'] = i
        claims.append(d)
    ws = wb['Должники']
    headers = [cell_str(c.value) for c in ws[1]]
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or all(v is None or str(v).strip() == '' for v in row):
            continue
        d = dict(zip(headers, row))
        d['_row'] = i
        debtors.append(d)
    f9 = {}
    if 'Ф-9 справки' in wb.sheetnames:   # привязанные сканы справок (из импорта «Портянки»)
        ws = wb['Ф-9 справки']
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not cell_str(row[0] if row else None):
                continue
            f9.setdefault(cell_str(row[0]), []).append({
                'file': cell_str(row[1] if len(row) > 1 else None),
                'page': cell_str(row[2] if len(row) > 2 else None),
                'text': cell_str(row[3] if len(row) > 3 else None),
                'graph': cell_str(row[4] if len(row) > 4 else None)})
    return cfg, claims, debtors, f9


# ------------------------------------------------ графы Ф-9 (СП: только они) ----
def fio_key(s):
    """Ключ сравнения ФИО: «Гаврилов Пётр Степанович» == «Гаврилов П.С.» == «гаврилов пс»."""
    parts = [x for x in re.split(r'[\s,;]+', str(s)) if x]
    if not parts:
        return ''
    inits = ''
    for tok in parts[1:4]:
        if '.' in tok:
            inits += tok.replace('.', '').lower()
        else:
            inits += tok[:1].lower()
    return parts[0].lower().replace('\u00b7', '') + '_' + inits


def age_years(dr, on=None):
    on = on or date.today()
    return on.year - dr.year - ((on.month, on.day) < (dr.month, dr.day))


def fem_heuristic(fio):
    f = str(fio).split()[0].lower() if str(fio).split() else ''
    return f.endswith(('а', 'я'))


def resolve_graph_person(gname, dlist, warnings, ls):
    """Лицо из графы Ф-9 → словарь person; паспорт/др.р. подтягиваются из листа «Должники»."""
    key = fio_key(gname)
    full, dr, place, pasport = gname, None, '', ''
    for dd in dlist:
        row_key = fio_key(cell_str(dd.get('ФИО полностью')))
        if row_key == key:
            full = cell_str(dd.get('ФИО полностью'))
            dr_raw = dd.get('Дата рождения')
            dr = parse_date(dr_raw)
            place = cell_str(dd.get('Место рождения'))
            pasport = cell_str(dd.get('Паспорт (кем и когда выдан)'))
            break
    if full != gname:
        pass
    else:
        warnings.append(f'ЛС {ls}: лицо из графы Ф-9 «{gname}» не найдено на листе «Должники» — '
                        'допишите его там полностью (дата рождения, паспорт) и перезапустите сбор.')
    return {'fio': full, 'fem': fem_heuristic(full), 'dr': dr,
            'place': place, 'pasport': pasport}


# ---------------------------------------------------------------- сверка ----
def normalize_claims(claims, debtors, errors, warnings, f9=None):
    norm = []
    seen_ls = {}
    by_ls_deb = {}
    for d in debtors:
        by_ls_deb.setdefault(cell_str(d.get('Лицевой счёт')), []).append(d)

    bad_claim = {}   # № заявления → [ошибки] — заявления с ошибками не собираются,
                     # остальные формируются без них (частичная генерация)

    for c in claims:
        row = c['_row']
        n = cell_str(c.get('№')) or str(len(norm) + 1)
        claim_errors = []

        def err(msg):
            claim_errors.append(msg)

        typ = cell_str(c.get('Тип документа')).upper()
        cat = cell_str(c.get('Категория')).lower()
        ls = cell_str(c.get('Лицевой счёт'))
        if typ not in ('СП', 'ИСК'):
            err(f'Заявление {n} (строка {row}): тип документа «{typ or "не указан"}» — укажите СП или ИСК.')
        if cat not in ('наниматель', 'собственник'):
            err(f'Заявление {n} (строка {row}): категория «{cat or "не указана"}» — укажите наниматель или собственник.')
        if not ls:
            err(f'Заявление {n} (строка {row}): не указан лицевой счёт.')
        elif ls in seen_ls:
            err(f'Заявление {n} (строка {row}: ЛС {ls}) повторяет строку {seen_ls[ls]} — лицевой счёт используется в двух заявлениях, разберитесь до формирования.')
        else:
            seen_ls[ls] = row
        if not cell_str(c.get('Суд (шапка)')):
            err(f'Заявление {n} (строка {row}): не указан суд.')
        if not cell_str(c.get('Адрес регистрации')):
            err(f'Заявление {n} (строка {row}): не указан адрес регистрации.')

        dates, sums = {}, {}
        for col, key in (('По состоянию на', 'sost'), ('Долг: с', 'd1'), ('Долг: по', 'd2'),
                         ('Пени: с', 'p1'), ('Пени: по', 'p2')):
            dates[key] = parse_date(c.get(col))
            if c.get(col) not in (None, '') and dates[key] is None:
                err(f'Заявление {n} (строка {row}): не распознана дата в поле «{col}»: «{c.get(col)}». Формат: ДД.ММ.ГГГГ.')
        for col, key in (('Долг, руб.', 'dolg'), ('Пени, руб.', 'peni'), ('Госпошлина, руб.', 'poshlina')):
            sums[key] = parse_sum(c.get(col))
            if sums[key] is None:
                err(f'Заявление {n} (строка {row}): не распознана сумма в поле «{col}»: «{c.get(col)}».')

        if dates.get('d1') and dates.get('d2') and dates['d1'] > dates['d2']:
            err(f'Заявление {n} (строка {row}): «Долг: с» позже «Долг: по».')
        if dates.get('p1') and dates.get('p2') and dates['p1'] > dates['p2']:
            err(f'Заявление {n} (строка {row}): «Пени: с» позже «Пени: по».')
        if dates.get('p2') and dates.get('sost') and dates['p2'] > dates['sost']:
            warnings.append(f'Заявление {n} (строка {row}): конец периода пени позже даты «По состоянию на».')
        for col in ('По состоянию на', 'Долг: с', 'Долг: по'):
            if c.get(col) in (None, ''):
                err(f'Заявление {n} (строка {row}): не заполнено обязательное поле «{col}».')
        if (sums.get('peni') or 0) > 0 and (dates.get('p1') is None or dates.get('p2') is None):
            err(f'Заявление {n} (строка {row}): указано пени, но не заполнен период пени («Пени: с» / «Пени: по»).')
        if dates.get('d2') and dates.get('sost') and dates['d2'] > dates['sost']:
            warnings.append(f'Заявление {n} (строка {row}): конец периода долга позже даты «По состоянию на».')

        if typ == 'СП' and sums.get('dolg') and sums['dolg'] + (sums.get('peni') or 0) > 500_000:
            err(f'Заявление {n} (строка {row}): долг+пени превышают 500 000 ₽ — судебный приказ невозможен (ст. 121 ГПК РФ), нужен иск.')

        cena = parse_sum(c.get('Цена иска, руб.'))
        if cena is not None and sums.get('dolg') is not None:
            auto = sums['dolg'] + (sums.get('peni') or 0)
            if abs(cena - auto) > 0.011:
                warnings.append(f'Заявление {n} (строка {row}): указанная цена иска {auto:.2f} ≠ долг + пени ({cena:.2f}). Проверьте.')

        # лица из графы «собственник/наниматель» справки Ф-9 (объединяем все её строки)
        graphs = []
        for ent in (f9 or {}).get(ls, []):
            if ent.get('graph'):
                graphs += [g.strip() for g in str(ent['graph']).split(';') if g.strip()]
        graphs = list(dict.fromkeys(graphs))
        sp_graph_mode = typ == 'СП' and bool(graphs)
        isk_graph_mode = typ == 'ИСК' and bool(graphs)
        if sp_graph_mode and len(graphs) > 1:
            warnings.append(f'Заявление {n} (ЛС {ls}): {len(graphs)} лиц(а) в графе Ф-9 — '
                            'судебный приказ формируется ОТДЕЛЬНО на каждого.')
        if isk_graph_mode:
            warnings.append(f'Заявление {n} (ЛС {ls}): из Ф-9 взяты собственник/наниматель '
                            '(графа) и совершеннолетние зарегистрированные — одним иском.')
        graph_keys = {fio_key(g) for g in graphs}

        # должники по ЛС
        dlist = by_ls_deb.get(ls, [])
        if not dlist:
            if sp_graph_mode or isk_graph_mode:
                warnings.append(f'Заявление {n} (ЛС {ls}): на листе «Должники» нет строк — '
                                'лица взяты только из графы Ф-9, паспортные данные не заполнены. '
                                'Перезапустите импорт, чтобы дописать.')
            else:
                err(f'Заявление {n} (строка {row}): по ЛС «{ls}» на листе «Должники» нет ни одной записи.')
        persons = []
        fios = set()
        for dd in dlist:
            fio = re.sub(r'\s+', ' ', cell_str(dd.get('ФИО полностью')))
            pol = cell_str(dd.get('Пол')).lower()[:1]
            fem = pol == 'ж' or (pol != 'м' and fio.split()[0].lower().endswith(('а', 'я'))) if fio else False
            graph_mode = sp_graph_mode or isk_graph_mode
            used = not graph_mode or fio_key(fio) in graph_keys
            if not fio:
                err(f'Лист «Должники», строка {dd["_row"]}: пустое ФИО (ЛС «{ls}»).')
            elif len(fio.split()) < 3 and used:
                err(f'Лист «Должники», строка {dd["_row"]}: ФИО «{fio}» — должно быть «Фамилия Имя Отчество». Краткое ФИО строится автоматически, оно должно быть точным.')
            elif not FIO_RE.fullmatch(fio) and used:
                err(f'Лист «Должники», строка {dd["_row"]}: странные символы или регистр в ФИО «{fio}» (пример: Иванов Иван Иванович).')
            else:
                if CYR.search(fio) and LAT.search(fio):
                    warnings.append(f'Лист «Должники», строка {dd["_row"]}: в ФИО «{fio}» смешаны русские и латинские буквы — проверьте.')
                dr = parse_date(dd.get('Дата рождения'))
                if dr is None:
                    # где искать: в справке Ф-9 по этому ЛС (страница) или в самой справке
                    src = ''
                    for ent in (f9 or {}).get(ls, []):
                        if ent.get('text') and fio.split()[0][:4].lower() in ent['text'].lower():
                            src = f' (см. справку Ф-9, стр. {ent.get("page") or "?"})'
                            break
                    msg = (f'Лист «Должники», строка {dd["_row"]}: не распознана дата рождения '
                           f'«{dd.get("Дата рождения")}» (формат ДД.ММ.ГГГГ)'
                           + (f'; ЛС {ls}{src}' if ls else '') + '.')
                    if not graph_mode or used:
                        err(msg)
                    else:
                        warnings.append(f'ЛС {ls}: у «{fio}» нет даты рождения, но лицо не входит '
                                        f'в {typ} (по графе Ф-9) — на сбор не влияет.')
                person = {'fio': fio, 'fem': fem, 'dr': dr,
                          'place': cell_str(dd.get('Место рождения')),
                          'pasport': cell_str(dd.get('Паспорт (кем и когда выдан)'))}
                if fio in fios:
                    # одна и та же строка приходит и из портянки, и из таблицы Ф-9 —
                    # одну запись дополняем недостающими данными второй, в документе не дублируем
                    warnings.append(f'Заявление {n}: должник «{fio}» указан в лицевом счёте дважды — строки объединены.')
                    dup = persons[-1] if persons and persons[-1]['fio'] == fio else \
                        next((q for q in reversed(persons) if q['fio'] == fio), None)
                    if dup:
                        for kk, vv in person.items():
                            if not dup.get(kk) and vv:
                                dup[kk] = vv
                    continue
                fios.add(fio)
                persons.append(person)
        base = {'num': n, 'row': row, 'typ': typ, 'cat': cat, 'ls': ls,
                'sud': cell_str(c.get('Суд (шапка)')).split('\n'),
                'addr': cell_str(c.get('Адрес регистрации')),
                'housing': cell_str(c.get('Тип жилья (текст)')) or 'отдельной квартиры',
                'sost': dates.get('sost'), 'd1': dates.get('d1'), 'd2': dates.get('d2'),
                'p1': dates.get('p1'), 'p2': dates.get('p2'),
                'dolg': sums.get('dolg'), 'peni': sums.get('peni') or 0.0,
                'poshlina': sums.get('poshlina'),
                'otmena': cell_str(c.get('Отмена СП (да/нет)')).lower().startswith('да'),
                'f9_texts': (f9 or {}).get(ls, [])}
        if sp_graph_mode:
            # СП (по указанию юриста): приказ на КАЖДОЕ лицо графы «собственник/наниматель»,
            # зарегистрированные НЕ входят; приказов ровно столько, сколько лиц в графе
            for gi, g in enumerate(graphs, 1):
                match = next((q for q in persons if fio_key(q['fio']) == fio_key(g)), None)
                p = match if match else resolve_graph_person(g, dlist, warnings, ls)
                norm.append({**base, 'persons': [p], 'part': gi, 'parts': len(graphs)})
            if claim_errors:
                bad_claim[n] = claim_errors
            continue
        if isk_graph_mode:
            # ИСК (по указанию юриста): собственник/наниматель из графы + совершеннолетние
            # зарегистрированные — ВСЕ в одном документе; несовершеннолетние исключаются
            persons_isk, have = [], set()
            for g in graphs:
                match = next((q for q in persons if fio_key(q['fio']) == fio_key(g)), None)
                if match:
                    persons_isk.append(match)
                else:
                    persons_isk.append(resolve_graph_person(g, dlist, warnings, ls))
                have.add(fio_key(g))
            for q in persons:
                k = fio_key(q['fio'])
                if k in have:
                    continue
                have.add(k)
                if q['dr'] is None:
                    persons_isk.append(q)
                    warnings.append(f'ЛС {ls}: у «{q["fio"]}» нет даты рождения — включён в иск, '
                                    'проверьте возраст вручную.')
                elif age_years(q['dr']) >= 18:
                    persons_isk.append(q)
                else:
                    warnings.append(f'ЛС {ls}: исключён из иска (несовершеннолетний): '
                                    f'{q["fio"]}, {q["dr"].strftime("%d.%m.%Y")}.')
            norm.append({**base, 'persons': persons_isk})
            if claim_errors:
                bad_claim[n] = claim_errors
            continue
        norm.append({**base, 'persons': persons})
        if claim_errors:
            bad_claim[n] = claim_errors
    return norm, bad_claim


# ---------------------------------------------------------------- текст ----
def intro_text(vz):
    return (f'В силу своей деятельности {vz} на основании заключенных договоров с поставщиками '
            'жилищных и коммунальных услуг производит начисление и сбор платы за ЖКУ, а также '
            'ведет работу по взысканию задолженности за ЖКУ путем подачи заявлений на выдачу '
            'судебных приказов и исковых заявлений от имени и в интересах поставщиков.')


def housing_text(typ, cat, housing, n_persons):
    person = 'Должник' if typ == 'СП' else 'Ответчик'
    plural_person = 'Должники' if typ == 'СП' else 'Ответчики'
    who = f'{person} является' if n_persons == 1 else f'{plural_person} являются'
    form = 'нанимателем' if n_persons == 1 else 'нанимателями'
    if cat == 'собственник':
        form = 'собственником' if n_persons == 1 else 'собственниками'
    return f'{who} {form} {housing} по вышеуказанному адресу.'


def legal_block(cat):
    if cat == 'наниматель':
        return [
            'На основании п. 5 ч. 3 ст. 67 ЖК РФ и ст. 678 Гражданского кодекса РФ наниматель жилого '
            'помещения по договору социального найма обязан своевременно вносить плату за жилое '
            'помещение и коммунальные услуги.',
            'Плата за жилое помещение должна вноситься нанимателем в сроки, предусмотренные договором '
            'найма жилого помещения. Если договором сроки не предусмотрены, то, как следует из п. 3 ст. 682 '
            'ГК РФ, плата должна вноситься нанимателем ежемесячно в порядке, установленном ЖК РФ.',
            'В соответствии с ч. 1 ст. 153 ЖК РФ граждане и организации обязаны своевременно и полностью '
            'вносить плату за жилое помещение и коммунальные услуги.',
            'В соответствии со ст. 155 Жилищного кодекса РФ, плата за коммунальные услуги вносится '
            'ежемесячно до десятого числа месяца, следующего за истекшим месяцем.',
            'В соответствии со ст. 156-157 Жилищного кодекса РФ, размер платы за жилье и коммунальные '
            'услуги устанавливается и рассчитывается по тарифам, установленным органами местного '
            'самоуправления в порядке, установленном федеральным законом.',
            'Согласно статье 69 Жилищного кодекса РФ члены семьи нанимателя жилого помещения по '
            'договору социального найма имеют равные с нанимателем права и обязанности и несут '
            'солидарную ответственность по обязательствам, вытекающим из договора социального найма.',
        ]
    return [
        'В соответствии с п.1 ст. 153 Жилищного Кодекса РФ, граждане обязаны своевременно и полностью '
        'вносить плату за коммунальные услуги.',
        'На основании ч. 3 ст. 30 Жилищного кодекса РФ собственник жилого помещения несет бремя '
        'содержания данного помещения и, если данное помещение является квартирой, общего имущества '
        'собственников помещений в соответствующем многоквартирном доме.',
        'Согласно ч. 3 ст. 154 ЖК РФ, собственники жилых домов несут расходы на их содержание и ремонт, '
        'а также оплачивают коммунальные услуги в соответствии с договорами, заключенными с лицами, '
        'осуществляющими соответствующие виды деятельности.',
        'В соответствии со ст. 155 Жилищного кодекса РФ, плата за коммунальные услуги вносится '
        'ежемесячно до десятого числа месяца, следующего за истекшим месяцем.',
        'В соответствии со ст. 156-157 Жилищного кодекса РФ, размер платы за жилье и коммунальные '
        'услуги устанавливается и рассчитывается по тарифам, установленным в порядке законодательства РФ.',
    ]


PERENI = (
    'В соответствии с ч. 14 ст. 155 Жилищного кодекса РФ, за нарушение сроков внесения платы за '
    'жилое помещение и коммунальные услуги с должника взимается пеня в размере одной трехсотой '
    'ставки рефинансирования (ключевой ставки) Центрального банка Российской Федерации от не '
    'выплаченной в срок суммы за каждый день просрочки, начиная с тридцать первого дня, '
    'следующего за днем наступления установленного срока оплаты.',
    'Согласно п. 1 ст. 330 Гражданского кодекса РФ неустойкой (штрафом, пеней) признается '
    'определенная законом денежная сумма, которую должник обязан уплатить кредитору в случае '
    'неисполнения или ненадлежащего исполнения обязательства, в частности при просрочке исполнения.',
)

SOLID = (
    'На основании ст. 322 ГК РФ солидарная обязанность (ответственность) или солидарное требование '
    'возникает, если солидарность обязанности или требования предусмотрена договором или установлена '
    'законом, в частности при неделимости предмета обязательства.',
    'Вместе с тем кредитор, не получивший полного удовлетворения от одного из солидарных должников, '
    'имеет право требовать недополученное от остальных солидарных должников, что предусмотрено '
    'п. 2 ст. 323 ГК РФ.',
)

ST309 = ('Статьей 309 Гражданского кодекса РФ установлено, что обязательства должны исполняться '
         'надлежащим образом в соответствии с условиями обязательства и требованиями закона.')

RUKOVODSTVO = {
    ('СП', 'наниматель'): 'Руководствуясь ст.ст. 309, 330, 678, 682 ГК РФ, ст.ст. 67, 153, 155 ЖК РФ, ст.ст. 121, 122, 131, 132 ГПК РФ',
    ('СП', 'собственник'): 'Руководствуясь ст.ст. 309, 330, 540, 548 ГК РФ, ст.ст. 30, 31, 69, 154, 155 ЖК РФ, ст.ст. 121, 122, 131, 132 ГПК РФ',
    ('ИСК', 'наниматель'): 'Руководствуясь ст.ст. 309, 330, 678, 682 ГК РФ, ст. 67, 155 ЖК РФ, ст.ст. 131, 132 ГПК РФ',
    ('ИСК', 'собственник'): 'Руководствуясь ст.ст. 309, 330, 540, 548 ГК РФ, ст.ст. 30, 31, 154, 155 ЖК РФ, ст. 131, 132 ГПК РФ',
}


def prishu_text(c, vz):
    shorts = ', '.join(fio_short(p['fio'], p['fem']) for p in c['persons'])
    dolg = sum_words(c['dolg'])
    posh = sum_words(c['poshlina'])
    period = f'за период с {fmt_date(c["d1"])} по {fmt_date(c["d2"])} года'
    parts = [f'задолженности за ЖКУ {period} в сумме {dolg}'] if c['typ'] == 'СП' else \
            [f'задолженность за ЖКУ {period} в сумме {dolg}']
    if c['peni'] > 0:
        pperiod = f'за период с {fmt_date(c["p1"])} по {fmt_date(c["p2"])} года'
        if c['typ'] == 'СП':
            parts.append(f'пени за нарушение сроков внесения платы за ЖКУ {pperiod} в сумме {sum_words(c["peni"])}')
        else:
            parts.append(f'пени за нарушение сроков внесения платы за ЖКУ {pperiod} в сумме {sum_words(c["peni"])}')
    parts.append(f'расходы по оплате государственной пошлины в сумме {posh}')
    tail = 'с перечислением указанных сумм на расчетный счет ' + vz + '.'
    if c['typ'] == 'СП':
        sol = ' солидарно' if len(c['persons']) > 1 else ''
        return (f'Выдать судебный приказ на имя {shorts} о взыскании{sol} ' +
                ', '.join(parts) + ', ' + tail)
    sol = '' if len(c['persons']) == 1 else ' солидарно с '
    with_ = sol + shorts if sol else ' с ' + shorts
    return f'Взыскать{with_} ' + ', '.join(parts) + ', ' + tail


def attachments(c, vz, cfg):
    trust = f'Копия доверенности № {cfg.get("доверенность_номер", "0001")} от {cfg.get("доверенность_дата", "09.01.2025")} года'
    if c['typ'] == 'СП':
        items = ['Справка формы № 9;', 'Платежное поручение;', 'Карточка лицевого счета.',
                 'История начислений по ЛС.', 'Отчет по поставщикам и услугам;',
                 f'Устав {vz}', 'Договоры с поставщиками услуг', 'Выписка из ЕГРЮЛ',
                 f'Протокол о продлении полномочий ген.директора {vz}',
                 'Копия реестра почтовых отправлений.', trust]
    else:
        items = ['Платежное поручение;', 'Справка формы № 9;', 'Карточка лицевого счета.',
                 'История начислений по лицевому счету.']
        if c['otmena']:
            items.append('Определение об отмене судебного приказа.')
        items += ['Копия свидетельств о госрегистрации.', f'Копия Устава {vz}.',
                  'Выписка из ЕГРЮЛ', 'Копии договоров', f'Копия доверенности представителя {vz}',
                  'Копия почтовых отправлений ответчикам']
    return items


def person_line(p):
    bits = [f'{fmt_date(p["dr"])} г.р.']
    if p['place'] and p['pasport']:
        bits.append(f'ур. {p["place"]}, паспорт: {p["pasport"]}')
    elif p['place']:
        bits.append(f'ур. {p["place"]}, паспортные данные неизвестны')
    elif p['pasport']:
        bits.append(f'паспорт: {p["pasport"]}')
    else:
        bits.append('место рождения и паспортные данные неизвестны')
    return f'{p["fio"]}, ' + ', '.join(bits)


# ---------------------------------------------------------------- word ----
def add_par(doc, text, align='justify', indent=True, bold=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(1)
    if indent:
        pf.first_line_indent = Cm(1.25)
    p.alignment = {'center': WD_ALIGN_PARAGRAPH.CENTER, 'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
                   'left': WD_ALIGN_PARAGRAPH.LEFT, 'right': WD_ALIGN_PARAGRAPH.RIGHT}[align]
    r = p.add_run(text)
    r.bold = bold
    return p


def build_docx(c, cfg, out_path):
    vz = cfg.get('взыскатель_наименование', 'ООО «Рога и Копыта»')
    doc = Document()
    st = doc.styles['Normal']
    st.font.name = 'Times New Roman'
    st.font.size = Pt(11)
    st.paragraph_format.space_after = Pt(0)
    st.paragraph_format.line_spacing = 1.0
    for s in doc.sections:
        s.left_margin, s.right_margin = Cm(3), Cm(1.5)
        s.top_margin, s.bottom_margin = Cm(2), Cm(2)

    # --- шапка (по правому краю) ---
    for line in c['sud']:
        add_par(doc, line, 'right', indent=False)
    add_par(doc, '------------------------------------------------', 'right', indent=False)
    head_word = 'Взыскатель' if c['typ'] == 'СП' else 'Истец'
    add_par(doc, f'{head_word}: {vz}', 'right', indent=False)
    for key, label in ((('юр_адрес'), 'Юр.адрес'), (('почт_адрес'), 'Почт.адрес')):
        if cfg.get(key):
            add_par(doc, f'{label}: {cfg[key]}', 'right', indent=False)
    if cfg.get('телефон'):
        add_par(doc, f'Тел. {cfg["телефон"]}', 'right', indent=False)
    if cfg.get('банк'):
        add_par(doc, cfg['банк'], 'right', indent=False)
        add_par(doc, f'р/с {cfg.get("р_с", "")} к/с {cfg.get("к_с", "")}', 'right', indent=False)
        add_par(doc, f'БИК {cfg.get("бик", "")} ИНН {cfg.get("инн", "")} КПП {cfg.get("кпп", "")}',
                'right', indent=False)

    label = 'Должник:' if c['typ'] == 'СП' else 'Ответчик(и):'
    lines = [f'{label} {person_line(c["persons"][0])}'] + \
            [person_line(p) for p in c['persons'][1:]]
    add_par(doc, '; '.join(lines), 'right', indent=False)
    who = 'Зарегистрированные по адресу' if len(c['persons']) > 1 else 'Зарегистрированная по адресу'
    add_par(doc, f'{who}: {c["addr"]}', 'right', indent=False)

    cena = c['dolg'] + c['peni']
    add_par(doc, f'Цена иска: {fmt_price(cena)}', 'right', indent=False)
    add_par(doc, f'Государственная пошлина: {fmt_price(c["poshlina"])}', 'right', indent=False)
    add_par(doc, '', 'center', indent=False)

    title = 'ЗАЯВЛЕНИЕ О ВЫДАЧЕ СУДЕБНОГО ПРИКАЗА' if c['typ'] == 'СП' else 'ИСКОВОЕ ЗАЯВЛЕНИЕ'
    add_par(doc, title, 'center', indent=False, bold=True)
    add_par(doc, '', 'center', indent=False)

    # --- мотивировка ---
    add_par(doc, intro_text(vz))
    add_par(doc, housing_text(c['typ'], c['cat'], c['housing'], len(c['persons'])))
    add_par(doc, f'Открыт отдельный лицевой счет № {c["ls"]}.')
    dolg_phrase = ('задолженность должника перед истцом составляет' if c['typ'] == 'СП'
                   else 'задолженность ответчика перед истцом составляет')
    add_par(doc, f'По состоянию на {fmt_date(c["sost"])} года, за период с {fmt_date(c["d1"])} '
                 f'по {fmt_date(c["d2"])} года, {dolg_phrase} {sum_words(c["dolg"])}.')

    if c['peni'] > 0:
        add_par(doc, PERENI[0])
        add_par(doc, PERENI[1])
        add_par(doc, f'За период с {fmt_date(c["p1"])} по {fmt_date(c["p2"])} года сумма пени за '
                     f'нарушение сроков внесения платы за ЖКУ составляет {sum_words(c["peni"])}.')

    for par in legal_block(c['cat']):
        add_par(doc, par)
    if c['cat'] == 'собственник' and len(c['persons']) > 1:
        for par in SOLID:
            add_par(doc, par)
    add_par(doc, ST309)
    add_par(doc, RUKOVODSTVO[(c['typ'], c['cat'])])
    add_par(doc, '', 'center', indent=False)
    add_par(doc, 'ПРОШУ:', 'center', indent=False, bold=True)
    add_par(doc, prishu_text(c, vz))
    add_par(doc, '', 'center', indent=False)

    add_par(doc, 'Приложения:', indent=False)
    for i, item in enumerate(attachments(c, vz, cfg), 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(1.0)
        p.paragraph_format.space_after = Pt(0)
        p.add_run(f'{i}. {item}')

    add_par(doc, '', 'center', indent=False)
    add_par(doc, f'Представитель {vz}', 'left', indent=False)
    add_par(doc, f'по доверенности № {cfg.get("доверенность_номер", "0001")} от '
                 f'{cfg.get("доверенность_дата", "09.01.2025")} года '
                 f'_________________ {cfg.get("представитель", "")}', 'left', indent=False)

    doc.save(out_path)


# ---------------------------------------------------------------- главная ----
def main():
    book = Path(sys.argv[1]) if len(sys.argv) > 1 else BASE / 'input-workbook.xlsx'
    if not book.exists():
        sys.exit(f'Книга не найдена: {book}\nСоздайте её: python3 create_workbook.py')

    cfg, claims, debtors, f9_texts = load_workbook_data(book)
    if not claims:
        sys.exit('Лист «Заявления» пуст.')

    missing = [k for k in ('взыскатель_наименование', 'доверенность_номер', 'доверенность_дата',
                           'представитель') if k not in cfg]
    if missing:
        print(f'ВНИМАНИЕ: на листе «Настройки» не заполнены: {", ".join(missing)}')

    errors, warnings = [], []
    warnings.append('Склонение фамилий в графе «ПРОШУ» выполняется автоматически — выборочно '
                    'просмотрите его в готовых документах.')
    docs, bad_claim = normalize_claims(claims, debtors, errors, warnings, f9_texts)
    docs = [c for c in docs if c['num'] not in bad_claim]
    for c in docs:   # текст справки Ф-9 в документ не вставляется (её прикладывают
        if not c.get('f9_texts'):   # вручную), но по ней определяется состав лиц —
            warnings.append(f'Заявление {c["num"]} (ЛС {c["ls"]}): справка формы 9 не привязана — '
                            'состав должников может быть неполным; проверьте импорт Ф-9.')

    stamp = datetime.now().strftime('%Y-%m-%d_%H-%M')
    out_dir = BASE / 'output' / stamp

    report = ['ОТЧЁТ-СВЕРКИ', f'Книга: {book.name}', f'Заявлений: {len(docs)}', '']
    if bad_claim:
        report.append(f'ОШИБКИ — заявлений не сформировано: {len(bad_claim)} (остальное собрано):')
        for nn in sorted(bad_claim, key=lambda x: (len(x), x)):
            report.append(f'  ✗ № {nn}:')
            report += [f'      {m}' for m in bad_claim[nn]]
        report.append('')
    if errors:
        report.append('ОШИБКИ (прочие):')
        report += [f'  ✗ {e}' for e in errors]
        report.append('')
    if warnings:
        report.append('ПРЕДУПРЕЖДЕНИЯ (проверьте вручную):')
        report += [f'  ⚠ {w}' for w in warnings]
        report.append('')

    rows = []
    out_dir.mkdir(parents=True, exist_ok=True)
    ok, failed = 0, 0
    for di, c in enumerate(docs, 1):
        print(f'[{di}/{len(docs)}] собираю документ № {c["num"]}…', flush=True)
        try:
            shorts = fio_nom_short(c['persons'][0]['fio'])
            fname = f'{c["num"]} {c["typ"]} {c["cat"]} {shorts} ЛС {c["ls"]}'
            if c.get('parts', 1) > 1:
                fname += f' ({c["part"]} из {c["parts"]})'
            fname += '.docx'
            fname = platform_utils.ascii_filename(fname)
            build_docx(c, cfg, out_dir / fname)
            print(f'[{di}/{len(docs)}] собрано: {fname}', flush=True)
            rows.append([c['num'], fname, c['typ'], c['cat'], c['ls'],
                         '; '.join(p['fio'] for p in c['persons']), c['addr'],
                         fmt_date(c['d1']), fmt_date(c['d2']), c['dolg'],
                         fmt_date(c['p1']), fmt_date(c['p2']), c['peni'],
                         c['dolg'] + c['peni'], c['poshlina'],
                         ' / '.join(map(str, c['sud']))])
            ok += 1
        except Exception as ex:  # noqa: BLE001
            failed += 1
            errors.append(f'Заявление {c["num"]}: не удалось сформировать документ — {ex}')
    if failed:
        report += ['ОШИБКИ при формировании:', *[f'  ✗ {e}' for e in errors[-failed:]]]

    # реестр
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Реестр'
    heads = ['№', 'Файл', 'Тип', 'Категория', 'Лицевой счёт', 'Должники', 'Адрес',
             'Долг: с', 'Долг: по', 'Долг, руб.', 'Пени: с', 'Пени: по', 'Пени, руб.',
             'Цена иска, руб.', 'Госпошлина, руб.', 'Суд']
    ws.append(heads)
    for r in rows:
        ws.append(r)
    sums_row = ['ИТОГО', '', '', '', '', '', '', '', '',
                sum(r[9] for r in rows), '', '', sum(r[12] for r in rows),
                sum(r[13] for r in rows), sum(r[14] for r in rows), '']
    ws.append(sums_row)
    for j, w in enumerate((5, 40, 6, 12, 12, 35, 35, 11, 11, 12, 11, 11, 12, 14, 14, 30), 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    for cell in ws[1]:
        cell.font = openpyxl.styles.Font(bold=True)
    wb.save(out_dir / 'registry.xlsx')
    report.append(f'Сформировано документов: {ok} в папку: {out_dir}')

    text = '\n'.join(report)
    report_path = (out_dir if out_dir.exists() else BASE) / 'validation-report.txt'
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(text, encoding='utf-8')
    print(text)


if __name__ == '__main__':
    main()
