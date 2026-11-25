#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Общие функции раздела «Досудебные претензии».

Шаблон претензии пользователя не переписывается: приложение само находит в нём
переменные места (получатель, адрес, номер и дата претензии, дата состояния,
период, сумма, срок оплаты) и подставляет значения построчно — по одной
претензии на должника, как в готовых файлах юриста.

Шапка (реквизиты взыскателя) читается из шаблона при загрузке списка и
кладётся в лист «Настройки» — её можно поправить в таблице вручную.
"""
import re
from copy import deepcopy
from datetime import date, datetime
from pathlib import Path

import docx
from docx.oxml.ns import qn
from docx.text.run import Run

BASE = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE / 'pretension-templates'
BOOK = 'pretension-input.xlsx'
REPORT = 'validation-report-pretension.txt'

MONTHS = ('января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
          'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря')

# «Фамилия И. О.» и перечисление таких («Фамилия И. О., Фамилия И. О.»)
FIO_RE = re.compile(
    r'[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.'
    r'(?:\s*,\s*[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.)*')
# адрес получателя: почтовый индекс и/или населённый пункт, улица, дом, квартира
ADDR_RE = re.compile(
    r'(?:\d{6},\s*)?[А-ЯЁ][а-яё\-]*[^;|\n]{0,90}?д\.\s*\d+[А-Яа-яА-Я]*'
    r'\s*кв\.?\s*\d*')
NUM_RE = re.compile(r'№\s*(\d+)\s*от\s*(\d{1,2}\.\d{1,2}\.\d{4})')
STATE_RE = re.compile(r'по состоянию на\s+(.{1,40}?)\s+года')
PERIOD_RE = re.compile(r'\bза\s+(.{3,80}?)\s+по состоянию на')
SUM_RE = re.compile(r'^\s*([\d\s]+[\d.,]+)\s*рублей')
TERM_RE = re.compile(r'в течение\s*(\d+)\s*дн')

# ───────────────────────────────────────────────────────────────── шаблон ──
def template_files():
    """Шаблоны претензии, которые лежат рядом с программой."""
    if not TEMPLATE_DIR.is_dir():
        return []
    return sorted(TEMPLATE_DIR.glob('*.docx'))


def pick_template(settings=None, addresses=()):
    """Выбрать шаблон претензии.

    Сначала — тот, что указан в настройках (юрист выбирает его раз и меняет
    при необходимости); если не указан — первый из папки шаблонов.
    """
    settings = settings or {}
    name = str(settings.get('шаблон') or '').strip()
    if name:
        p = TEMPLATE_DIR / name
        if p.is_file():
            return p
    files = template_files()
    return files[0] if files else None


def _all_runs(par):
    """Все run'ы абзаца по порядку — включая текст внутри гиперссылок."""
    runs = []
    for child in par._p:
        tag = child.tag.split('}')[-1]
        if tag == 'r':
            runs.append(Run(child, par))
        elif tag == 'hyperlink':
            for r in child.findall(qn('w:r')):
                runs.append(Run(r, par))
    return runs


def _run_map(par):
    """[(run, начало, конец)] — границы символов абзаца по run'ам."""
    out, pos = [], 0
    for r in _all_runs(par):
        out.append((r, pos, pos + len(r.text)))
        pos += len(r.text)
    return out


def _has_text(run):
    """Есть ли в run'е обычный текст (а не табуляция или поле)."""
    return run._r.find(qn('w:t')) is not None


def unlink_fields(par):
    """Расцепить поля Word (MERGEFIELD и пр.) — оставить только текст.

    В шаблонах юриста ФИО получателя бывает полем слияния: пока поле живо,
    Word может перерисовать его при открытии файла. Убираем служебные run'ы,
    текст остаётся.
    """
    for r in list(_all_runs(par)):
        el = r._r
        if (el.find(qn('w:fldChar')) is not None
                or el.find(qn('w:instrText')) is not None):
            el.getparent().remove(el)


def replace_span(par, start, end, new_text):
    """Заменить символы [start:end) абзаца, сохранив оформление первого run'а."""
    first = True
    for r, s, e in _run_map(par):
        if e <= start or s >= end:
            continue
        if not _has_text(r):
            continue          # табуляция или служебный элемент — не трогаем
        ls, le = max(start, s) - s, min(end, e) - s
        text = r.text
        if first:
            r.text = text[:ls] + new_text + text[le:]
            first = False
        else:
            r.text = text[:ls] + text[le:]


def _after_tabs(par):
    """(начало, конец) текста правее последней табуляции — «правая колонка»."""
    text = par.text
    pos = text.rfind('\t')
    if pos == -1:
        return None
    return pos + 1, len(text)


def _last_run_span(par):
    """(начало, конец) последнего непустого run'а абзаца."""
    for r, s, e in reversed(_run_map(par)):
        if e > s:
            return s, e
    return None


def _left_run_span(par):
    """(начало, конец) первого непустого run'а абзаца — левая колонка шапки."""
    for r, s, e in _run_map(par):
        if e > s:
            return s, e
    return None


def analyze(doc):
    """Найти в шаблоне переменные места.

    Возвращает {поле: (абзац, начало, конец)} — границы в тексте абзаца.
    """
    slots = {}
    paras = doc.paragraphs
    for i, par in enumerate(paras):
        text = par.text
        if not text.strip():
            continue
        low = text.lower()

        if 'fio' not in slots:
            span = _after_tabs(par)
            m = FIO_RE.search(text[span[0]:span[1]]) if span else None
            if not m and not low.startswith(('исполнитель', 'генеральный')):
                m = FIO_RE.search(text)
            if m:
                s = (span[0] + m.start()) if span else m.start()
                e = (span[0] + m.end()) if span else m.end()
                if text[e:e + 1] == '.' and text[e + 1:e + 2] in ('', ' ', '\t'):
                    e += 1          # «Акбаева Ж. Ш. » + «. » — точка не нужна
                slots['fio'] = (i, s, e)

        if 'addr' not in slots:
            span = _last_run_span(par)
            if span:
                m = ADDR_RE.search(text[span[0]:span[1]])
                if m:
                    slots['addr'] = (i, span[0] + m.start(), span[0] + m.end())

        if 'number' not in slots:
            m = NUM_RE.search(text)
            if m:
                slots['number'] = (i, m.start(1), m.end(1))
                slots['date'] = (i, m.start(2), m.end(2))

        if 'state_date' not in slots:
            m = STATE_RE.search(text)
            if m:
                slots['state_date'] = (i, m.start(1), m.end(1))

        if 'period' not in slots:
            m = PERIOD_RE.search(text)
            if m:
                slots['period'] = (i, m.start(1), m.end(1))

        if 'sum' not in slots:
            m = SUM_RE.search(text)
            if m:
                slots['sum'] = (i, m.start(1), m.end(1))

        if 'term' not in slots:
            m = TERM_RE.search(text)
            if m:
                slots['term'] = (i, m.start(1), m.end(1))
    return slots


def describe_slots(slots):
    """Человекочитаемое описание найденных мест — показываем юристу."""
    names = {'fio': 'получатель (ФИО)', 'addr': 'адрес получателя',
             'number': 'номер претензии', 'date': 'дата претензии',
             'state_date': 'дата состояния («по состоянию на …»)',
             'period': 'период («за … »)', 'sum': 'сумма долга',
             'term': 'срок оплаты (дней)'}
    return [(names.get(k, k), v[0] + 1) for k, v in slots.items()]


# ───────────────────────────────────────────────────────────── значения ──
def money(value):
    """12345.6 → «12 345,60»."""
    try:
        num = float(str(value).replace(' ', '').replace(',', '.'))
    except (TypeError, ValueError):
        return str(value or '').strip()
    return '{:,.2f}'.format(num).replace(',', ' ').replace('.', ',')


def date_words(value):
    """Дата словами: «01 ноября 2026»."""
    d = _as_date(value)
    if not d:
        return str(value or '').strip()
    return '%02d %s %d' % (d.day, MONTHS[d.month - 1], d.year)


def date_plain(value):
    """«25.11.2026»."""
    d = _as_date(value)
    return d.strftime('%d.%m.%Y') if d else str(value or '').strip()


def _as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value or '').strip()
    if not s:
        return None
    for fmt in ('%d.%m.%Y', '%d.%m.%y', '%Y-%m-%d', '%d %B %Y', '%d %b %Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    m = re.search(r'(\d{1,2})\.(\d{1,2})\.(\d{2,4})', s)
    if m:
        d, mo, y = (int(x) for x in m.groups())
        if y < 100:
            y += 2000
        try:
            return date(y, mo, d)
        except ValueError:
            return None
    return None



# ──────────────────────────────────────────────────────── сборка книги ──
DEFAULT_SETTINGS = [
    ('шаблон', ''),
    ('взыскатель', ''),
    ('наименование_1', ''),
    ('наименование_2', ''),
    ('адрес_рассрочки', ''),
    ('адрес_регион', ''),
    ('телефон_рассрочки', ''),
    ('юр_адрес_1', ''),
    ('юр_адрес_2', ''),
    ('почт_адрес_1', ''),
    ('почт_адрес_2', ''),
    ('телефон', ''),
    ('телефон_2', ''),
    ('факс', ''),
    ('email', ''),
    ('р_с', ''),
    ('к_с', ''),
    ('БИК', ''),
    ('ИНН', ''),
    ('КПП', ''),
    ('директор', ''),
    ('исполнитель', ''),
    ('претензия_номер_от', '1'),
    ('дата_претензии', ''),
    ('срок_оплаты, дней', '30'),
    ('дата_состояния', ''),
    ('период', ''),
    ('точные повторы пропускать', 'да'),
]

CLAIM_HEADERS = ['№', 'Лицевой счёт', 'ФИО', 'Адрес', 'Сумма', 'Дата состояния',
                 'Период', '№ претензии', 'Дата претензии', 'Примечание']


def sheet_settings(ws):
    """Настройки из листа «Настройки» → словарь."""
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        key = str(row[0]).strip().lower().replace(' ', '_')
        val = row[1] if len(row) > 1 else ''
        out[key] = '' if val is None else val
    return out


def sheet_claims(ws):
    """Строки листа «Претензии» → список словарей."""
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or all(v is None or str(v).strip() == '' for v in row):
            continue
        vals = list(row) + [None] * (10 - len(row))
        out.append({
            '№': vals[0], 'лс': vals[1], 'фио': vals[2], 'адрес': vals[3],
            'сумма': vals[4], 'дата_состояния': vals[5], 'период': vals[6],
            'номер': vals[7], 'дата_претензии': vals[8], 'примечание': vals[9],
        })
    return out


def read_book(path):
    """Читает книгу раздела: (настройки, претензии)."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    settings = sheet_settings(wb['Настройки']) if 'Настройки' in wb.sheetnames else {}
    claims = sheet_claims(wb['Претензии']) if 'Претензии' in wb.sheetnames else []
    return settings, claims


def header_values(settings):
    """Что подставлять в шапку (пустое — не трогаем шаблон)."""
    return {k: str(v).strip() for k, v in settings.items() if str(v or '').strip()}


# ─────────────────────────────────────────────────────────────── шапка ──
DASHES = '-----'


def _span_after(text, label, stop=None):
    """(начало, конец) текста после метки — до запятой/табуляции или конца."""
    pos = text.find(label)
    if pos == -1:
        return None
    start = pos + len(label)
    end = len(text)
    if stop:
        for s in stop:
            p = text.find(s, start)
            if p != -1:
                end = min(end, p)
    return start, end


def extract_header(doc):
    """Прочитать реквизиты взыскателя из шаблона — чтобы показать их в таблице."""
    out = {}
    paras = doc.paragraphs
    for i, p in enumerate(paras):
        t = p.text
        if 'Адрес юридический:' in t:
            span = _span_after(t, 'Адрес юридический:')
            out['юр_адрес_1'] = t[span[0]:span[1]].strip()
            if i + 1 < len(paras):
                out['юр_адрес_2'] = _left_text(paras[i + 1])
        elif 'Адрес почтовый:' in t:
            span = _span_after(t, 'Адрес почтовый:', stop=(DASHES,))
            out['почт_адрес_1'] = t[span[0]:span[1]].strip()
            if i + 1 < len(paras):
                out['почт_адрес_2'] = _left_text(paras[i + 1])
                out['адрес_регион'] = _right_text(paras[i + 1])
        elif t.strip().startswith('Тел.'):
            span = _span_after(t, 'Тел.', stop=(',', '\t'))
            out['телефон'] = t[span[0]:span[1]].strip()
        elif re.match(r'^\s*8\(', t) and 'телефон_2' not in out:
            out['телефон_2'] = t.split(',')[0].strip()
        elif t.strip().startswith('факс'):
            span = _span_after(t, 'факс', stop=('\t',))
            out['факс'] = t[span[0]:span[1]].strip()
        elif 'Электронный адрес' in t:
            span = _span_after(t, 'Электронный адрес', stop=('\t',))
            out['email'] = t[span[0]:span[1]].strip()
        elif t.strip().startswith('Р/с'):
            out['р_с'] = t.split('Р/с', 1)[1].strip()
        elif t.strip().startswith('К/с'):
            out['к_с'] = t.split('К/с', 1)[1].strip()
        elif t.strip().startswith('БИК'):
            rest = t.split('БИК', 1)[1].strip()
            parts = rest.split('ИНН')
            out['БИК'] = parts[0].strip()
            if len(parts) > 1:
                out['ИНН'] = parts[1].strip()
        elif t.strip().startswith('КПП'):
            out['КПП'] = t.split('КПП', 1)[1].strip()
        elif t.strip().startswith('Генеральный директор'):
            # после подчёркиваний — фамилия подписанта
            parts = [x for x in re.split(r'_+', t.split('директор', 1)[1]) if x.strip()]
            out['директор'] = parts[-1].strip() if parts else ''
        elif t.strip().startswith('Исполнитель:'):
            out['исполнитель'] = t.split('Исполнитель:', 1)[1].strip()
        elif 'по адресу:' in t:
            rest = t.split('по адресу:', 1)[1].strip()
            m = re.search(r'\.\s+(?=[А-ЯЁ][а-яё]+\s)', rest)
            out['адрес_рассрочки'] = (rest[:m.start()] if m else rest).strip()
        elif 'по телефону:' in t:
            rest = t.split('по телефону:', 1)[1]
            m = re.search(r'\s+для\s', rest)
            out['телефон_рассрочки'] = (rest[:m.start()] if m else rest).strip()
        else:
            m = re.search(r'перед\s+(.+?)\s+за\s', t)
            if m and 'взыскатель' not in out:
                out['взыскатель'] = m.group(1).strip()
    # название организации: строка с формой и строка с именем
    for i, par in enumerate(paras):
        txt = par.text.strip()
        if re.match(r'(ОБЩЕСТВО|АКЦИОНЕРНОЕ|ИНДИВИДУАЛЬНЫЙ)', txt):
            out['наименование_1'] = txt
            if i + 1 < len(paras) and paras[i + 1].text.strip():
                out['наименование_2'] = paras[i + 1].text.strip()
            break
    # значения, которые есть в шаблоне, но в настройках их не было бы
    slots = analyze(doc)
    if 'period' in slots:
        pi, s, e = slots['period']
        out.setdefault('период', paras[pi].text[s:e].strip())
    if 'term' in slots:
        pi, s, e = slots['term']
        out.setdefault('срок_оплаты, дней', paras[pi].text[s:e].strip())
    return out


def apply_header(paras, head, tpl=None):
    """Подставить реквизиты взыскателя в шапку документа.

    tpl — то, что лежит в шаблоне: по нему ищем старые значения и заменяем
    на новые (например, название организации во всех местах текста).
    """
    if not head:
        return
    head = {str(k).strip().lower(): v for k, v in head.items()}
    tpl = {str(k).strip().lower(): v for k, v in (tpl or {}).items()}
    for key in ('взыскатель', 'наименование_1', 'наименование_2',
                'адрес_рассрочки', 'телефон_рассрочки'):
        old = str(tpl.get(key) or '').strip()
        new = str(head.get(key) or '').strip()
        if not old or not new or old == new:
            continue
        for par in paras:
            while True:
                i = par.text.find(old)
                if i == -1:
                    break
                replace_span(par, i, i + len(old), new)
    for i, p in enumerate(paras):
        t = p.text
        if 'Адрес юридический:' in t and head.get('юр_адрес_1'):
            _replace_after_label(p, 'Адрес юридический:', head['юр_адрес_1'])
            if i + 1 < len(paras) and head.get('юр_адрес_2'):
                _replace_left(paras[i + 1], head['юр_адрес_2'])
        elif 'Адрес почтовый:' in t and head.get('почт_адрес_1'):
            _replace_after_label(p, 'Адрес почтовый:', head['почт_адрес_1'],
                                 stop=(DASHES,))
            if i + 1 < len(paras) and head.get('почт_адрес_2'):
                _replace_left(paras[i + 1], head['почт_адрес_2'])
            if i + 1 < len(paras) and head.get('адрес_регион'):
                _replace_right(paras[i + 1], head['адрес_регион'])
        elif t.strip().startswith('Тел.') and head.get('телефон'):
            _replace_after_label(p, 'Тел.', head['телефон'], stop=(',', '\t'))
        elif re.match(r'^\s*8\(', t) and head.get('телефон_2'):
            stop = t.find(',')
            if stop == -1:
                stop = len(t)
            replace_span(p, 0, stop, head['телефон_2'])
        elif t.strip().startswith('факс') and head.get('факс'):
            _replace_after_label(p, 'факс', head['факс'], stop=('\t',))
        elif 'Электронный адрес' in t and head.get('email'):
            _replace_after_label(p, 'Электронный адрес', head['email'], stop=('\t',))
        elif t.strip().startswith('Р/с') and head.get('р_с'):
            _replace_after_label(p, 'Р/с', head['р_с'])
        elif t.strip().startswith('К/с') and head.get('к_с'):
            _replace_after_label(p, 'К/с', head['к_с'])
        elif t.strip().startswith('БИК') and head.get('бик'):
            _replace_after_label(p, 'БИК', head['бик'], stop=(' ИНН',))
            if head.get('инн'):
                _replace_after_label(p, 'ИНН', head['инн'])
        elif t.strip().startswith('КПП') and head.get('кпп'):
            _replace_after_label(p, 'КПП', head['кпп'])
        elif t.strip().startswith('Генеральный директор') and head.get('директор'):
            pos = t.find('_')
            if pos == -1:
                pos = t.lower().find('директор') + len('директор')
            replace_span(p, pos, len(t), head['директор'])
        elif t.strip().startswith('Исполнитель:') and head.get('исполнитель'):
            _replace_after_label(p, 'Исполнитель:', head['исполнитель'])
        elif head.get('взыскатель'):
            m = re.search(r'перед\s+(.+?)\s+за\s', t)
            if m:
                replace_span(p, m.start(1), m.end(1), head['взыскатель'])


def _replace_after_label(par, label, value, stop=None):
    """Заменить значение после метки, сохранив пробелы вокруг."""
    span = _span_after(par.text, label, stop)
    if not span:
        return False
    orig = par.text[span[0]:span[1]]
    lead = orig[:len(orig) - len(orig.lstrip())]
    replace_span(par, span[0], span[1], lead + value)
    return True


def _right_span(par):
    """(начало, конец) правой колонки абзаца — после табуляции или пробелов."""
    text = par.text
    best = -1
    pos = text.rfind('\t')
    if pos != -1:
        best = pos
    for m in re.finditer(r' {3,}', text):
        if m.start() > best:
            best = m.end()
    if best == -1:
        return None
    return best, len(text)


def _right_text(par):
    span = _right_span(par)
    return par.text[span[0]:span[1]].strip() if span else ''


def _replace_right(par, value):
    span = _right_span(par)
    if span:
        replace_span(par, span[0], span[1], value)


def _left_text(par):
    """Левая колонка абзаца — первый run (в шапке там адрес взыскателя)."""
    runs = _all_runs(par)
    return runs[0].text.strip() if runs else par.text.strip()


def _replace_left(par, value):
    """Заменить левую колонку абзаца (первый run), не потеряв отступы."""
    span = _left_run_span(par)
    if not span:
        return
    s, e = span
    core = par.text[s:e].strip()
    if core:
        replace_span(par, s, s + len(core), value)


# ───────────────────────────────────────────────────── сборка документа ──
def build_document(template_path, claims, settings, out_path, report=None):
    """Собрать документ: одна претензия на строку списка, каждая с новой страницы."""
    doc = docx.Document(str(template_path))
    slots = analyze(doc)
    tpl = {str(k).strip().lower(): v for k, v in extract_header(doc).items()}
    missing = [k for k in ('fio', 'addr', 'sum') if k not in slots]
    if missing:
        raise RuntimeError(
            'в шаблоне не найдены: ' + ', '.join(missing) +
            ' — проверьте шаблон претензии')

    body = doc.element.body
    sect = None
    for child in list(body):
        if child.tag.endswith('}sectPr'):
            sect = child
        body.remove(child)
    src = docx.Document(str(template_path))
    template_block = [c for c in src.element.body
                      if not c.tag.endswith('}sectPr')]

    settings = {str(k).strip().lower(): v for k, v in settings.items()}
    head = header_values(settings)
    default_state = settings.get('дата_состояния') or ''
    default_period = settings.get('период') or ''
    default_term = settings.get('срок_оплаты, дней') or '30'
    default_date = settings.get('дата_претензии') or date.today().strftime('%d.%m.%Y')
    try:
        first_number = int(str(settings.get('претензия_номер_от') or 1).strip())
    except ValueError:
        first_number = 1

    for n, claim in enumerate(claims):
        before = len(doc.paragraphs)
        for el in template_block:
            body.append(deepcopy(el))
        paras = doc.paragraphs[before:]
        for par in paras:          # поля Word (MERGEFIELD) — расцепить
            unlink_fields(par)

        number = claim.get('номер')
        number = int(str(number).strip()) if str(number or '').strip().isdigit() \
            else first_number + n
        claim_date = claim.get('дата_претензии') or default_date
        state_date = claim.get('дата_состояния') or default_state \
            or date.today().strftime('%d.%m.%Y')
        period = claim.get('период') or default_period
        term = default_term
        fio = str(claim.get('фио') or '').strip()
        addr = str(claim.get('адрес') or '').strip()
        amount = money(claim.get('сумма'))

        values = {
            'fio': fio,
            'addr': addr,
            'number': str(number),
            'date': date_plain(claim_date),
            'state_date': date_words(state_date),
            'period': period,
            'sum': amount,
            'term': str(term),
        }
        # замены внутри одного абзаца — с конца к началу, чтобы смещения не «поехали»
        by_para = {}
        for key, (pi, s, e) in slots.items():
            by_para.setdefault(pi, []).append((s, e, values.get(key, '')))
        for pi in sorted(by_para, reverse=True):
            if pi >= len(paras):
                continue
            for s, e, val in sorted(by_para[pi], reverse=True):
                replace_span(paras[pi], s, e, val)
        if n and paras:
            paras[0].paragraph_format.page_break_before = True
        apply_header(paras, head, tpl)

    if sect is not None:
        body.append(sect)
    doc.save(str(out_path))
    return slots
