#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Общие функции раздела «Исполнительное производство» (ФССП).

Задача раздела: из судебных приказов (сканы) и реестра исполнительных
документов собрать заявления об исполнении и сопроводительное письмо.

Три источника данных:
  1. «Судебные приказы» — PDF со сканами (OCR). Из приказа берём: номер
     дела, дату, ФИО должников, адрес, сумму долга, госпошлину, суд.
  2. Реестр исполнительных документов в ФССП — таблица Excel: одна строка
     на должника (ФИО, адрес, № исполнительного документа, суд, суммы).
  3. Шаблоны юриста (.docx) — заявление об исполнении судебного приказа,
     заявление об исполнении исполнительного листа, сопровод приставу.

Шаблоны юриста НЕ переписываются: приложение само находит в них переменные
места по опорным словам («в сумме …», «выданному: …», «зарегистрирован(а)
по адресу: …») и подставляет значения построчно.
"""
import re
from copy import deepcopy
from datetime import date, datetime
from pathlib import Path

import docx
from docx.oxml.ns import qn
from docx.text.run import Run

BASE = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE / 'iss-templates'
BOOK = 'iss-input.xlsx'
REPORT = 'validation-report-iss.txt'
OUT_DIR = BASE / 'output-iss'
OCR_DIR = BASE / 'output' / 'ocr-iss'

MONTHS = ('января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
          'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря')

# ─────────────────────────────────────────────────── разбор судебного приказа ──
CASE_NO_RE = re.compile(r'Дело\s*№\s*([\d\-–/]+)', re.IGNORECASE)
DATE_RE = re.compile(r'(\d{1,2})\s+([а-яё]+)\s+(\d{4})\s*года', re.IGNORECASE)
GOSPOSHLINA_RE = re.compile(
    r'госпошлины\s+в\s+размере\s+(?P<rub>[\d\s]+)\s*руб\.'
    r'(?:\s*(?P<kop>[\d\s]+)\s*коп\.)?', re.IGNORECASE)
DEBT_RE = re.compile(
    r'задолженност[^\n]{0,140}?в\s+размере\s+(?P<rub>[\d\s]+)\s*руб\.'
    r'(?:\s*(?P<kop>[\d\s]+)\s*коп\.)?', re.IGNORECASE)
COURT_RE = re.compile(
    r'Судебн(?:ый|ого|ым|ом)\s+участок[а]?\s*№\s*(\d+)', re.IGNORECASE)
FIO_RE = re.compile(
    r'([А-ЯЁ][а-яё\-]+(?:\s+[А-ЯЁ][а-яё\-]+){1,2})'
    r'(?:\s*,\s*\d{1,2}\.\d{1,2}\.\d{4}\s*года\s+рождения)?')
# ФИО вместе с датой рождения — самый надёжный признак в приказе
FIO_DOB_RE = re.compile(
    r'([А-ЯЁ][а-яё\-]+(?:\s+[А-ЯЁ][а-яё\-]+){1,2})\s*,'
    r'\s*\d{1,2}\.\d{1,2}\.\d{4}\s*года\s+рождения')
ADDR_RE = re.compile(
    r'зарегистрированн(?:ых|ой|ого|ом)\s+по\s+адресу:\s*(?P<addr>[^\n]+)',
    re.IGNORECASE)
ADDR_CUT_RE = re.compile(
    r'\s*,?\s*(?:исследовав|вынесен|и\s+приложенных|руководствуясь|'
    r'Должник\s+вправе|Указанные\s+суммы|паспорт)', re.IGNORECASE)


def _num(value):
    """'5913' / '5 913' → 5913.0; None → None."""
    if value is None:
        return None
    s = re.sub(r'[\s ]', '', str(value))
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def money(value):
    """12345.6 → «12 345,60»."""
    try:
        num = float(str(value).replace(' ', '').replace(',', '.'))
    except (TypeError, ValueError):
        return str(value or '').strip()
    return '{:,.2f}'.format(num).replace(',', ' ').replace('.', ',')


def money_words(value):
    """Сумма в виде шаблона юриста: «51885 руб. 77 коп.»."""
    try:
        num = float(str(value).replace(' ', '').replace(',', '.'))
    except (TypeError, ValueError):
        return str(value or '').strip()
    rub = int(num)
    kop = int(round((num - rub) * 100))
    if kop == 100:
        rub += 1
        kop = 0
    return '%d руб. %02d коп.' % (rub, kop)


def _as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value or '').strip()
    if not s:
        return None
    for fmt in ('%d.%m.%Y', '%d.%m.%y', '%Y-%m-%d'):
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


def date_plain(value):
    d = _as_date(value)
    return d.strftime('%d.%m.%Y') if d else str(value or '').strip()


def date_words(value):
    """«06 октября 2026 года»."""
    d = _as_date(value)
    if not d:
        return str(value or '').strip()
    return '%02d %s %d года' % (d.day, MONTHS[d.month - 1], d.year)


def parse_case_date(text):
    """Дата приказа: «06 октября 2026 года» → date."""
    m = DATE_RE.search(text)
    if not m:
        return None
    day, mon, year = m.groups()
    low = mon.lower().strip()
    for i, name in enumerate(MONTHS, 1):
        if name.startswith(low[:3]) or low.startswith(name[:3]):
            try:
                return date(int(year), i, int(day))
            except ValueError:
                return None
    return None


# ──────────────────────── ФИО: родительный падеж (из приказа) → именительный ──
FEMALE_PATRONYMIC_RE = re.compile(
    r'(вна|евна|ична|инична|вны|евны|ичны|иничны)$', re.IGNORECASE)
MALE_PATRONYMIC_RE = re.compile(
    r'(ович|евич|ич|ог|ов|ев|ин|ын|ик|овича|евича|ога|ова|ева|ина)$',
    re.IGNORECASE)


def fio_to_nominative(fio):
    """«Фамилиной Имени Отчеством» → «Фамилия Имя Отчество».

    В судебном приказе ФИО стоит в родительном падеже («с должников — …»),
    а в заявлении нужно именительный. Пол определяется по отчеству —
    и по именительному, и по родительному окончанию.
    Если пол не определяется — фамилия и имя оставляем как есть.
    """
    parts = str(fio or '').split()
    if len(parts) < 2:
        return fio
    surname, name = parts[0], parts[1]
    patronymic = parts[2] if len(parts) > 2 else ''

    female = bool(patronymic and FEMALE_PATRONYMIC_RE.search(patronymic))
    male = bool(patronymic and MALE_PATRONYMIC_RE.search(patronymic))

    if female:
        low = surname.lower()
        if low.endswith('ой') and len(surname) > 3:
            surname = surname[:-2] + 'а'
        elif low.endswith('ей') and len(surname) > 3:
            surname = surname[:-2] + 'я'
        if name.lower().endswith('ы'):
            name = name[:-1] + 'а'
        if patronymic.lower().endswith('вны'):
            patronymic = patronymic[:-1] + 'а'
        elif patronymic.lower().endswith('евны'):
            patronymic = patronymic[:-1] + 'а'
        elif patronymic.lower().endswith('ичны'):
            patronymic = patronymic[:-1] + 'а'
        elif patronymic.lower().endswith('иничны'):
            patronymic = patronymic[:-1] + 'а'
    elif male:
        if name.lower().endswith('а') and len(name) > 2:
            name = name[:-1]
        if patronymic.lower().endswith('а') and len(patronymic) > 3:
            patronymic = patronymic[:-1]

    return ' '.join([surname, name] + ([patronymic] if patronymic else []))


def _norm_text(text):
    return re.sub(r'\s+', ' ', text or '')


def parse_order(text):
    """Разобрать текст одного судебного приказа.

    Возвращает словарь полей (пустые — если не нашлось).
    """
    out = {'номер_дела': '', 'дата': '', 'суд': '', 'номер_участка': '',
           'фио': '', 'адрес': '', 'сумма': '', 'госпошлина': '',
           'период': '', 'должников': []}
    flat = _norm_text(text)

    m = CASE_NO_RE.search(text)
    if m:
        out['номер_дела'] = m.group(1).strip()
    d = parse_case_date(text)
    if d:
        out['дата'] = d

    # суд — из шапки приказа: «Судебный участок № 75 Тосненского района …»
    m = COURT_RE.search(text)
    if m:
        out['номер_участка'] = m.group(1)
        head = ''
        for line in text.splitlines():
            if 'участок' in line.lower() and '№' in line:
                head = re.sub(r'\s+', ' ', line).strip()
                break
        head = re.sub(r'^\d{6}\b.*$', '', head).strip()
        out['суд'] = head or ('Судебный участок № %s' % m.group(1))

    m = GOSPOSHLINA_RE.search(flat)
    if m:
        rub = _num(m.group('rub'))
        kop = _num(m.group('kop')) or 0
        if rub is not None:
            out['госпошлина'] = rub + kop / 100

    m = DEBT_RE.search(flat)
    if m:
        rub = _num(m.group('rub'))
        kop = _num(m.group('kop')) or 0
        if rub is not None:
            out['сумма'] = rub + kop / 100

    m = re.search(r'за\s+период\s+с\s+([\d.]+)\s+по\s+([\d.]+)', flat)
    if m:
        out['период'] = 'с %s по %s' % (m.group(1), m.group(2))

    m = ADDR_RE.search(flat)
    if m:
        addr = ADDR_CUT_RE.split(m.group('addr'))[0]
        out['адрес'] = addr.strip(' ,.')

    # должники — между «с должник(ов)» и «зарегистрированн»
    block = ''
    mb = re.search(r'с\s+должник\w*\s*[—:-]?\s*(?P<body>.+?)'
                   r'(?=зарегистрированн)', flat, re.IGNORECASE)
    if mb:
        block = mb.group('body')
    names = []
    # надёжный признак — ФИО, за которым идёт дата рождения
    for m in FIO_DOB_RE.finditer(block):
        fio = ' '.join(m.group(1).split()).strip(',')
        if fio and fio not in names:
            names.append(fio)
    if not names:
        # запасной вариант: ФИО в начале каждого фрагмента списка
        for piece in re.split(r',(?=\s*[А-ЯЁ][а-яё]+\s+[А-ЯЁ])', block):
            piece = piece.strip()
            m = FIO_RE.match(piece)
            if not m:
                continue
            fio = ' '.join(m.group(1).split()).strip(',')
            # обрывок вида «Город Область» — это не ФИО
            if not fio or fio in names or len(fio.split()) < 2:
                continue
            if fio.lower().startswith(('урожен', 'паспорт')):
                continue
            names.append(fio)
    out['должников'] = names
    if names:
        out['фио'] = ', '.join(names)
    return out


def ocr_pdf(path, cache_dir=None):
    """PDF со сканами → текст каждой страницы.

    Страницы кэшируются: повторный импорт не гоняет OCR заново.
    """
    import platform_utils
    cache_dir = Path(cache_dir or OCR_DIR)
    cache_dir.mkdir(parents=True, exist_ok=True)
    path = Path(path)
    stem = platform_utils.ascii_filename(path.stem)
    cache = cache_dir / (stem + '.txt')
    if cache.exists() and cache.stat().st_size > 0:
        raw = cache.read_text(encoding='utf-8', errors='replace')
    else:
        import subprocess
        img_dir = cache_dir / ('img-' + stem)
        img_dir.mkdir(parents=True, exist_ok=True)
        subprocess.run(['pdftoppm', '-r', '300', '-gray', '-png',
                        str(path), str(img_dir / 'p')], check=True)
        pages = []
        for img in sorted(img_dir.glob('p-*.png')):
            out = cache_dir / platform_utils.ascii_filename(img.stem)
            subprocess.run(['tesseract', str(img), str(out), '-l', 'rus',
                            '--psm', '6'], check=True, capture_output=True)
            pages.append((out.with_suffix('.txt')).read_text(
                encoding='utf-8', errors='replace'))
            try:
                img.unlink()
            except OSError:
                pass
        raw = '\n\n=== страница ===\n\n'.join(pages)
        cache.write_text(raw, encoding='utf-8')
    return split_pages(raw)


def split_pages(raw):
    """Текст с разделителями страниц → список страниц."""
    parts = re.split(r'\n\s*===?\s*страница\s*===?\s*\n', raw)
    return [p for p in parts if p.strip()]


# ───────────────────────────────────────────────────── работа с шаблонами ──
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
    return run._r.find(qn('w:t')) is not None


def unlink_fields(par):
    """Расцепить поля Word (MERGEFIELD) — оставить только текст.

    Важно: сами run'ы остаются (в них может быть оформление места для
    значения), удаляются только служебные элементы поля. Иначе смещения
    в абзаце «поедут» и значения попадут не туда.
    """
    for r in list(_all_runs(par)):
        el = r._r
        for tag in ('w:fldChar', 'w:instrText'):
            node = el.find(qn(tag))
            if node is not None:
                el.remove(node)


# ──────────────────────────────── карта абзаца: текст и пустые run'ы ──
# В шаблонах юриста места для значений — пустые run'ы Word (без текста).
# Поэтому работаем не по символам, а по run'ам: пустой run'ы = «место
# для значения», обычный текст = уже заполненный образец.
def par_map(par):
    """(синтетический текст, позиции).

    Позиция: (run, смещение в run'е, признак «пустое место»).
    Пустой run даёт в синтетическом тексте один символ-заглушку, чтобы
    границы мест не «съезжали».
    """
    text = []
    positions = []
    for r in _all_runs(par):
        if _has_text(r):
            for j, _ch in enumerate(r.text):
                text.append(r.text[j])
                positions.append((r, j, False))
        else:
            text.append('\x00')
            positions.append((r, 0, True))
    return ''.join(text), positions


def _covered(selection):
    """Сгруппировать позиции по run'ам: [(run, от, до)]."""
    out = []
    for run, off, _empty in selection:
        if out and out[-1][0] is run:
            out[-1] = (run, out[-1][1], off + 1)
        else:
            out.append((run, off, off + 1))
    return out


def _cut(run, lo, hi):
    """Убрать из run'а символы [lo:hi)."""
    text = run.text
    if lo >= len(text):
        return
    run.text = text[:lo] + text[min(hi, len(text)):]


def apply_value(par, start, end, value):
    """Поставить значение в место [start:end) абзаца (границы в par_map).

    Если в месте есть текст (заполненный образец) — заменяем его,
    сохраняя оформление. Если места нет (пустой шаблон) — пишем в
    пустые run'ы: оформление места сохраняется.
    """
    _text, positions = par_map(par)
    selection = positions[start:end]
    if not selection:
        return False
    text_runs = [p for p in selection if not p[2]]
    if text_runs:
        first = True
        for run, lo, hi in _covered(text_runs):
            text = run.text
            rest = text[hi:]
            tail = '' if not rest.strip() else rest
            if first:
                run.text = text[:lo] + value + tail
                first = False
            else:
                run.text = text[:lo] + tail
        return True
    placeholders = [p for p in selection if p[2]]
    if placeholders:
        run, off, _ = placeholders[0]
        run.text = run.text[:off] + value + run.text[off:]
        return True
    return False


def replace_span(par, start, end, new_text):
    """Заменить символы [start:end) абзаца, сохранив оформление первого run'а."""
    first = True
    for r, s, e in _run_map(par):
        if e <= start or s >= end:
            continue
        if not _has_text(r):
            continue
        ls, le = max(start, s) - s, min(end, e) - s
        text = r.text
        if first:
            r.text = text[:ls] + new_text + text[le:]
            first = False
        else:
            r.text = text[:ls] + text[le:]


def insert_at(par, pos, text):
    """Вставить текст в позицию pos текста абзаца."""
    for r, s, e in _run_map(par):
        if s <= pos <= e and _has_text(r):
            r.text = r.text[:pos - s] + text + r.text[pos - s:]
            return True
    par.add_run(text)
    return False


def set_span(par, start, end, text):
    """Поставить текст в диапазон [start:end); если диапазон пуст — вставить."""
    if end > start:
        replace_span(par, start, end, text)
    else:
        insert_at(par, start, text)


def template_files():
    """Шаблоны раздела, которые лежат рядом с программой."""
    if not TEMPLATE_DIR.is_dir():
        return []
    return sorted(TEMPLATE_DIR.glob('*.docx'))


def pick_template(kind, settings=None):
    """Выбрать шаблон по виду: 'сп' / 'ил' / 'сопровод'.

    Сначала — тот, что указан в настройках; иначе первый подходящий
    по имени файла.
    """
    settings = settings or {}
    name = str(settings.get('шаблон_' + kind) or '').strip()
    if name:
        p = TEMPLATE_DIR / name
        if p.is_file():
            return p
    files = template_files()
    for f in files:
        low = f.name.lower()
        if kind == 'сп' and 'приказ' in low:
            return f
        if kind == 'ил' and 'лист' in low:
            return f
        if kind == 'сопровод' and ('сопровод' in low or 'сопр' in low):
            return f
    return files[0] if files else None


# ───────────────────────────────────────── разбор шаблона заявления ──
# Переменные места ищем по опорным словам. Значение подставляется между
# ними. Так работает и с пустым шаблоном (места пустые), и с заполненным
# образцом — в обоих случаях находим границы по одним и тем же словам.
SUM_MARKS = [
    ('взысканные в сумме ', ' и государственную пошлину'),
    ('в сумме ', ' и государственную пошлину'),
]
GOSH_MARKS = [
    ('пошлину в сумме ', ', по исполнительному листу'),
    ('пошлину в сумме ', ' по судебному приказу'),
    ('пошлину в сумме ', ', по судебному приказу'),
]
DOCNO_MARKS = [
    ('исполнительному листу № ', ' года, выданному'),
    ('исполнительному листу № ', ', выданному'),
    ('судебному приказу дело № ', ', выданному'),
    ('судебному приказу № ', ', выданному'),
]
COURT_MARKS = [
    ('выданному: ', ' в отношении должника'),
    ('выданному:', ' в отношении должника'),
]


def _span_between(text, start_mark, end_mark):
    """(начало, конец) значения между опорными словами — без лишних пробелов."""
    i = text.find(start_mark)
    if i == -1:
        return None
    start = i + len(start_mark)
    j = text.find(end_mark, start)
    if j == -1:
        return None
    while start < j and text[start] == ' ':
        start += 1
    end = j
    while end > start and text[end - 1] == ' ':
        end -= 1
    return start, end


def _find_marks(text, marks):
    for sm, em in marks:
        span = _span_between(text, sm, em)
        if span:
            return span
    return None


def analyze(doc):
    """Найти в шаблоне заявления переменные места.

    Возвращает {поле: (абзац, начало, конец)} — границы в карте абзаца
    (см. par_map). Пустой диапазон (начало == конец) — место вставки.
    """
    slots = {}
    paras = doc.paragraphs

    # абзац с «Прошу перечислить…» — суммы, номер документа, суд
    for i, par in enumerate(paras):
        text, _pos = par_map(par)
        if 'Прошу перечислить' not in text:
            continue
        span = _find_marks(text, SUM_MARKS)
        if span:
            slots['сумма'] = (i,) + span
        span = _find_marks(text, GOSH_MARKS)
        if span:
            slots['госпошлина'] = (i,) + span
        span = _find_marks(text, DOCNO_MARKS)
        if span:
            slots['номер_документа'] = (i,) + span
        span = _find_marks(text, COURT_MARKS)
        if span:
            slots['суд'] = (i,) + span
        break

    # «Подлинник исполнительного листа дела № …» — тот же номер документа
    for i, par in enumerate(paras):
        text, _pos = par_map(par)
        if 'Подлинник' not in text:
            continue
        for sm, em in (('исполнительного листа дела № ', '. в 1 экз'),
                       ('исполнительного листа дела № ', ' в 1 экз'),
                       ('судебного приказа дела № ', '. в 1 экз'),
                       ('судебного приказа дела № ', ' в 1 экз')):
            span = _span_between(text, sm, em)
            if span:
                slots['номер_приложения'] = (i,) + span
                break
        break

    # дата и номер заявления — строка «От __15.07.2026 г. № б/н…»
    for i, par in enumerate(paras):
        text, _pos = par_map(par)
        if not text.strip().startswith('От '):
            continue
        m = re.search(r'От\s+_*[\d]{1,2}\.[\d]{1,2}\.[\d]{4}_*\s*г_*', text)
        if m:
            slots['дата_заявления'] = (i, m.start(), m.end())
        m = re.search(r'№\s*_*б/н_*', text)
        if m:
            slots['номер_заявления'] = (i, m.start(), m.end())
        break

    # ФИО и адрес — абзац с «зарегистрирован(а) по адресу:»
    for i, par in enumerate(paras):
        text, positions = par_map(par)
        if 'зарегистрирован(а) по адресу' not in text:
            continue
        k = text.find('зарегистрирован(а) по адресу')
        # ФИО стоит непосредственно перед «зарегистрирован(а) по адресу».
        # В одних шаблонах оно уже стоит (образец), в других — на его месте
        # пустые run'ы: тогда ФИО вписывается туда же.
        trimmed = text[:k].rstrip()
        name = trimmed[:-1].rstrip() if trimmed.endswith(',') else trimmed
        if name:
            start = text.rfind(name, 0, k)
            if start == -1:
                start = 0
            slots['фио'] = (i, start, start + len(name))
        else:
            # пустое место — ищем по run'ам перед «зарегистрирован(а)»
            empty_pos = [n for n, (_r, _o, empty) in enumerate(positions[:k])
                         if empty]
            if empty_pos:
                slots['фио'] = (i, empty_pos[0], empty_pos[-1] + 1)
            elif i > 0:
                prev, prev_pos = par_map(paras[i - 1])
                mark = 'в отношении должника:'
                p = prev.find(mark)
                if p != -1:
                    slots['фио'] = (i - 1, p + len(mark), p + len(mark))
                else:
                    slots['фио'] = (i, 0, 0)
            else:
                slots['фио'] = (i, 0, 0)
        # адрес — после «по адресу:»
        p = text.find('по адресу:', k)
        if p != -1:
            a_start = p + len('по адресу:')
            a_end = len(text)
            while a_end > a_start and text[a_end - 1] in ' .':
                a_end -= 1
            while a_start < a_end and text[a_start] == ' ':
                a_start += 1
            slots['адрес'] = (i, a_start, a_end)
        break
    return slots


def describe_slots(slots):
    """Человекочитаемое описание найденных мест — показываем юристу."""
    names = {'сумма': 'сумма долга', 'госпошлина': 'госпошлина',
             'номер_документа': '№ исполнительного документа',
             'номер_приложения': '№ документа в приложениях',
             'дата_заявления': 'дата заявления',
             'суд': 'суд, выдавший документ', 'фио': 'ФИО должника',
             'адрес': 'адрес должника'}
    return [(names.get(k, k), v[0] + 1) for k, v in slots.items()]


def extract_requisites(doc):
    """Реквизиты взыскателя из шапки шаблона — для листа «Настройки»."""
    out = {}
    paras = doc.paragraphs
    for i, p in enumerate(paras):
        t = p.text
        if 'Адрес юридический:' in t:
            out['юр_адрес_1'] = t.split('Адрес юридический:', 1)[1].strip()
            if i + 1 < len(paras):
                out['юр_адрес_2'] = paras[i + 1].text.strip()
        elif 'Адрес почтовый:' in t:
            out['почт_адрес_1'] = t.split('Адрес почтовый:', 1)[1].strip()
            if i + 1 < len(paras):
                out['почт_адрес_2'] = paras[i + 1].text.strip()
        elif t.strip().startswith('Тел.'):
            out['телефон'] = t.split('Тел.', 1)[1].strip().rstrip(',')
        elif t.strip().startswith('факс'):
            out['факс'] = t.split('факс', 1)[1].strip()
        elif 'Электронный адрес' in t:
            out['email'] = t.split('Электронный адрес', 1)[1].lstrip(': ')
        elif t.strip().startswith('Р/с'):
            out['р_с'] = t.split('Р/с', 1)[1].strip()
        elif t.strip().startswith('К/с'):
            out['к_с'] = t.split('К/с', 1)[1].strip()
        elif t.strip().startswith('БИК'):
            rest = t.split('БИК', 1)[1].strip()
            parts = rest.split('ИНН')
            out['бик'] = parts[0].strip()
            if len(parts) > 1:
                out['инн'] = parts[1].strip()
        elif t.strip().startswith('КПП'):
            out['кпп'] = t.split('КПП', 1)[1].strip()
        elif 'ОГРН' in t:
            m = re.search(r'ОГРН[:\s]*([\d]+)', t)
            if m:
                out['огрн'] = m.group(1)
        elif 'Взыскатель:' in t:
            m = re.search(r'Взыскатель:\s*([^,;]+)', t)
            if m:
                out['взыскатель'] = m.group(1).strip()
        elif t.strip().startswith('Зарегистрировано:'):
            out['регистрация'] = t.split('Зарегистрировано:', 1)[1].strip(' ;')
    for i, par in enumerate(paras):
        txt = par.text.strip()
        if re.match(r'(ОБЩЕСТВО|АКЦИОНЕРНОЕ|ИНДИВИДУАЛЬНЫЙ|ПУБЛИЧНОЕ)', txt):
            out['наименование_1'] = txt
            if i + 1 < len(paras) and paras[i + 1].text.strip():
                out['наименование_2'] = paras[i + 1].text.strip()
            break
    return out


# ──────────────────────────────────────────────────────── книга раздела ──
DEFAULT_SETTINGS = [
    ('шаблон_сп', ''),
    ('шаблон_ил', ''),
    ('шаблон_сопровод', ''),
    ('взыскатель', ''),
    ('наименование_1', ''),
    ('наименование_2', ''),
    ('юридический адрес', ''),
    ('почтовый адрес', ''),
    ('телефон', ''),
    ('факс', ''),
    ('email', ''),
    ('р_с', ''),
    ('к_с', ''),
    ('БИК', ''),
    ('ИНН', ''),
    ('КПП', ''),
    ('ОГРН', ''),
    ('регистрация', ''),
    ('дата_заявления', ''),
    ('номер_заявления', ''),
]

DOC_HEADERS = ['№', 'Вид документа', '№ исполнительного документа',
               'Дата документа', 'ФИО должника', 'Адрес',
               'Сумма долга', 'Госпошлина', 'Период', 'Суд', 'Примечание']


def sheet_settings(ws):
    """Настройки из листа «Настройки» → словарь (ключи в нижнем регистре)."""
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        key = str(row[0]).strip().lower().replace(' ', '_')
        val = row[1] if len(row) > 1 else ''
        out[key] = '' if val is None else val
    return out


def sheet_docs(ws):
    """Строки листа «Документы» → список словарей."""
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or all(v is None or str(v).strip() == '' for v in row):
            continue
        vals = list(row) + [None] * (len(DOC_HEADERS) - len(row))
        out.append({
            '№': vals[0], 'вид': vals[1], 'номер': vals[2], 'дата': vals[3],
            'фио': vals[4], 'адрес': vals[5], 'сумма': vals[6],
            'госпошлина': vals[7], 'период': vals[8], 'суд': vals[9],
            'примечание': vals[10],
        })
    return out


def read_book(path):
    """Читает книгу раздела: (настройки, заявления)."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    settings = (sheet_settings(wb['Настройки'])
                if 'Настройки' in wb.sheetnames else {})
    docs = sheet_docs(wb['Заявления']) if 'Заявления' in wb.sheetnames else []
    return settings, docs


def ensure_settings(ws):
    """Добавить в лист «Настройки» недостающие строки параметров."""
    rows = {}
    for r in range(2, ws.max_row + 1):
        key = str(ws.cell(row=r, column=1).value or '').strip().lower()
        if key:
            rows[key] = r
    for key, val in DEFAULT_SETTINGS:
        low = key.strip().lower()
        if low not in rows:
            r = ws.max_row + 1
            ws.cell(row=r, column=1, value=key)
            ws.cell(row=r, column=2, value=val)
            rows[low] = r
    return rows


# ──────────────────────────────────────── сборка заявлений (Word) ──
def doc_kind(doc):
    """Вид заявления по шаблону: «сп» (судебный приказ) или «ил» (лист)."""
    text = '\n'.join(p.text for p in doc.paragraphs).lower()
    if 'исполнительном листе' in text or 'исполнительного листа' in text:
        return 'ил'
    if 'судебному приказу' in text or 'судебного приказа' in text:
        return 'сп'
    return 'сп'


def _values_for_row(d, settings):
    """Значения, которые подставляются в шаблон по строке таблицы."""
    values = {}
    for key, field in (('фио', 'фио'), ('адрес', 'адрес'),
                       ('суд', 'суд'), ('номер_документа', 'номер')):
        val = str(d.get(field) or '').strip()
        if val:
            values[key] = val
    amount = d.get('сумма')
    if amount not in ('', None):
        values['сумма'] = money_words(amount)
    duty = d.get('госпошлина')
    if duty not in ('', None):
        values['госпошлина'] = money_words(duty)
    doc_date = d.get('дата')
    if doc_date not in ('', None):
        values['дата_документа'] = date_plain(doc_date)
    # номер документа — вместе с датой («2-1526/2026 от 19.05.2026»)
    if values.get('номер_документа'):
        full = values['номер_документа']
        if values.get('дата_документа') and ' от ' not in full:
            full = '%s от %s' % (full, values['дата_документа'])
        values['номер_документа'] = full
        values['номер_приложения'] = full
    claim_date = str(settings.get('дата_заявления') or '').strip()
    if claim_date:
        values['дата_заявления'] = 'От %s г' % date_plain(claim_date)
    claim_no = str(settings.get('номер_заявления') or '').strip()
    if claim_no:
        values['номер_заявления'] = '№ %s' % claim_no
    return values


def build_document(items, settings, out_path):
    """Собрать заявления: по одному на документ, каждое с новой страницы.

    items — список пар (шаблон, строка таблицы). Шаблон указывает вид
    заявления: судебный приказ или исполнительный лист.

    Возвращает (слоты по шаблонам, предупреждения).
    """
    from copy import deepcopy

    if not items:
        raise RuntimeError('нет строк для сборки')

    first_path = items[0][0]
    doc = docx.Document(str(first_path))
    body = doc.element.body
    sect = None
    for child in list(body):
        if child.tag.endswith('}sectPr'):
            sect = child
        body.remove(child)

    settings = {str(k).strip().lower(): v for k, v in (settings or {}).items()}
    notes = []
    all_slots = {}
    first = True

    for n, (template_path, d) in enumerate(items, 1):
        src = docx.Document(str(template_path))
        slots = analyze(src)
        all_slots.setdefault(str(template_path), slots)
        template_block = [c for c in src.element.body
                          if not c.tag.endswith('}sectPr')]

        before = len(doc.paragraphs)
        for el in template_block:
            body.append(deepcopy(el))
        paras = doc.paragraphs[before:]
        for par in paras:
            unlink_fields(par)

        values = _values_for_row(d, settings)
        if not str(d.get('фио') or '').strip():
            notes.append('строка %d: нет ФИО — проверьте' % n)

        by_para = {}
        for key, (pi, s, e) in slots.items():
            if key not in values:
                continue
            by_para.setdefault(pi, []).append((s, e, values[key]))
        for pi in sorted(by_para, reverse=True):
            if pi >= len(paras):
                continue
            for s, e, val in sorted(by_para[pi], reverse=True):
                apply_value(paras[pi], s, e, val)

        if not first and paras:
            paras[0].paragraph_format.page_break_before = True
        first = False

    if sect is not None:
        body.append(sect)
    doc.save(str(out_path))
    return all_slots, notes


# ──────────────────────────────────────── сопровод приставу (таблица) ──
def build_accompanying(template_path, docs, out_path):
    """Заполнить таблицу в шаблоне сопровода: одна строка на документ.

    docs — список словарей со строками таблицы раздела.
    Возвращает число заполненных строк.
    """
    doc = docx.Document(str(template_path))
    if not doc.tables:
        return 0
    table = doc.tables[0]
    rows = table.rows
    if len(rows) < 2:
        return 0

    def fill(row, d):
        cells = row.cells
        values = [
            str(d.get('№') or ''),
            str(d.get('фио') or ''),
            str(d.get('адрес') or ''),
            str(d.get('номер') or ''),
            str(d.get('суд') or ''),
            money_words(d.get('сумма')) if d.get('сумма') not in ('', None) else '',
            money_words(d.get('госпошлина')) if d.get('госпошлина') not in ('', None) else '',
        ]
        for i, cell in enumerate(cells):
            if i < len(values):
                _set_cell(cell, values[i])

    # образец строки — строка с данными (в шаблоне юриста их много)
    sample = None
    for row in rows[1:]:
        if any(c.text.strip() for c in row.cells):
            sample = row
            break
    if sample is None:
        sample = rows[1]
    sample_tr = sample._tr

    # лишние строки убираем — оставляем заголовок и одну образцовую
    while len(table.rows) > 2:
        table._tbl.remove(table.rows[-1]._tr)

    for n, d in enumerate(docs, 1):
        if n == 1:
            fill(rows[1], d)
            continue
        table._tbl.append(deepcopy(sample_tr))
        fill(table.rows[-1], d)
    doc.save(str(out_path))
    return len(docs)


def _set_cell(cell, text):
    """Записать текст в ячейку, сохранив оформление первого run'а."""
    if not text:
        return
    par = cell.paragraphs[0] if cell.paragraphs else None
    if par is None:
        return
    runs = list(par.runs)
    if not runs:
        par.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r.text = ''
