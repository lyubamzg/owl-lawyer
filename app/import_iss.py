#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Загрузка данных в раздел «Исполнительное производство».

Принимает файлы, выбранные в приложении:
  *.pdf  — сканы судебных приказов: приложение распознаёт их (OCR) и
            вытаскивает номер дела, дату, должников, адрес, суммы;
  *.xlsx — готовый реестр исполнительных документов в ФССП;
  *.docx — шаблоны заявлений (необязательно: есть свои в iss-templates/).

Результат — книга iss-input.xlsx, лист «Заявления»: одна строка на
должника. Данные из разных файлов дописываются в конец списка.
"""
import re
import shutil
import sys
from pathlib import Path

import openpyxl

import iss_lib as il

BASE = Path(__file__).resolve().parent
BOOK = BASE / il.BOOK
REPORT = BASE / il.REPORT

# заголовки реестра ФССП → наши поля
HEADER_ALIASES = {
    'ф': 'фио', 'фио': 'фио', 'ф.и.о.должника': 'фио',
    'ф.и.о': 'фио', 'должник': 'фио', 'фио должника': 'фио',
    'адрес': 'адрес', 'адрес должника': 'адрес',
    '№ исполнительного документа': 'номер',
    'номер исполнительного документа': 'номер',
    'исполнительный документ': 'номер',
    'судебный орган, выдавший исполнительный документ': 'суд',
    'судебный орган': 'суд', 'суд': 'суд',
    'сумма основного долга': 'сумма', 'сумма долга': 'сумма',
    'сумма': 'сумма', 'основной долг': 'сумма',
    'сумма гос.пошлины': 'госпошлина', 'госпошлина': 'госпошлина',
    'сумма госпошлины': 'госпошлина',
    '№ п/п': 'номер_п', 'номер п/п': 'номер_п',
    'дата документа': 'дата', 'дата': 'дата',
}


def norm_key(text):
    t = str(text or '').strip().lower().replace('ё', 'е')
    t = re.sub(r'\s+', ' ', t)
    return t


def clean_amount(value):
    """«7916 руб. 74 коп.» → 7916.74."""
    if value is None:
        return ''
    if isinstance(value, (int, float)):
        return round(float(value), 2)
    s = str(value).strip()
    if not s:
        return ''
    m = re.search(r'(\d[\d\s]*)[,.](\d{1,2})', s)
    if m:
        return round(float(m.group(1).replace(' ', '')) +
                     float('0.' + m.group(2)), 2)
    m = re.search(r'([\d\s]+)', s)
    if m:
        try:
            return round(float(m.group(1).replace(' ', '')), 2)
        except ValueError:
            return s
    return s


def parse_registry(path):
    """Читает реестр ФССП: (строки, предупреждения)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    rows, warnings = [], []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        header = None
        for r, row in enumerate(ws.iter_rows(values_only=True), 1):
            vals = list(row)
            if all(v is None or str(v).strip() == '' for v in vals):
                continue
            if header is None:
                header = {}
                for i, v in enumerate(vals):
                    key = HEADER_ALIASES.get(norm_key(v))
                    if key and key not in header:
                        header[key] = i
                if not header:
                    # шапки нет: пробуем по порядку столбцов
                    idx = [i for i, v in enumerate(vals)
                           if v is not None and str(v).strip() != '']
                    if len(idx) >= 4:
                        header = {'фио': idx[0], 'адрес': idx[1],
                                  'номер': idx[2], 'суд': idx[3],
                                  'сумма': idx[4] if len(idx) > 4 else None}
                    else:
                        warnings.append('лист «%s», строка %d: не понял '
                                        'столбцы — пропущена' % (sheet, r))
                        continue
                continue
            rec = {}
            for key, i in header.items():
                rec[key] = vals[i] if i is not None and i < len(vals) else None
            fio = str(rec.get('фио') or '').strip()
            if not fio:
                warnings.append('лист «%s», строка %d: нет ФИО — пропущена'
                                % (sheet, r))
                continue
            rec['лист'] = sheet
            rec['строка'] = r
            rows.append(rec)
    return rows, warnings


def parse_orders_pdf(path):
    """Сканы судебных приказов → (строки, предупреждения).

    Один и тот же приказ может быть отсканирован дважды — такие строки
    пропускаем (по номеру дела и дате).
    """
    warnings = []
    try:
        pages = il.ocr_pdf(path)
    except Exception as ex:  # noqa: BLE001
        return [], ['не удалось распознать %s: %s' % (path.name, ex)]
    rows = []
    seen = {}
    for n, page in enumerate(pages, 1):
        parsed = il.parse_order(page)
        if not parsed.get('номер_дела') and not parsed.get('сумма'):
            warnings.append('страница %d: не похоже на судебный приказ — '
                            'пропущена' % n)
            continue
        key = (parsed.get('номер_дела', ''), str(parsed.get('дата') or ''))
        if key in seen:
            warnings.append('страница %d: повтор приказа %s (уже есть '
                            'строка %d) — пропущена'
                            % (n, parsed.get('номер_дела') or '?',
                               seen[key]))
            continue
        seen[key] = len(rows) + 1
        fio = parsed.get('фио') or ''
        names = [il.fio_to_nominative(x) for x in parsed.get('должников') or []]
        if not names and fio:
            names = [il.fio_to_nominative(fio)]
        if not names:
            warnings.append('страница %d: не нашли ФИО должника — '
                            'пропущена' % n)
            continue
        # один судебный приказ — одна строка: все солидарные должники
        # перечисляются в одной графе «ФИО должника»
        rows.append({
            'фио': ', '.join(names),
            'адрес': parsed.get('адрес', ''),
            'номер': parsed.get('номер_дела', ''),
            'дата': parsed.get('дата', ''),
            'суд': parsed.get('суд', ''),
            'сумма': parsed.get('сумма', ''),
            'госпошлина': parsed.get('госпошлина', ''),
            'период': parsed.get('период', ''),
            'вид': 'судебный приказ',
            'примечание': 'стр. %d' % n,
        })
    return rows, warnings


def _find_row(ws, key):
    key = str(key).strip().lower()
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=1).value or '').strip().lower() == key:
            return r
    return None


def _set_setting(ws, key, value):
    r = _find_row(ws, key)
    if r:
        ws.cell(row=r, column=2, value=value)


def main(argv):
    files = [Path(a) for a in argv[1:]]
    pdfs = [f for f in files if f.suffix.lower() == '.pdf']
    books = [f for f in files if f.suffix.lower() in ('.xlsx', '.xlsm', '.ods')]
    templates = [f for f in files if f.suffix.lower() == '.docx']
    if not pdfs and not books:
        print('Нужен PDF со сканами судебных приказов или Excel с реестром.')
        print('Пример: python3 import_iss.py "приказы.pdf" "реестр ФССП.xlsx"')
        return 2

    if not BOOK.exists():
        import create_iss_workbook
        create_iss_workbook.build(BOOK)

    new_rows = []
    warnings = []

    for pdf in pdfs:
        rows, warns = parse_orders_pdf(pdf)
        new_rows.extend(rows)
        warnings.extend('%s: %s' % (pdf.name, w) for w in warns)
        print('Распознано приказов в «%s»: %d' % (pdf.name, len(rows)))

    for book in books:
        rows, warns = parse_registry(book)
        for rec in rows:
            new_rows.append({
                'фио': str(rec.get('фио') or '').strip(),
                'адрес': str(rec.get('адрес') or '').strip(),
                'номер': str(rec.get('номер') or '').strip(),
                'дата': rec.get('дата') or '',
                'суд': str(rec.get('суд') or '').strip(),
                'сумма': clean_amount(rec.get('сумма')),
                'госпошлина': clean_amount(rec.get('госпошлина')),
                'период': '',
                'вид': 'исполнительный лист',
                'примечание': '%s, стр. %s' % (rec.get('лист', ''),
                                                rec.get('строка', '')),
            })
        warnings.extend('%s: %s' % (book.name, w) for w in warns)
        print('Прочитано строк реестра из «%s»: %d' % (book.name, len(rows)))

    # шаблоны юриста — в папку раздела
    il.TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    for tpl in templates:
        dst = il.TEMPLATE_DIR / tpl.name
        shutil.copy2(tpl, dst)
        print('Шаблон: %s' % tpl.name)

    wb = openpyxl.load_workbook(BOOK)
    import create_iss_workbook
    create_iss_workbook.refresh_instructions(wb)
    if 'Заявления' not in wb.sheetnames:
        create_iss_workbook.build(BOOK)
        wb = openpyxl.load_workbook(BOOK)
    ws = wb['Заявления']

    # дописываем новые строки в конец
    start = ws.max_row
    while start > 1 and all(
            ws.cell(row=start, column=c).value in (None, '')
            for c in range(1, len(il.DOC_HEADERS) + 1)):
        start -= 1
    row_no = start

    for rec in new_rows:
        row_no += 1
        values = [
            row_no - 1,
            rec.get('вид') or '',
            rec.get('номер') or '',
            rec.get('дата') or '',
            rec.get('фио') or '',
            rec.get('адрес') or '',
            rec.get('сумма') if rec.get('сумма') != '' else '',
            rec.get('госпошлина') if rec.get('госпошлина') != '' else '',
            rec.get('период') or '',
            rec.get('суд') or '',
            rec.get('примечание') or '',
        ]
        for i, v in enumerate(values, 1):
            ws.cell(row=row_no, column=i, value=v)
        if not rec.get('адрес'):
            warnings.append('строка %d: нет адреса' % (row_no - 1))
        if rec.get('сумма') in ('', None):
            warnings.append('строка %d: не указана сумма долга' % (row_no - 1))

    # настройки: подставляем имена шаблонов, какие нашлись
    st = wb['Настройки']
    il.ensure_settings(st)
    current_settings = il.sheet_settings(st)
    for kind in ('сп', 'ил', 'сопровод'):
        key = 'шаблон_' + kind
        current = str(current_settings.get(key) or '').strip()
        if current and (il.TEMPLATE_DIR / current).is_file():
            continue
        found = il.pick_template(kind, {})
        if found:
            _set_setting(st, key, found.name)
    wb.save(BOOK)

    lines = ['ЗАГРУЗКА ДАННЫХ — ИСПОЛНИТЕЛЬНОЕ ПРОИЗВОДСТВО', '=' * 44, '']
    lines.append('Книга: %s' % il.BOOK)
    lines.append('Записано строк: %d' % len(new_rows))
    lines.append('')
    lines.append('ЧТО ПРОВЕРИТЬ В ТАБЛИЦЕ:')
    lines.append('  • лист «Заявления» — ФИО, адрес, номер документа, суммы;')
    lines.append('  • вид заявления — «судебный приказ» или «исполнительный лист»;')
    lines.append('  • лист «Настройки» — вид заявлений по умолчанию, шаблоны,')
    lines.append('    дату и номер заявления.')
    lines.append('')
    if warnings:
        lines.append('ПРЕДУПРЕЖДЕНИЯ (%d):' % len(warnings))
        for w in warnings[:60]:
            lines.append('  ⚠ ' + w)
        if len(warnings) > 60:
            lines.append('  … и ещё %d' % (len(warnings) - 60))
    else:
        lines.append('Предупреждений нет.')
    REPORT.write_text('\n'.join(lines) + '\n', encoding='utf-8')

    print('ЗАГРУЗКА ДАННЫХ — ИСПОЛНИТЕЛЬНОЕ ПРОИЗВОДСТВО')
    print('  строк записано: %d' % len(new_rows))
    if warnings:
        print('  предупреждений: %d (см. %s)' % (len(warnings), il.REPORT))
    print('Книга сохранена: %s' % BOOK)
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
