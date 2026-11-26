#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Загрузка списка должников в раздел «Досудебные претензии».

Принимает файлы, выбранные в приложении:
  *.xlsx — список должников (лицевой счёт, ФИО, адрес, сумма);
  *.docx — шаблон претензии (необязательно: если его нет, берётся шаблон
            рядом с программой; см. pretension-templates/).

Список читается как есть: шапки может не быть, пустые строки пропускаются,
по одной строке — одна претензия.
"""
import re
import shutil
import sys
from pathlib import Path

import openpyxl

import pretension_lib as pl

BASE = Path(__file__).resolve().parent
BOOK = BASE / pl.BOOK
REPORT = BASE / pl.REPORT

HEADER_ALIASES = {
    'ф': 'фио', 'фио': 'фио', 'ф.и.о': 'фио', 'ф и о': 'фио',
    'должник': 'фио', 'должники': 'фио', 'плательщик': 'фио',
    'адрес': 'адрес', 'адрес должника': 'адрес', 'адрес регистрации': 'адрес',
    'сумма': 'сумма', 'сумма долга': 'сумма', 'долг': 'сумма',
    'сумма задолженности': 'сумма', 'начислено': 'сумма',
    'лс': 'лс', 'лицевой счет': 'лс', 'лицевой счёт': 'лс',
    'лицевойсчет': 'лс', 'номер лицевого счета': 'лс',
}


def norm_key(text):
    t = str(text or '').strip().lower().replace('ё', 'е')
    t = re.sub(r'[.,()]', ' ', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t


def looks_like_header(row):
    cells = [norm_key(c) for c in row if c not in (None, '')]
    if not cells:
        return False
    hits = sum(1 for c in cells if c in HEADER_ALIASES)
    return hits >= 2


def parse_list(path):
    """Читает список: (строки, заголовки, предупреждения)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    rows, warnings = [], []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        header = None
        for r, row in enumerate(ws.iter_rows(values_only=True), 1):
            vals = list(row)
            if all(v is None or str(v).strip() == '' for v in vals):
                continue
            if header is None and looks_like_header(vals):
                header = {}
                for i, v in enumerate(vals):
                    key = HEADER_ALIASES.get(norm_key(v))
                    if key and key not in header:
                        header[key] = i
                continue
            if header is None:
                # без шапки: лицевой счёт, ФИО, адрес, сумма
                idx = [i for i, v in enumerate(vals)
                       if v is not None and str(v).strip() != '']
                if len(idx) >= 4:
                    header = {'лс': idx[0], 'фио': idx[1], 'адрес': idx[2],
                              'сумма': idx[3]}
                elif len(idx) == 3:
                    header = {'фио': idx[0], 'адрес': idx[1], 'сумма': idx[2]}
                else:
                    warnings.append(
                        'лист «%s», строка %d: не понял столбцы — пропущена' % (sheet, r))
                    continue
            rec = {}
            for key, i in header.items():
                rec[key] = vals[i] if i < len(vals) else None
            fio = str(rec.get('фио') or '').strip()
            if not fio:
                warnings.append('лист «%s», строка %d: нет ФИО — пропущена' % (sheet, r))
                continue
            rec['лист'] = sheet
            rec['строка'] = r
            rows.append(rec)
    return rows, warnings


def clean_amount(value):
    if value is None:
        return ''
    if isinstance(value, (int, float)):
        return round(float(value), 2)
    s = str(value).strip().replace(' ', '').replace(' ', '')
    s = s.replace(',', '.')
    try:
        return round(float(s), 2)
    except ValueError:
        return value


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


def _set_stored(ws, key, value):
    r = _find_row(ws, key)
    if r:
        ws.cell(row=r, column=2, value=value)
    else:
        r = ws.max_row + 1
        ws.cell(row=r, column=1, value=key)
        ws.cell(row=r, column=2, value=value)


def main(argv):
    files = [Path(a) for a in argv[1:]]
    lists = [f for f in files if f.suffix.lower() in ('.xlsx', '.xlsm', '.ods')]
    templates = [f for f in files if f.suffix.lower() == '.docx']
    if not lists:
        print('Нужен файл Excel со списком должников (xlsx).')
        print('Пример: python3 import_pretension.py "список должников.xlsx"')
        return 2
    list_path = lists[0]
    if not BOOK.exists():
        import create_pretension_workbook
        create_pretension_workbook.build(BOOK)

    rows, warnings = parse_list(list_path)

    # шаблон: свой файл или встроенный (по адресам из списка)
    addresses = [str(r.get('адрес') or '') for r in rows]
    template = None
    if templates:
        template = pl.TEMPLATE_DIR / templates[0].name
        pl.TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
        shutil.copy2(templates[0], template)
        print('Шаблон претензии: %s' % templates[0].name)
    else:
        template = pl.pick_template({}, addresses)

    wb = openpyxl.load_workbook(BOOK)
    import create_pretension_workbook
    create_pretension_workbook.refresh_instructions(wb)
    if 'Шаблон' not in wb.sheetnames:
        tm = wb.create_sheet('Шаблон')
        tm['A1'], tm['B1'] = 'Параметр', 'Значение из шаблона'
    ws = wb['Претензии']
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.value = None

    # точные повторы (та же ФИО, адрес и сумма) по умолчанию пропускаем:
    # иначе один и тот же должник получит два одинаковых письма
    settings_now = pl.sheet_settings(wb['Настройки'])
    skip_dupes = str(settings_now.get('точные_повторы_пропускать') or 'да').strip().lower() \
        not in ('нет', 'no', 'off', '0', 'выкл')

    seen = {}
    exact = {}
    written = 0
    skipped = 0
    for n, rec in enumerate(rows, 1):
        fio = str(rec.get('фио') or '').strip()
        addr = str(rec.get('адрес') or '').strip()
        amount = clean_amount(rec.get('сумма'))
        ls = rec.get('лс')
        ls = '' if ls is None else str(ls).strip()
        if not addr:
            warnings.append('строка %d: нет адреса' % rec['строка'])
        if amount in ('', None):
            warnings.append('%s: не указана сумма долга' % fio)
        key = (fio.lower(), addr.lower())
        exact_key = key + (str(amount),)
        if skip_dupes and exact_key in exact:
            skipped += 1
            warnings.append('%s: точный повтор строки %d — пропущен (строка %d)' %
                            (fio, exact[exact_key], rec['строка']))
            continue
        exact.setdefault(exact_key, rec['строка'])
        if key in seen:
            warnings.append('%s: повтор в строке %d (строка %d) — проверьте' %
                            (fio, seen[key], rec['строка']))
        else:
            seen[key] = rec['строка']
        values = [n, ls, fio, addr, amount, '', '', '', '', '']
        for i, v in enumerate(values, 1):
            ws.cell(row=n + 1, column=i, value=v)
        written += 1

    # настройки: шапка из шаблона (если её ещё не заполняли)
    st = wb['Настройки']
    srows = pl.ensure_settings(st)
    current = pl.sheet_settings(st)
    # старый параметр «претензия_номер_от» → «исходящий_номер_и_дата»
    legacy = str(current.get('претензия_номер_от') or '').strip()
    if legacy:
        if not str(current.get('исходящий_номер_и_дата') or '').strip():
            st.cell(row=srows['исходящий_номер_и_дата'], column=2,
                    value=legacy)
            warnings.append('исходящий номер и дата перенесены из «претензия_номер_от»')
        r = srows.get('претензия_номер_от')
        if r:
            st.delete_rows(r)
            srows = pl.ensure_settings(st)
    current = pl.sheet_settings(st)
    if template and template.is_file():
        import docx
        # ключи приводим к нижнему регистру — так же, как в листе «Настройки»
        head = {str(k).strip().lower(): v
                for k, v in pl.extract_header(docx.Document(str(template))).items()}
        head['шаблон'] = template.name
        # значения, которые в прошлый раз были взяты из шаблона: если юрист их
        # не менял руками — обновляем под новый шаблон; если менял — оставляем
        tm = wb['Шаблон']
        stored = {}
        for r in range(2, tm.max_row + 1):
            k = str(tm.cell(row=r, column=1).value or '').strip().lower()
            if k:
                stored[k] = tm.cell(row=r, column=2).value
        kept = []
        for key, val in head.items():
            if key == 'шаблон':
                _set_setting(st, key, val)
                continue
            now = str(current.get(key) or '').strip()
            if not now:
                if val:                       # пусто — берём из шаблона
                    _set_setting(st, key, val)
                continue
            if val and now != str(val).strip():
                # в таблице стоит другое значение, чем в шаблоне — это правка юриста
                kept.append(key)
        for key, val in head.items():
            _set_stored(tm, key, val)
        if kept:
            warnings.append('шапка: значения, отличающиеся от шаблона, оставлены '
                            'как есть (правка юриста): ' + ', '.join(kept))
    wb.save(BOOK)

    # отчёт
    lines = ['ЗАГРУЗКА СПИСКА ПРЕТЕНЗИЙ', '=' * 40, '']
    lines.append('Список: %s' % list_path.name)
    lines.append('Строк в списке: %d, записано в таблицу: %d' % (len(rows), written))
    if skipped:
        lines.append('Пропущено точных повторов: %d (настройка «точные повторы пропускать»)' % skipped)
    if template:
        lines.append('Шаблон претензии: %s' % template.name)
    lines.append('Книга: %s' % pl.BOOK)
    lines.append('')
    lines.append('ЧТО ПРОВЕРИТЬ В ТАБЛИЦЕ (лист «Настройки»):')
    lines.append('  • шапка — реквизиты взыскателя;')
    lines.append('  • исходящий_номер_и_дата — «100 от 10.11.2026»: исходящий номер')
    lines.append('    и дата претензий; одни на весь пакет (заполняет сотрудник);')
    lines.append('  • дата_состояния — «по состоянию на …» (долг на какую дату);')
    lines.append('    если не заполнено — берётся 1-е число месяца исходящей даты;')
    lines.append('  • период — «за жилищно-коммунальные услуги» (что именно за период);')
    lines.append('  • срок_оплаты, дней — сколько дней на оплату.')
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

    print('ЗАГРУЗКА СПИСКА ПРЕТЕНЗИЙ')
    print('  список: %s — строк: %d' % (list_path.name, len(rows)))
    print('  записано в «%s»: %d%s' % (pl.BOOK, written,
          (' (пропущено повторов: %d)' % skipped) if skipped else ''))
    if template:
        print('  шаблон: %s' % template.name)
    if warnings:
        print('  предупреждений: %d (см. %s)' % (len(warnings), pl.REPORT))
    print('Книга сохранена: %s' % BOOK)
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
