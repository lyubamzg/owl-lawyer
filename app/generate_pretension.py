#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сборка документов раздела «Досудебные претензии».

Читает книгу pretension-input.xlsx, берёт шаблон претензии и собирает
файл Word: одна претензия на строку списка, каждая с новой страницы.
Рядом кладёт реестр (registry.xlsx) и отчёт сверки.
"""
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

import pretension_lib as pl

BASE = Path(__file__).resolve().parent
BOOK = BASE / pl.BOOK
REPORT = BASE / pl.REPORT
OUT_DIR = BASE / 'output-pretension'


def main(argv):
    book = Path(argv[1]) if len(argv) > 1 else BOOK
    if not book.is_file():
        print('Нет книги %s — сначала загрузите список.' % book.name)
        return 2
    settings, claims = pl.read_book(book)
    if not claims:
        print('В книге нет строк претензий — загрузите список.')
        return 2

    addresses = [str(c.get('адрес') or '') for c in claims]
    template = pl.pick_template(settings, addresses)
    if not template or not template.is_file():
        print('Не найден шаблон претензии. Положите .docx в папку %s'
              % pl.TEMPLATE_DIR.name)
        return 2

    stamp = datetime.now().strftime('%Y-%m-%d_%H-%M')
    out_dir = OUT_DIR / stamp
    out_dir.mkdir(parents=True, exist_ok=True)

    # имя файла — по исходящему номеру и дате из претензии:
    # «Претензии № 906 от 24.11.2026.docx»
    try:
        num_from = int(str(settings.get('претензия_номер_от') or 1).strip() or 1)
    except ValueError:
        num_from = 1
    first = claims[0]
    first_number = first.get('номер')
    first_number = (int(str(first_number).strip())
                    if str(first_number or '').strip().isdigit()
                    else num_from)
    first_date = (first.get('дата_претензии')
                  or settings.get('дата_претензии')
                  or datetime.now().strftime('%d.%m.%Y'))
    first_date = pl.date_plain(first_date)
    out_path = out_dir / ('Претензии № %s от %s.docx' % (first_number, first_date))

    warnings = []
    for n, c in enumerate(claims, 1):
        if not str(c.get('фио') or '').strip():
            warnings.append('строка %d: нет ФИО — претензия пропущена' % n)
            continue
        if not str(c.get('адрес') or '').strip():
            warnings.append('%s: нет адреса' % c.get('фио'))
        if c.get('сумма') in ('', None):
            warnings.append('%s: не указана сумма долга' % c.get('фио'))

    try:
        slots = pl.build_document(template, claims, settings, out_path)
    except Exception as ex:  # noqa: BLE001
        print('Не удалось собрать документы: %s' % ex)
        return 1

    # реестр
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Реестр'
    heads = ['№', 'Лицевой счёт', 'ФИО', 'Адрес', 'Сумма долга',
             '№ претензии', 'Дата претензии', 'Дата состояния', 'Период', 'Файл']
    for i, h in enumerate(heads, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.fill = PatternFill('solid', fgColor='FFF2CC')
        c.font = Font(bold=True)
    total = 0.0
    for n, c in enumerate(claims, 1):
        number = c.get('номер')
        number = int(str(number).strip()) if str(number or '').strip().isdigit() \
            else num_from + n - 1
        amount = pl.money(c.get('сумма'))
        try:
            total += float(str(c.get('сумма') or 0).replace(' ', '').replace(',', '.'))
        except (TypeError, ValueError):
            pass
        vals = [n, c.get('лс'), c.get('фио'), c.get('адрес'), amount, number,
                pl.date_plain(c.get('дата_претензии') or
                              settings.get('дата_претензии') or ''),
                pl.date_words(c.get('дата_состояния') or
                              settings.get('дата_состояния') or ''),
                c.get('период') or settings.get('период') or '',
                out_path.name]
        for i, v in enumerate(vals, 1):
            ws.cell(row=n + 1, column=i, value=v)
    ws.cell(row=len(claims) + 2, column=4, value='ИТОГО').font = Font(bold=True)
    ws.cell(row=len(claims) + 2, column=5, value=pl.money(total)).font = Font(bold=True)
    wb.save(out_dir / 'registry.xlsx')

    # отчёт
    lines = ['СБОРКА ДОСУДЕБНЫХ ПРЕТЕНЗИЙ', '=' * 40, '']
    lines.append('Книга: %s' % book.name)
    lines.append('Шаблон: %s' % template.name)
    lines.append('Претензий: %d' % len(claims))
    lines.append('Итого долга: %s руб.' % pl.money(total))
    lines.append('Файл: %s' % out_path.name)
    lines.append('')
    lines.append('ПОДСТАНОВЛЕННЫЕ ПОЛЯ ШАБЛОНА:')
    for name, para in pl.describe_slots(slots):
        lines.append('  • %s — абзац %d' % (name, para))
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

    print('СБОРКА ДОСУДЕБНЫХ ПРЕТЕНЗИЙ')
    print('  шаблон: %s' % template.name)
    print('  претензий: %d, итого долга: %s руб.' % (len(claims), pl.money(total)))
    if warnings:
        print('  предупреждений: %d (см. %s)' % (len(warnings), pl.REPORT))
    print('  файл: %s' % out_path.name)
    print('  все %d претензий — в одном файле, каждая с новой страницы' % len(claims))
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
