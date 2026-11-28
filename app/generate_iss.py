#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сборка документов раздела «Исполнительное производство».

Читает книгу iss-input.xlsx и собирает:
  • один файл Word — заявления об исполнении (по одному на должника,
    каждое с новой страницы);
  • сопроводительное письмо приставу — с таблицей должников;
  • реестр (registry.xlsx);
  • отчёт сверки (validation-report-iss.txt).
"""
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

import iss_lib as il

BASE = Path(__file__).resolve().parent
BOOK = BASE / il.BOOK
REPORT = BASE / il.REPORT
OUT_DIR = BASE / 'output-iss'


def kind_of(row):
    """Вид заявления по строке: «сп» или «ил»."""
    kind = str(row.get('вид') or '').strip().lower()
    if kind.startswith('исполн') or kind == 'ил':
        return 'ил'
    if kind.startswith('суд') or kind == 'сп':
        return 'сп'
    return ''


def main(argv):
    book = Path(argv[1]) if len(argv) > 1 else BOOK
    if not book.is_file():
        print('Нет книги %s — сначала загрузите данные.' % book.name)
        return 2
    settings, docs = il.read_book(book)
    if not docs:
        print('В книге нет строк — загрузите данные.')
        return 2

    warnings = []
    items = []
    for n, row in enumerate(docs, 1):
        kind = kind_of(row)
        if not kind:
            kind = 'ил' if row.get('номер') else 'сп'
        tpl = il.pick_template(kind, settings)
        if not tpl or not tpl.is_file():
            print('Не найден шаблон заявления (%s). Положите .docx в папку %s'
                  % (kind, il.TEMPLATE_DIR.name))
            return 2
        if not str(row.get('фио') or '').strip():
            warnings.append('строка %d: нет ФИО — заявление пропущено' % n)
            continue
        if not str(row.get('адрес') or '').strip():
            warnings.append('%s: нет адреса' % row.get('фио'))
        if row.get('сумма') in ('', None):
            warnings.append('%s: не указана сумма долга' % row.get('фио'))
        items.append((tpl, row))

    if not items:
        print('Нет строк, из которых можно собрать заявление.')
        return 2

    stamp = datetime.now().strftime('%Y-%m-%d_%H-%M')
    out_dir = OUT_DIR / stamp
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        slots, notes = il.build_document(items, settings,
                                         out_dir / 'Заявления об исполнении.docx')
    except Exception as ex:  # noqa: BLE001
        print('Не удалось собрать заявления: %s' % ex)
        return 1
    warnings.extend(notes)

    # сопроводительное письмо приставу
    acc_path = None
    acc_tpl = il.pick_template('сопровод', settings)
    if acc_tpl and acc_tpl.is_file():
        try:
            filled = il.build_accompanying(acc_tpl, [row for _t, row in items],
                                           out_dir / 'Сопровод приставу.docx')
            acc_path = out_dir / 'Сопровод приставу.docx'
            print('  сопровод: %d строк' % filled)
        except Exception as ex:  # noqa: BLE001
            warnings.append('сопровод приставу не собран: %s' % ex)

    # реестр
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Реестр'
    heads = ['№', 'Вид документа', '№ исполнительного документа',
             'Дата документа', 'ФИО должника', 'Адрес', 'Сумма долга',
             'Госпошлина', 'Период', 'Суд', 'Примечание']
    for i, h in enumerate(heads, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.fill = PatternFill('solid', fgColor='FFF2CC')
        c.font = Font(bold=True)
    total = 0.0
    for n, (tpl, row) in enumerate(items, 1):
        amount = row.get('сумма')
        try:
            total += float(str(amount).replace(' ', '').replace(',', '.'))
        except (TypeError, ValueError):
            pass
        vals = [n, row.get('вид') or kind_of(row), row.get('номер'),
                il.date_plain(row.get('дата')) if row.get('дата') else '',
                row.get('фио'), row.get('адрес'), amount,
                row.get('госпошлина'), row.get('период'), row.get('суд'),
                row.get('примечание')]
        for i, v in enumerate(vals, 1):
            ws.cell(row=n + 1, column=i, value=v)
    ws.cell(row=len(items) + 2, column=6, value='ИТОГО').font = Font(bold=True)
    ws.cell(row=len(items) + 2, column=7,
            value=il.money_words(total)).font = Font(bold=True)
    wb.save(out_dir / 'registry.xlsx')

    # отчёт
    lines = ['СБОРКА — ИСПОЛНИТЕЛЬНОЕ ПРОИЗВОДСТВО', '=' * 44, '']
    lines.append('Книга: %s' % book.name)
    lines.append('Заявлений: %d' % len(items))
    lines.append('Итого долга: %s руб.' % il.money_words(total))
    lines.append('Файл заявлений: Заявления об исполнении.docx')
    if acc_path:
        lines.append('Сопровод приставу: Сопровод приставу.docx')
    lines.append('Реестр: registry.xlsx')
    lines.append('')
    lines.append('ШАБЛОНЫ И НАЙДЕННЫЕ МЕСТА:')
    for path, sl in slots.items():
        lines.append('  %s:' % Path(path).name)
        for name, para in il.describe_slots(sl):
            lines.append('    • %s — абзац %d' % (name, para))
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

    print('СБОРКА — ИСПОЛНИТЕЛЬНОЕ ПРОИЗВОДСТВО')
    print('  заявлений: %d, итого долга: %s руб.'
          % (len(items), il.money_words(total)))
    print('  все заявления — в одном файле, каждое с новой страницы')
    if acc_path:
        print('  сопровод приставу — отдельным файлом')
    if warnings:
        print('  предупреждений: %d (см. %s)' % (len(warnings), il.REPORT))
    print('  папка: %s' % out_dir.name)
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
