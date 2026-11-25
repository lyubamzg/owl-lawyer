#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Создаёт книгу раздела «Досудебные претензии»: pretension-input.xlsx.

Листы:
  «Инструкция» — как пользоваться разделом;
  «Настройки»   — шапка (реквизиты взыскателя), номер и дата претензии,
                  дата состояния, период, срок оплаты, имя шаблона;
  «Претензии»   — сам список: по одной строке на претензию.
"""
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import pretension_lib as pl

BASE = Path(__file__).resolve().parent
OUT = BASE / pl.BOOK

HEAD_FILL = PatternFill('solid', fgColor='FFF2CC')
HEAD_FONT = Font(bold=True, size=11)


def build(out_path=OUT):
    wb = openpyxl.Workbook()

    ins = wb.active
    ins.title = 'Инструкция'
    ins.column_dimensions['A'].width = 110
    for i, line in enumerate([
        'РАЗДЕЛ «ДОСУДЕБНЫЕ ПРЕТЕНЗИИ» — рабочая таблица',
        '',
        'Как работать:',
        '  1. «Загрузить список» — выберите файл Excel со списком должников',
        '     (столбцы: лицевой счёт, ФИО, адрес, сумма; шапки может не быть).',
        '  2. «Открыть таблицу в редакторе» — проверьте и поправьте:',
        '        • список претензий (лист «Претензии»);',
        '        • шапку — реквизиты взыскателя (лист «Настройки»);',
        '        • исходный номер претензии — «претензия_номер_от»;',
        '        • дату претензии — «дата_претензии»;',
        '        • дату, на которую считается долг, — «дата_состояния»;',
        '        • период задолженности — «период».',
        '  3. «Собрать документы» — претензии соберутся в один файл Word,',
        '     каждая с новой страницы.',
        '',
        'В листе «Претензии» можно задать свои значения по строке:',
        '  «Дата состояния», «Период», «№ претензии», «Дата претензии».',
        'Если ячейка пустая — берётся значение из листа «Настройки».',
    ], 1):
        ins.cell(row=i, column=1, value=line)
    ins['A1'].font = Font(bold=True, size=13)

    st = wb.create_sheet('Настройки')
    st['A1'], st['B1'] = 'Параметр', 'Значение'
    for cell in (st['A1'], st['B1']):
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
    for i, (key, val) in enumerate(pl.DEFAULT_SETTINGS, 2):
        st.cell(row=i, column=1, value=key)
        st.cell(row=i, column=2, value=val)
    st.column_dimensions['A'].width = 26
    st.column_dimensions['B'].width = 58
    st.freeze_panes = 'A2'

    # значения шапки, считанные из шаблона: по ним видно, что юрист правил руками
    tm = wb.create_sheet('Шаблон')
    tm['A1'], tm['B1'] = 'Параметр', 'Значение из шаблона'
    for cell in (tm['A1'], tm['B1']):
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
    tm.column_dimensions['A'].width = 26
    tm.column_dimensions['B'].width = 58
    tm.sheet_state = 'hidden'

    cl = wb.create_sheet('Претензии')
    for i, name in enumerate(pl.CLAIM_HEADERS, 1):
        cell = cl.cell(row=1, column=i, value=name)
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cl.column_dimensions[get_column_letter(i)].width = \
            {1: 5, 2: 14, 3: 26, 4: 38, 5: 13, 6: 15, 7: 24, 8: 12, 9: 15, 10: 20}[i]
    cl.freeze_panes = 'A2'

    out_path = Path(out_path)
    wb.save(out_path)
    return out_path


def main():
    out = build()
    print('Создано: %s' % out)
    return 0


if __name__ == '__main__':
    sys.exit(main())
