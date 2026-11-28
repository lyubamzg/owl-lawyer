#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Создаёт книгу раздела «Исполнительное производство»: iss-input.xlsx.

Листы:
  «Инструкция» — как пользоваться разделом;
  «Настройки»   — вид заявлений, шаблоны, дата и номер заявления;
  «Заявления»   — сам список: по одной строке на должника.
"""
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import iss_lib as il

BASE = Path(__file__).resolve().parent
OUT = BASE / il.BOOK

HEAD_FILL = PatternFill('solid', fgColor='FFF2CC')
HEAD_FONT = Font(bold=True, size=11)

INSTRUCTIONS = [
    'РАЗДЕЛ «ИСПОЛНИТЕЛЬНОЕ ПРОИЗВОДСТВО» — рабочая таблица',
    '',
    'Как работать:',
    '  1. «Загрузить данные» — выберите:',
    '        • PDF со сканами судебных приказов — приложение распознает их',
    '          (OCR) и впишет должников, суммы, госпошлину и номер дела;',
    '        • готовый реестр исполнительных документов в ФССП (Excel);',
    '        • шаблоны заявлений (.docx) — если хотите свои.',
    '  2. «Открыть таблицу в редакторе» — проверьте и поправьте:',
    '        • список заявлений (лист «Заявления»);',
    '        • вид заявления — «судебный приказ» или «исполнительный лист»;',
    '        • шаблоны — имена файлов из папки «iss-templates»;',
    '        • дату и номер заявления — «дата_заявления», «номер_заявления».',
    '  3. «Собрать документы» — заявления соберутся в один файл Word',
    '     (каждое с новой страницы), рядом — реестр и сопровод приставу.',
    '',
    'В листе «Заявления» одна строка — одно заявление (один должник).',
    'ФИО пишется в именительном падеже («Иванов Иван Иванович»).',
    '',
    'Если ФИО распозналось в родительном падеже (как в судебном приказе —',
    '«с должников — Ивановой Марии Ивановны»), приложение переводит его',
    'в именительный. Проверьте результат по таблице.',
]


def refresh_instructions(wb):
    """Обновить лист «Инструкция» в уже созданной книге."""
    if 'Инструкция' not in wb.sheetnames:
        return
    ins = wb['Инструкция']
    for i, line in enumerate(INSTRUCTIONS, 1):
        ins.cell(row=i, column=1, value=line)
    for r in range(len(INSTRUCTIONS) + 1, ins.max_row + 1):
        ins.cell(row=r, column=1, value=None)
    ins['A1'].font = Font(bold=True, size=13)


def build(out_path=OUT):
    wb = openpyxl.Workbook()

    ins = wb.active
    ins.title = 'Инструкция'
    ins.column_dimensions['A'].width = 100
    for i, line in enumerate(INSTRUCTIONS, 1):
        ins.cell(row=i, column=1, value=line)
    ins['A1'].font = Font(bold=True, size=13)

    st = wb.create_sheet('Настройки')
    st['A1'], st['B1'] = 'Параметр', 'Значение'
    for cell in (st['A1'], st['B1']):
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
    for i, (key, val) in enumerate(il.DEFAULT_SETTINGS, 2):
        st.cell(row=i, column=1, value=key)
        st.cell(row=i, column=2, value=val)
    st.column_dimensions['A'].width = 26
    st.column_dimensions['B'].width = 58
    st.freeze_panes = 'A2'

    cl = wb.create_sheet('Заявления')
    for i, name in enumerate(il.DOC_HEADERS, 1):
        cell = cl.cell(row=1, column=i, value=name)
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cl.column_dimensions[get_column_letter(i)].width = {
            1: 5, 2: 22, 3: 30, 4: 14, 5: 34, 6: 44,
            7: 14, 8: 14, 9: 26, 10: 40, 11: 26}[i]
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
