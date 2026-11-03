#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OCR распознавание паспортных данных должников со сканов/фото.

Использование:
    python3 ocr_debtors.py [файлы или папка]

Складывает сканы (jpg, png, pdf, tif) в папку «scans» рядом со скриптом
(или передаёт их как аргументы). Файл лучше назвать с номером ЛС:
    «ЛС000001111 справка.jpg» — тогда номер подтянется в таблицу.

Результат: «recognized-debtors.xlsx» для проверки глазами
рядом со скриптом + тексты распознавания в output/ocr/.
Данные НИКУДА не отправляются — всё распознаётся на этом компьютере.
"""
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

import platform_utils

BASE = Path(__file__).resolve().parent
SRC_DEFAULT = BASE / 'scans'
OUT_DIR = BASE / 'output' / 'ocr'

IMAGE_EXT = {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.pdf'}

FIO_RE = re.compile(
    r'([А-ЯЁ][а-яё\-]+ +[А-ЯЁ][а-яё\-]+ +[А-ЯЁ][а-яё\-]+(?:ович|евич|овна|евна|ична|инична))\b')
DATE_RE = re.compile(r'\b(\d{2}[.,]\d{2}[.,](19|20)\d{2})\b')
PASPORT_RE = re.compile(r'\b(\d{2} ?\d{2})[\s№]*(\d{6})\b')
KEM_VIDAN_RE = re.compile(r'выдан[^\n]{10,200}', re.IGNORECASE)


def ocr_image(path: Path) -> str:
    """Распознаёт файл (pdf → страницы) и возвращает текст."""
    texts = []
    if path.suffix.lower() == '.pdf':
        stem = platform_utils.ascii_filename(path.stem)
        subprocess.run(['pdftoppm', '-r', '300', '-gray', '-png',
                        str(path), str(OUT_DIR / stem)], check=True)
        for img in sorted(OUT_DIR.glob(f'{stem}-*.png')):
            texts.append(run_tesseract(img))
    else:
        texts.append(run_tesseract(path))
    return '\n'.join(texts)


def run_tesseract(img: Path) -> str:
    out = OUT_DIR / platform_utils.ascii_filename(img.stem)
    subprocess.run(['tesseract', str(img), str(out), '-l', 'rus', '--psm', '4'],
                   check=True, capture_output=True)
    return (out.with_suffix('.txt')).read_text(encoding='utf-8')


def extract_persons(text: str):
    """Вытаскивает из текста кандидатов: ФИО + дата рождения + паспорт."""
    persons = []
    for m in FIO_RE.finditer(text):
        after = text[m.end():m.end() + 160]
        before = text[max(0, m.start() - 100):m.start()]
        # дата рождения — ближайшая после ФИО; отбрасываем даты выдачи паспорта сразу после «выдан»
        dr = None
        for dm in DATE_RE.finditer(after[:60]):
            dr = dm.group(1).replace(',', '.')
            if KEM_VIDAN_RE.search(text[m.end() + dm.start():m.end() + dm.start() + 30] or 'x'):
                dr = None
            else:
                break
        pasp = PASPORT_RE.search(before + ' ' + after)
        series_num = f'{pasp.group(1)} {pasp.group(2)}' if pasp else ''
        kem = ''
        km = KEM_VIDAN_RE.search(after)
        if km:
            kem = ' '.join(km.group(0).split()).rstrip('.,;')
        persons.append({'fio': m.group(1), 'dr': dr or '', 'pasport': series_num,
                        'kem': kem})
    return persons


def main():
    args = [Path(a) for a in sys.argv[1:]]
    if args:
        files = [f for a in args if (BASE / a).exists() or a.exists()
                 for f in ([BASE / a] if a.exists() else [a])]
    else:
        SRC_DEFAULT.mkdir(exist_ok=True)
        files = sorted(p for p in SRC_DEFAULT.iterdir()
                       if p.suffix.lower() in IMAGE_EXT and not p.name.startswith('~'))
    if not files:
        sys.exit('Нет файлов. Положите сканы в папку «scans» рядом со скриптом '
                 'или передайте пути: python3 ocr_debtors.py скан1.pdf скан2.png')

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'К проверке'
    heads = ['Лицевой счёт (из имени файла или руками)', 'ФИО полностью', 'Дата рождения',
             'Паспорт', 'Кем и когда выдан', 'Источник (файл)', '⚠ Проверьте по оригиналу']
    ws.append(heads)
    report = ['OCR-РАСПОЗНАВАНИЕ — отчёт']

    for f in files:
        f = Path(f)
        ls = re.search(r'ЛС\s*([0-9]{5,12})', f.stem, re.IGNORECASE)
        ls = ls.group(1) if ls else ''
        report.append(f'{f.name}:')
        try:
            text = ocr_image(f)
        except Exception as ex:  # noqa: BLE001
            report.append(f'  ! не удалось распознать: {ex}')
            continue
        (OUT_DIR / platform_utils.ascii_filename(f.stem + '.txt')).write_text(text, encoding='utf-8')
        persons = extract_persons(text)
        if not persons:
            report.append('  ! ФИО не найдены — смотри полный текст в output/ocr/')
        for p in persons:
            ws.append([ls, p['fio'], p['dr'], p['pasport'], p['kem'], f.name, ''])
            report.append(f'  {p["fio"]}, {p["dr"]}, {p["pasport"]} {p["kem"]}')

    for i, w in enumerate((16, 34, 14, 16, 44, 30, 10), 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill('solid', fgColor='DDEBF7')
        c.alignment = Alignment(wrap_text=True, vertical='top')

    out = BASE / 'recognized-debtors.xlsx'
    wb.save(out)
    print('\n'.join(report))
    print(f'\nТаблица для проверки: {out}')
    print('Проверьте каждую строку по оригиналу — OCR ошибается, '
          'особенно в цифрах паспорта!')


if __name__ == '__main__':
    platform_utils.fix_console()
    not_found = platform_utils.missing_tools()
    if not_found:
        print('Для OCR не хватает программ:\n  • '
              + '\n  • '.join(not_found)
              + '\n\nУстановите их по инструкции пакета Сова-юрист и повторите.')
        sys.exit(2)
    main()
