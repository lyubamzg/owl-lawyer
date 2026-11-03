#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Семейное право: анкета «family-questionnaire.xlsx» → исковое заявление .docx.

Один файл анкеты = одно дело = один иск с тремя требованиями:
  1) расторжение брака, 2) взыскание алиментов, 3) место жительства детей.
Шапка «Истец/Ответчик» заполняется из анкеты при генерации (в образце пусто).

Запуск: python3 generate_family_claim.py [Анкета.xlsx]
        python3 generate_family_claim.py --демо   (демо-анкета + демо-иск)
Выход: output-family/<штамп>/… .docx + validation-report-family.txt
"""
import re
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
import platform_utils

from generate import (add_par, sum_words, fmt_date, parse_date, parse_sum,
                      cell_str, fio_short, fio_nom_short, FIO_RE)

from docx import Document
from docx.shared import Pt, Cm

BASE = Path(__file__).resolve().parent
ANKETA = BASE / 'family-questionnaire.xlsx'
OUT_BASE = BASE / 'output-family'
REPORT_NAME = 'validation-report-family.txt'
POSHLINA_DEFAULT = 5000  # пп. 5 п. 1 ст. 333.19 НК РФ (с 09.2024)

# По ст. 81 СК РФ: на одного 1/4, на двух 1/3, на трёх и более 1/2
ALIM_FRAC = {1: '1/4', 2: '1/3', 3: '1/2'}
ALIM_COL = 'Доля алиментов (1/4, 1/3, 1/2 или твердая сумма)'


# ---------------------------------------------------------------- чтение ----
def read_sheet(ws):
    heads = {str(c.value).strip(): c.column for c in ws[1] if c.value}
    if not heads:
        return []
    rows = []
    for r in range(2, ws.max_row + 1):
        row = {k: cell_str(ws.cell(row=r, column=col).value)
               for k, col in heads.items()}
        if any(v for v in row.values()):
            rows.append(row)
    return rows


def load_anketa(path: Path):
    """Читает анкету → (дело, дети)."""
    if not path.exists():
        sys.exit(f'Анкета не найдена: {path}\nСоздайте её через раздел «Семейное '
                 f'право» → «Загрузить сканы документов» или флагом --демо.')
    wb = openpyxl.load_workbook(path)
    for need in ('Дело', 'Дети'):
        if need not in wb.sheetnames:
            sys.exit(f'В анкете нет листа «{need}» — файл повреждён или чужой.')
    dela = read_sheet(wb['Дело'])
    deti = read_sheet(wb['Дети'])
    if len(dela) > 1:
        print(f'ВНИМАНИЕ: на листе «Дело» больше одной строки ({len(dela)}) — '
              f'беру первую. Для второго дела скопируйте файл анкеты.')
    return (dela[0] if dela else {}), deti


# ------------------------------------------------------------ алименты ----
def plural_deti(n):
    if n % 100 in range(11, 20) or n % 10 in (0, 5, 6, 7, 8, 9):
        return 'детей'
    return 'ребёнка' if n % 10 == 1 else 'детей'


def parse_alim(text, n):
    """«1/3» → {'kind':'frac','frac':'1/3'}; «5000» → {'kind':'sum','x':5000.0}."""
    s = str(text).strip()
    m = re.fullmatch(r'(\d)\s*/\s*(\d)', s)
    if m:
        return {'kind': 'frac', 'frac': f'{m.group(1)}/{m.group(2)}'}
    num = parse_sum(s) if s else None
    if num:
        return {'kind': 'sum', 'x': num}
    return {'kind': 'frac', 'frac': ALIM_FRAC.get(n, '1/3')}


# ------------------------------------------------------------- валидация ----
def validate_case(d, deti, errors, warnings):
    """Проверяет анкету → dict-дело для сборки (c['n'] и т.д.)."""
    c = {}
    if not d:
        errors.append('Лист «Дело» пуст — заполните его.')
        return c

    c['fio'] = d.get('Истец ФИО', '')
    if not c['fio']:
        errors.append('Не указана ФИО истца (лист «Дело»).')
    elif not FIO_RE.match(c['fio']):
        warnings.append(f'ФИО истца выглядит неполным (нужно «Фамилия Имя Отчество»): '
                        f'«{c["fio"]}» — продолжаю, проверьте результат.')
    c['is_fem'] = d.get('Истец пол (ж/м)', '').strip().lower() in ('ж', 'жен', 'женский')

    c['dr'] = parse_date(d.get('Истец дата рождения', ''))
    if not c['dr']:
        errors.append('Не указана дата рождения истца (формат ДД.ММ.ГГГГ).')

    c['addr'] = d.get('Истец адрес регистрации', '')
    c['pasp'] = d.get('Паспорт истца (кем и когда выдан)', '')
    c['tel'] = d.get('Истец телефон', '')

    sud = [s.strip() for s in d.get('Суд (шапка)', '').split('\n') if s.strip()]
    c['sud'] = sud
    if not sud:
        errors.append('Не указан суд (лист «Дело», колонка «Суд (шапка)») — '
                      'он вписывается вручную, в две строки через Alt+Enter.')
    if sud and any('миров' in s.lower() for s in sud):
        warnings.append('В шапке суда указан мирОвой судья: иск со спором о детях '
                        '(место жительства) подсуден РАЙОННОМУ суду (ч. 3 ст. 23 ГПК РФ) — проверьте.')

    c['poshlina'] = parse_sum(d.get('Госпошлина, руб.', ''))
    if c['poshlina'] is None:
        c['poshlina'] = POSHLINA_DEFAULT
        warnings.append(f'Госпошлина не указана — подставлено {POSHLINA_DEFAULT} ₽ '
                        f'(пп. 5 п. 1 ст. 333.19 НК РФ). Сверьте по платёжке.')

    c['ot_fio'] = d.get('Ответчик ФИО', '')
    if not c['ot_fio']:
        errors.append('Не указана ФИО ответчика — она вписывается вручную '
                      '(распознаётся только паспорт истца).')
    elif not FIO_RE.match(c['ot_fio']):
        warnings.append(f'ФИО ответчика выглядит неполным: «{c["ot_fio"]}».')
    c['ot_dr'] = parse_date(d.get('Дата рождения ответчика', ''))
    c['ot_addr'] = d.get('Адрес ответчика', '')
    c['ot_pasp'] = d.get('Паспорт ответчика (кем и когда выдан)', '')
    if c['ot_fio']:
        oth = [w for w, v in (('адрес', c['ot_addr']), ('паспортные данные', c['ot_pasp']))
               if not v]
        if oth:
            warnings.append('У ответчика не указаны ' + ' и '.join(oth) +
                            ' — в иске будет «данные неизвестны», впишите позже вручную.')

    c['seria'] = d.get('Свидетельство о браке: серия', '')
    c['act_num'] = d.get('Номер записи акта о браке', '')
    c['brak_date'] = parse_date(d.get('Дата заключения брака', ''))
    c['zags'] = d.get('Наименование ЗАГС', '')
    brak_missing = [k for k, v in (('серия свидетельства', c['seria']),
                                   ('номер актовой записи', c['act_num']),
                                   ('дата заключения брака', fmt_date(c['brak_date'])),
                                   ('наименование ЗАГС', c['zags'])) if not v]
    if brak_missing:
        errors.append('Не хватает данных о браке: ' + ', '.join(brak_missing) +
                      '. Загрузите скан свидетельства или впишите вручную.')

    c['stop_date'] = parse_date(d.get('Дата прекращения совместной жизни', ''))
    if not c['stop_date']:
        warnings.append('«Дата прекращения совместной жизни» пуста — в иске будет '
                        'формулировка без даты.')

    c['detya'] = []
    nd = 0
    for k in deti:
        if not k.get('ФИО ребёнка'):
            continue
        nd += 1
        dr = parse_date(k.get('Дата рождения', ''))
        if not dr:
            errors.append(f'Ребёнок {nd}: нет даты рождения.')
        if not FIO_RE.match(k['ФИО ребёнка']):
            warnings.append(f'Ребёнок {nd}: «{k["ФИО ребёнка"]}» — не полное ФИО в именительном '
                            f'падеже. На бланке имя пишется датively («Родившейся Ивановой '
                            f'Полине…») — впишите «Иванова Полина Ириковна».')
        if not k.get('Свид-во о рождении (серия, номер записи)'):
            warnings.append(f'Ребёнок {nd}: не указано свидетельство о рождении.')
        c['detya'].append({'fio': k['ФИО ребёнка'], 'dr': dr,
                           'sv': k.get('Свид-во о рождении (серия, номер записи)', '')})
    if nd == 0:
        errors.append('На листе «Дети» нет строк с ФИО — иск о детях требует хотя бы '
                      'одного ребёнка.')
    c['n'] = nd

    raw_alim = d.get(ALIM_COL, '')
    c['alim'] = parse_alim(raw_alim, nd)
    if not raw_alim:
        warnings.append(f'Доля алиментов не указана — взято по ст. 81 СК РФ: '
                        f'{c["alim"]["frac"]} на {nd} {plural_deti(nd)}. Сверьте.')

    with_whom = (d.get('Дети проживают с (истец/ответчик)', '') or 'истец').lower()
    if with_whom != 'истец':
        warnings.append(f'«Дети проживают с» = {with_whom!r}, а иск просит определить '
                        f'место жительства С ИСТЦОМ — сверьте.')
    return c


# ------------------------------------------------------------ документ ----
def new_doc():
    doc = Document()
    st = doc.styles['Normal']
    st.font.name = 'Times New Roman'
    st.font.size = Pt(11)
    st.paragraph_format.space_after = Pt(0)
    st.paragraph_format.line_spacing = 1.0
    for s in doc.sections:
        s.left_margin, s.right_margin = Cm(3), Cm(1.5)
        s.top_margin, s.bottom_margin = Cm(2), Cm(2)
    return doc


def person_family_line(fio, dr, addr, pasp, fem):
    """«ФИО, дд.мм.гггг г.р., зарегистрирован(а): … , паспорт …» (правый край)."""
    bits = [f'{fmt_date(dr)} г.р.']
    bits.append(f'зарегистрирован(а): {addr}' if addr
                else 'адрес регистрации неизвестен')
    bits.append(f'паспорт: {pasp}' if pasp else 'паспортные данные неизвестны')
    return f'{fio}, ' + ', '.join(bits)


def on_kogo(c):
    if c['n'] == 1:
        k = c['detya'][0]
        return f'несовершеннолетнего ребёнка {k["fio"]}, {fmt_date(k["dr"])} г.р.'
    kids = '; '.join(f'{k["fio"]}, {fmt_date(k["dr"])} г.р.' for k in c['detya'])
    return f'несовершеннолетних детей: {kids}'


def alim_text(c, num):
    ot = fio_short(c['ot_fio'], fem=False) if FIO_RE.match(c['ot_fio']) else c['ot_fio']
    a = c['alim']
    do_sov = 'до достижения ребёнком совершеннолетия.' if c['n'] == 1 \
        else 'до достижения детьми совершеннолетия.'
    if a['kind'] == 'frac':
        razmer = f'в размере {a["frac"]} всех видов заработка и иного дохода ежемесячно'
    else:
        razmer = f'в твёрдой денежной сумме в размере {sum_words(a["x"])} ежемесячно'
    return (f'{num}. Взыскать с {ot} алименты на содержание {on_kogo(c)} {razmer}, '
            f'начиная со дня подачи настоящего искового заявления и {do_sov}')


LEGAL_FAMILY = [
    'В соответствии со ст. 21 СК РФ расторжение брака при наличии у супругов общих '
    'несовершеннолетних детей производится в судебном порядке.',
    'В соответствии со ст. 24 СК РФ при расторжении брака в судебном порядке суд '
    'принимает меры по определению места жительства детей и взыскивает алименты на '
    'содержание детей, если между супругами не достигнуто соглашение по этим вопросам.',
    'В силу п. 3 ст. 65 СК РФ и п. 2 ст. 66 СК РФ место жительства детей при раздельном '
    'проживании родителей устанавливается соглашением родителей, а при отсутствии '
    'такого соглашения — судом, исходя из интересов детей.',
    'Согласно ст. 80 СК РФ родители обязаны содержать своих несовершеннолетних детей. '
    'В соответствии со ст. 81 СК РФ размер алиментов составляет: на одного ребёнка — '
    '1/4, на двух — 1/3, на трёх и более детей — 1/2 всех видов заработка и иного '
    'дохода родителей.',
    'Спор о детях входит в компетенцию районного суда (ч. 3 ст. 23 ГПК РФ). Настоящее '
    'исковое заявление подаётся с соблюдением требований ст. 131, 132 ГПК РФ. '
    'Государственная пошлина уплачена согласно пп. 5 п. 1 ст. 333.19 НК РФ; при подаче '
    'иска о взыскании алиментов истец освобождён от уплаты государственной пошлины '
    '(ст. 333.36 НК РФ).',
]


def build_docx_familnoe(c, out_path):
    doc = new_doc()

    # шапка правым краем
    for line in c['sud']:
        add_par(doc, line, 'right', indent=False)
    add_par(doc, '------------------------------------------------', 'right', indent=False)
    add_par(doc, f'Истец: {person_family_line(c["fio"], c["dr"], c["addr"], c["pasp"], c["is_fem"])}',
            'right', indent=False)
    if c['tel']:
        add_par(doc, f'Тел. {c["tel"]}', 'right', indent=False)
    add_par(doc, f'Ответчик: {person_family_line(c["ot_fio"], c["ot_dr"], c["ot_addr"], c["ot_pasp"], False)}',
            'right', indent=False)
    add_par(doc, f'Государственная пошлина: {sum_words(c["poshlina"])}', 'right', indent=False)
    add_par(doc, '', 'center', indent=False)

    title = ('ИСКОВОЕ ЗАЯВЛЕНИЕ о расторжении брака, взыскании алиментов и '
             'определении места жительства ' + ('ребёнка' if c['n'] == 1 else 'детей'))
    add_par(doc, title, 'center', indent=False, bold=True)
    add_par(doc, '', 'center', indent=False)

    add_par(doc, f'Брак между мною, {c["fio"]}, и {c["ot_fio"]} заключён '
                 f'{fmt_date(c["brak_date"])} года в {c["zags"]}, о чём составлена запись '
                 f'акта о заключении брака № {c["act_num"]}. '
                 f'Свидетельство о заключении брака: серия {c["seria"]}.')

    kids_list = '; '.join(f'{k["fio"]} ({fmt_date(k["dr"])} г.р.)' for k in c['detya'])
    add_par(doc, f'От брака имеем '
                 f'{"несовершеннолетнего ребёнка" if c["n"] == 1 else "несовершеннолетних детей"}: '
                 f'{kids_list}. Свидетельства о рождении: '
                 + '; '.join(k['sv'] for k in c['detya']) + '.')

    if c['stop_date']:
        add_par(doc, f'Дальнейшая совместная жизнь и сохранение семьи невозможны. '
                     f'Семейные отношения фактически прекращены с {fmt_date(c["stop_date"])} '
                     f'года, совместное хозяйство не ведётся, примирение нецелесообразно.')
    else:
        add_par(doc, 'Дальнейшая совместная жизнь и сохранение семьи невозможны, '
                     'примирение нецелесообразно.')
    add_par(doc, 'Дети проживают со мной. Ответчик в содержании детей участия не '
                 'принимает; соглашение об уплате алиментов между супругами не достигнуто.')
    add_par(doc, 'Соглашение о месте жительства детей между родителями не достигнуто: '
                 'ответчик на определение места жительства детей с истцом не согласен.')

    for par in LEGAL_FAMILY:
        add_par(doc, par)
    add_par(doc, '', 'center', indent=False)
    add_par(doc, 'ПРОШУ:', 'center', indent=False, bold=True)

    add_par(doc, f'1. Расторгнуть брак между {c["fio"]} и {c["ot_fio"]}, зарегистрированный '
                 f'{fmt_date(c["brak_date"])} года ({c["zags"]}, актовая запись № {c["act_num"]}).')
    add_par(doc, alim_text(c, 2))
    nabor = (c['detya'][0]['sv'] if c['n'] == 1
             else '; '.join(f'{i + 1}) {k["fio"]}: {k["sv"]}' for i, k in enumerate(c['detya'])))
    add_par(doc, f'3. Определить место жительства {"ребёнка" if c["n"] == 1 else "детей"} '
                 f'({nabor}) с истцом {c["fio"]}.')
    add_par(doc, '', 'center', indent=False)

    add_par(doc, 'Приложения:', indent=False)
    apps = ['документ об уплате государственной пошлины',
            'копия искового заявления для ответчика',
            f'копия свидетельства о заключении брака ({c["seria"]})']
    apps += [f'копия свидетельства о рождении: {k["sv"]}' for k in c['detya']]
    if c['alim']['kind'] == 'sum':
        apps.append('расчёт взыскиваемых алиментов')
    if c['addr']:
        apps.append('выписка из домовой книги (справка о регистрации)')
    apps.append('копия паспорта истца')
    for i, item in enumerate(apps, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(1.0)
        p.paragraph_format.space_after = Pt(0)
        p.add_run(f'{i}. {item}')

    add_par(doc, '', 'center', indent=False)
    add_par(doc, f'Истец _________________ / {c["fio"]}', 'left', indent=False)
    add_par(doc, '«___» _______________ 20___ г.', 'left', indent=False)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)


# ------------------------------------------------------------------------ main ----
def main():
    platform_utils.fix_console()
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    flags = {a for a in sys.argv[1:] if a.startswith('--')}
    demo = '--демо' in flags or '--demo' in flags

    if demo:
        import import_family
        book = BASE / 'demo-family-questionnaire.xlsx'
        import_family.ensure_anketa(book, demo=True)
        print(f'Демо-анкета: {book.name} (вымышленные данные)')
    else:
        book = Path(args[0]) if args else ANKETA

    delo, deti = load_anketa(book)
    errors, warnings = [], []
    warnings.append('Формулировки — ЧЕРНОВИК: склонения и юридические обороты '
                    'обязательно проверьте перед подачей.')
    c = validate_case(delo, deti, errors, warnings)

    out_dir = OUT_BASE / datetime.now().strftime('%Y-%m-%d_%H-%M')
    out_path = None
    if not errors and c.get('fio'):
        try:
            out_path = out_dir / platform_utils.ascii_filename(
                f'family-claim-{fio_nom_short(c["fio"])}.docx')
            build_docx_familnoe(c, out_path)
        except Exception as ex:  # noqa: BLE001
            errors.append(f'Не удалось сформировать документ — {ex}')
            out_path = None

    report = [f'ОТЧЁТ-СВЕРКИ (семейные)', f'Анкета: {book.name}', '']
    if errors:
        report.append('ОШИБКИ (иск не сформирован):')
        report += [f'  ✗ {e}' for e in errors]
        report.append('')
    if warnings:
        report.append('ПРЕДУПРЕЖДЕНИЯ (проверьте вручную):')
        report += [f'  ⚠ {w}' for w in warnings]
        report.append('')
    if out_path:
        report.append(f'Иск сформирован: {out_path.relative_to(BASE)}')
        report.append('')
        report.append('ПЕРЕД ПОДАЧЕЙ: сверьте данные анкеты с оригиналами сканов')
        report.append('(лист «Сканы» показывает, что извлек OCR), проверьте падежи')
        report.append('в «ПРОШУ» — и подсудность: иск идёт в РАЙОННЫЙ суд (ст. 23 ГПК РФ).')

    text = '\n'.join(report)
    report_path = (out_dir if not errors else BASE) / REPORT_NAME
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(text, encoding='utf-8')
    print(text)
    print()
    print(f'Отчёт: {report_path.relative_to(BASE)}')


if __name__ == '__main__':
    main()
