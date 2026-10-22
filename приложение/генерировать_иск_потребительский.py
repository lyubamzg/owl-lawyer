#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Защита прав потребителя: анкета «Анкета-потребительские.xlsx» → иск .docx.

Один файл анкеты = одно дело = один иск. ОПИСАТЕЛЬНАЯ ЧАСТЬ (проблематика)
берётся с листа «Проблематика» — свободный текст, каждая строка = абзац;
шаблон её не трогает. Фиксированные блоки: статьи ЗоЗПП/ГК/ГПК/НК и ПРОШУ:
товар → неустойка → моральный вред → штраф 50 %.

Запуск: python3 генерировать_иск_потребительский.py [Анкета.xlsx]
        python3 генерировать_иск_потребительский.py --демо
Выход: Выход-потребительские/<штамп>/… .docx + Отчёт-сверки (потребительские).txt
"""
import re
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
import платформа

from generate import add_par, sum_words, fmt_date, parse_date, parse_sum, cell_str, FIO_RE
from генерировать_иск import new_doc, person_family_line, fio_nom_short
from docx.shared import Pt, Cm

BASE = Path(__file__).resolve().parent
ANKETA = BASE / 'Анкета-потребительские.xlsx'
OUT_BASE = BASE / 'Выход-потребительские'
REPORT_NAME = 'Отчёт-сверки (потребительские).txt'

PRICE_LIMIT_POST = 1_000_000  # п. 3 ст. 333.36 НК РФ: освобождение потребителей до 1 млн ₽


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
    if not path.exists():
        sys.exit('Анкета не найдена: ' + str(path) +
                 '\nСоздайте её: раздел «Защита прав потребителя» → '
                 '«Загрузить сканы документов» или флагом --демо.')
    wb = openpyxl.load_workbook(path)
    for need in ('Дело', 'Проблематика'):
        if need not in wb.sheetnames:
            sys.exit(f'В анкете нет листа «{need}» — файл повреждён или чужой.')
    dela = read_sheet(wb['Дело'])
    if len(dela) > 1:
        print('ВНИМАНИЕ: на листе «Дело» больше одной строки — беру первую.')
    # проблематика: строки колонки A; каждая заполненная строка (или строка с Alt+Enter) — абзац
    paras = []
    ws = wb['Проблематика']
    for r in range(2, ws.max_row + 1):
        v = cell_str(ws.cell(row=r, column=1).value)
        if not v:
            continue
        if '\n' in v:
            paras += [s.strip() for s in v.split('\n') if s.strip()]
        else:
            paras.append(v)
    return (dela[0] if dela else {}), paras


# ------------------------------------------------------------- валидация ----
def validate_case(d, para, errors, warnings):
    c = {}
    if not d:
        errors.append('Лист «Дело» пуст — заполните его.')
        return c

    c['fio'] = d.get('Истец ФИО', '')
    if not c['fio']:
        errors.append('Не указана ФИО истца.')
    elif not FIO_RE.match(c['fio']):
        warnings.append(f'ФИО истца выглядит неполным («Фамилия Имя Отчество»): «{c["fio"]}».')

    # пол в этой анкете не спрашиваем — эвристика «ва/на/ая» (влияет только на «зарегистрирован(а)»)
    fam_word = c['fio'].split()[0].lower() if c['fio'] else ''
    c['is_fem'] = bool(fam_word) and (fam_word.endswith(('ва', 'на', 'ая', 'овна', 'евна'))
                                      or fam_word[-1] in 'аян')

    c['dr'] = parse_date(d.get('Истец дата рождения', ''))
    c['addr'] = d.get('Истец адрес регистрации', '')
    c['pasp'] = d.get('Паспорт истца (кем и когда выдан)', '')
    c['tel'] = d.get('Истец телефон', '')
    if not c['dr']:
        errors.append('Не указана дата рождения истца (ДД.ММ.ГГГГ) — она нужна в шапке.')
    if not c['addr']:
        warnings.append('Не указан адрес регистрации истца — в шапке будет «адрес неизвестен».')

    sud = [s.strip() for s in d.get('Суд (шапка)', '').split('\n') if s.strip()]
    c['sud'] = sud
    if not sud:
        errors.append('Не указан суд (колонка «Суд (шапка)») — заполняется вручную.')

    c['ot_name'] = d.get('Ответчик (наименование организации/ИП)', '')
    if not c['ot_name']:
        errors.append('Не указан ответчик (наименование организации/ИП) — вручную.')
    c['ot_addr'] = d.get('Адрес ответчика', '')
    c['ot_inn'] = d.get('ИНН/ОГРН ответчика', '')
    if not c['ot_addr']:
        warnings.append('Не указан адрес ответчика — в иске будет «адрес неизвестен».')

    c['price'] = parse_sum(d.get('Стоимость товара (услуги), руб.', ''))
    if c['price'] is None:
        errors.append('Не указана стоимость товара (услуги) — она нужна для ПРОШУ '
                      'и отметки о госпошлине.')
    c['buy_date'] = parse_date(d.get('Дата покупки (заказа)', ''))
    if not c['buy_date']:
        warnings.append('Не указана дата покупки — в описании из проблематики она '
                        'уже может быть; дата используется только в ПРОШУ.')

    c['claim'] = (d.get('Что просим (возврат денег/замена товара/устранение недостатков)', '')
                  or 'возврат денег').strip().lower()
    if 'возврат' not in c['claim'] and 'замена' not in c['claim'] and 'устран' not in c['claim']:
        warnings.append(f'«Что просим» не в списке (текущее: {c["claim"]!r}) — '
                        f'беру «возврат денег». Исправьте в анкете, если нужно иное.')
        c['claim'] = 'возврат денег'

    c['ney'] = parse_sum(d.get('Неустойка, руб.', ''))
    c['moral'] = parse_sum(d.get('Моральный вред, руб.', '')) or 0.0
    c['ubi'] = parse_sum(d.get('Прочие убытки, руб.', '')) or 0.0
    if c['ney'] is None:
        warnings.append('Неустойка не указана — пункт о неустойке будет пропущен.')
    if c['moral'] == 0:
        warnings.append('Моральный вред не указан — пункт о компенсации будет пропущен.')

    posh = parse_sum(d.get('Госпошлина, руб.', ''))
    if posh is None:
        if (c['price'] or 0) <= PRICE_LIMIT_POST:
            c['poshlina'] = 0.0
            c['posh_note'] = 'истец освобождён от уплаты (п. 3 ст. 333.36 НК РФ — потребитель, цена иска до 1 млн ₽)'
        else:
            warnings.append('Цена иска выше 1 000 000 ₽ — потребитель НЕ освобождён от '
                            'пошлины; впишите её в анкету.')
            c['poshlina'] = None
            c['posh_note'] = ''
    else:
        c['poshlina'] = posh
        c['posh_note'] = ''

    c['apps_free'] = [s.strip(' .;') for s in (d.get('Приложения (свой список, через «;»)', '')
                                               or '').split(';') if s.strip(' .;')]

    if not para:
        errors.append('Лист «Проблематика» пуст — описательную часть иска вы заполняете '
                      'сами (свободный текст, строка = абзац).')
    c['para'] = para
    return c


# ------------------------------------------------------------ документ ----
def ney_proshu(c, num):
    return (f'{num}. Взыскать с ответчика неустойку в размере {sum_words(c["ney"])} '
            f'за нарушение срока удовлетворения требования потребителя '
            f'(ст. 22, 23 Закона о защите прав потребителей).')


def proshu_1(c, num):
    cena = sum_words(c['price']) if c['price'] else '__ (впишите цену в анкету)'
    d = fmt_date(c['buy_date']) if c['buy_date'] else '__ (дата покупки)'
    if 'замена' in c['claim']:
        return (f'{num}. Обязать ответчика заменить товар на однородный товар '
                f'надлежащего качества, уплаченную сумму в размере {cena} '
                f'принять в зачёт (ст. 18, 21 Закона о защите прав потребителей).')
    if 'устран' in c['claim']:
        return (f'{num}. Обязать ответчика безвозмездно устранить недостатки товара '
                f'в возможно короткий срок (ст. 18, 20 Закона о защите прав потребителей).')
    return (f'{num}. Расторгнуть договор купли-продажи от {d} и взыскать с ответчика '
            f'уплаченную за товар сумму в размере {cena} '
            f'(ст. 18 Закона о защите прав потребителей).')


LEGAL_CONSUMER = [
    'Отношения между истцом и ответчиком регулируются Законом РФ «О защите прав '
    'потребителей» от 07.02.1992 № 2300-1 и Гражданским кодексом РФ.',
    'В соответствии со ст. 4 Закона о защите прав потребителей продавец обязан передать '
    'потребителю товар, качество которого соответствует договору.',
    'Согласно ст. 18 Закона о защите прав потребителей потребитель в случае обнаружения '
    'в товаре недостатков вправе отказаться от исполнения договора купли-продажи и '
    'потребовать возврата уплаченной за товар суммы.',
    'Согласно ст. 19 Закона потребитель вправе предъявить требования к качеству товара '
    'в течение гарантийного срока; требование о возврате денег подлежит удовлетворению '
    'в 10-дневный срок (ст. 22 Закона).',
    'За нарушение установленного срока потребитель вправе взыскать неустойку в размере '
    'одного процента цены товара за каждый день просрочки (ст. 23 Закона).',
    'В силу ст. 15 Закона моральный вред, причинённый потребителю вследствие нарушения '
    'прав, подлежит компенсации; согласно п. 6 ст. 13 Закона при удовлетворении '
    'требований потребителя суд взыскивает штраф в размере пятидесяти процентов '
    'присуждённой суммы.',
    'Претензия потребителя подлежит удовлетворению по месту нахождения истца либо '
    'ответчика; иск подаётся по выбору истца (ст. 17 Закона о защите прав потребителей, '
    'ч. 7 ст. 29 ГПК РФ) и оформлен с соблюдением ст. 131, 132 ГПК РФ; при цене иска '
    'до 1 000 000 ₽ истец-потребитель освобождён от уплаты государственной пошлины '
    '(п. 3 ст. 333.36 НК РФ).',
]


def build_docx_potr(c, out_path):
    doc = new_doc()

    for line in c['sud']:
        add_par(doc, line, 'right', indent=False)
    add_par(doc, '------------------------------------------------', 'right', indent=False)
    add_par(doc, f'Истец: {person_family_line(c["fio"], c["dr"], c["addr"], c["pasp"], True)}',
            'right', indent=False)
    if c['tel']:
        add_par(doc, f'Тел. {c["tel"]}', 'right', indent=False)

    ot = c['ot_name']
    add_par(doc, f'Ответчик: {ot}', 'right', indent=False)
    if c['ot_addr']:
        add_par(doc, f'Адрес: {c["ot_addr"]}', 'right', indent=False)
    if c['ot_inn']:
        add_par(doc, f'ИНН/ОГРН: {c["ot_inn"]}', 'right', indent=False)

    if c['poshlina']:
        add_par(doc, f'Государственная пошлина: {sum_words(c["poshlina"])}', 'right', indent=False)
    elif c.get('posh_note'):
        add_par(doc, f'Государственная пошлина: не уплачивается ({c["posh_note"]})',
                'right', indent=False)
    else:
        add_par(doc, 'Государственная пошлина: __________ (впишите)', 'right', indent=False)
    add_par(doc, '', 'center', indent=False)

    add_par(doc, 'ИСКОВОЕ ЗАЯВЛЕНИЕ о защите прав потребителя', 'center', indent=False, bold=True)
    add_par(doc, '', 'center', indent=False)

    add_par(doc, 'Я, ' + c['fio'] + ', приобрёл(а) у ответчика товар (услугу) с недостатками. '
                 'Изложение обстоятельств дела:')
    for t in c['para']:
        add_par(doc, t)

    for par in LEGAL_CONSUMER:
        add_par(doc, par)
    add_par(doc, '', 'center', indent=False)
    add_par(doc, 'ПРОШУ:', 'center', indent=False, bold=True)

    add_par(doc, proshu_1(c, 1))
    if c['ney']:
        add_par(doc, ney_proshu(c, 2))
    if c['moral']:
        add_par(doc, f'3. Взыскать с ответчика компенсацию морального вреда '
                     f'в размере {sum_words(c["moral"])} (ст. 15 Закона о защите прав потребителей).')
    add_par(doc, f'4. Взыскать с ответчика штраф в размере пятидесяти процентов '
                 f'присуждённой в пользу истца суммы (п. 6 ст. 13 Закона о защите '
                 f'прав потребителей).')
    add_par(doc, '', 'center', indent=False)

    add_par(doc, 'Приложения:', indent=False)
    apps = ['копия искового заявления для ответчика',
            'документ об уплате государственной пошлины' if c['poshlina']
            else 'обоснование освобождения от госпошлины (п. 3 ст. 333.36 НК РФ)',
            'копия чека / документа об оплате',
            'копия претензии и доказательства её отправки']
    dup_words = ('чек', 'претензи')   # база уже содержит чек и претензию — свой
    apps += [item for item in c['apps_free']]   # список тоже бывает дублем
    apps = apps[:4] + [item for item in apps[4:]
                       if not any(w in item.lower() for w in dup_words)]
    seen, uniq = set(), []   # страховочная точная дедупликация
    for item in apps:
        key = item.strip().lower()
        if key not in seen:
            seen.add(key)
            uniq.append(item)
    apps = uniq
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
    платформа.fix_console()
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    flags = {a for a in sys.argv[1:] if a.startswith('--')}
    demo = '--демо' in flags or '--demo' in flags

    if demo:
        import импорт_потребительских as imp
        book = BASE / 'Демо-Анкета-потребительские.xlsx'
        imp.ensure_anketa(book, demo=True)
        print('Демо-анкета: ' + book.name + ' (вымышленные данные)')
    else:
        book = Path(args[0]) if args else ANKETA

    delo, para = load_anketa(book)
    errors, warnings = [], []
    warnings.append('Формулировки — ЧЕРНОВИК: описательную часть и юридические обороты '
                    'проверьте перед подачей.')
    c = validate_case(delo, para, errors, warnings)

    out_dir = OUT_BASE / datetime.now().strftime('%Y-%m-%d_%H-%M')
    out_path = None
    if not errors and c.get('fio'):
        try:
            out_path = out_dir / f'ИСК {fio_nom_short(c["fio"])} (защита прав потребителя).docx'
            build_docx_potr(c, out_path)
        except Exception as ex:  # noqa: BLE001
            errors.append(f'Не удалось сформировать документ — {ex}')
            out_path = None

    report = ['ОТЧЁТ-СВЕРКИ (потребительские)', 'Анкета: ' + book.name, '']
    if errors:
        report += ['ОШИБКИ (иск не сформирован):'] + [f'  ✗ {e}' for e in errors] + ['']
    if warnings:
        report += ['ПРЕДУПРЕЖДЕНИЯ:'] + [f'  ⚠ {w}' for w in warnings] + ['']
    if out_path:
        report.append('Иск сформирован: ' + out_path.relative_to(BASE).__str__())
        report.append('')
        report.append('ПЕРЕД ПОДАЧЕЙ: описательная часть взята с листа «Проблематика» как есть —')
        report.append('прочитайте иск целиком и поправьте падежи/формулировки.')

    text = '\n'.join(report)
    report_path = (out_dir if not errors else BASE) / REPORT_NAME
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(text, encoding='utf-8')
    print(text)
    print()
    print('Отчёт: ' + str(report_path.relative_to(BASE)))


if __name__ == '__main__':
    main()
