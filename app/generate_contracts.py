#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Конструктор договоров «Сова-юрист»: анкета «contracts-workbook.xlsx» → .docx.

Употребление:
  python3 generate_contracts.py [contracts-workbook.xlsx] [--демо]

Выход: «output-contracts/ГГГГ-ММ-ДД_ЧЧ-ММ/» — по одному .docx на договор,
при необходимости акт приёма-передачи + «validation-report-contracts.txt».
--демо строит демо-книгу (3 тестовых вымышленных договора) и собирает её.
"""
import datetime as dt
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))

import contract_rules as норма          # noqa: E402
import contract_texts as тексты        # noqa: E402
import platform_utils

OUT_ROOT = BASE / 'output-contracts'
DEFAULT_BOOK = BASE / 'contracts-workbook.xlsx'


# ───────────────────────────────── чтение анкеты ────────────────────────────

def _str(v):
    if v is None:
        return ''
    if hasattr(v, 'strftime'):
        return v.strftime('%d.%m.%Y')
    return str(v).strip()


def read_book(book_path):
    wb = load_workbook(book_path, data_only=True)
    ws = wb['Договоры']
    заголовки = [str(c.value).strip() if c.value else '' for c in ws[1]]
    договоры = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None and not (row and row[1]):
            continue
        d = {h: _str(v) for h, v in zip(заголовки, row)}
        d['№'] = строка_номер(d, len(договоры) + 1)
        договоры.append(d)

    перечень = {}
    if 'Перечень работ и услуг' in wb.sheetnames:
        pw = wb['Перечень работ и услуг']
        for row in pw.iter_rows(min_row=2, values_only=True):
            try:
                n = int(float(row[0])) if row[0] is not None else None
            except (ValueError, TypeError):
                continue
            if n is None or не_пуста(row[2]):
                if n is not None:
                    перечень.setdefault(n, []).append(_str(row[2]))

    график = {}
    if 'График платежей' in wb.sheetnames:
        gw = wb['График платежей']
        for row in gw.iter_rows(min_row=2, values_only=True):
            try:
                n = int(float(row[0])) if row[0] is not None else None
            except (ValueError, TypeError):
                continue
            if n is not None:
                график.setdefault(n, []).append(
                    '%s — %s (%s)' % (_str(row[1]), _str(row[2]), _str(row[3]) or 'платёж'))
    return договоры, перечень, график


def строка_номер(d, fallback):
    try:
        return int(float(str(d.get('№') or fallback)))
    except ValueError:
        return fallback


def не_пуста(v):
    return v is not None and str(v).strip() != ''


# ─────────────────────────── подстановка значений ───────────────────────────

def значения(d, перечень_lines, график_lines):
    номер = d['№']
    вид = d['Вид договора'].strip()
    роль_а, роль_б = норма.РОЛИ[вид]
    физ_а = норма.оба_граждане(d.get('Реквизиты стороны 1 (кто)'))
    физ_б = норма.оба_граждане(d.get('Реквизиты стороны 2 (кто)'))

    дата = норма.дата_текст(d.get('Дата заключения'))
    срок_с = норма.дата_текст(d.get('Срок действия (с)'))
    срок_по = норма.дата_текст(d.get('Срок действия (по)'))
    срок_работы = _str(d.get('Сроки выполнения работ (текст)'))
    if not срок_работы and (срок_с or срок_по):
        срок_работы = 'в период с %s по %s' % (срок_с or 'по настоящее время', срок_по)

    цена = _str(d.get('Цена (сумма), руб.')) or '0'
    оплата = _str(d.get('Порядок оплаты')) or \
        'оплата производится при подписании Договора'
    аванс = _str(d.get('Аванс/задаток, руб.')) or '0 (не уплачивается)'
    депозит = _str(d.get('Депозит (залог), руб.')) or '0 (не уплачивается)'
    коммуналка = _str(d.get('Коммунальные платежи (кто платит)')) or роль_б
    неустойка = _str(d.get('Неустойка (текст)')) or \
        '0 (штрафные санкции не установлены; применяются нормы ст. 395 ГК РФ)'
    особые = _str(d.get('Особые условия')) or 'особые условия не установлены'

    if вид == 'Заём':
        проценты_поле = _str(d.get('Проценты по займу (% год)'))
        if проценты_поле and норма.не_проценты_указано(проценты_поле):
            проценты = 'сумма займа передаётся под %s процентов в год (%s)' % (
                проценты_поле, 'начисление по правилам ст. 809 ГК РФ')
        elif физ_а and физ_б and норма.сумма(цена) is not None and норма.сумма(цена) <= 100000:
            проценты = ('проценты по Договору не начисляются — заём '
                        'беспроцентный (п. 4 ст. 809 ГК РФ)')
        else:
            проценты = ('размер процентов не установлен: проценты '
                        'определяются ключевой ставкой Банка России '
                        '(п. 1 ст. 809 ГК РФ)')
        возврат = _str(d.get('Возврат займа (срок)')) or \
            'в течение тридцати (30) дней со дня предъявления требования ' \
            '(п. 1 ст. 810 ГК РФ)'
    else:
        проценты = ''
        возврат = ''

    if перечень_lines:
        список = '\n'.join('     %s) %s' % (i, line)
                           for i, line in enumerate(перечень_lines, 1))
    elif график_lines:
        список = '\n'.join('     %s) %s' % (i, line)
                           for i, line in enumerate(график_lines, 1))
    else:
        список = '     — (детализация не приведена; состав определяется предметом)'

    return {
        'номер': str(номер),
        'дата': дата,
        'место': _str(d.get('Место заключения')) or 'г. Пример',
        'роль_а': роль_а,
        'субъ_а': 'гражданин(ка) РФ' if физ_а else '',
        'имя_а': _str(d.get('Сторона 1 (кто)')),
        'реквизиты_а': _str(d.get('Реквизиты стороны 1 (кто)')) or 'реквизиты не заполнены',
        'роль_б': роль_б,
        'субъ_б': 'гражданин(ка) РФ' if физ_б else '',
        'имя_б': _str(d.get('Сторона 2 (кто)')),
        'реквизиты_б': _str(d.get('Реквизиты стороны 2 (кто)')) or 'реквизиты не заполнены',
        'предмет': _str(d.get('Предмет договора')),
        'цена': цена,
        'оплата': оплата,
        'аванс': аванс,
        'депозит': депозит,
        'коммуналка': коммуналка,
        'срок_с': срок_с or 'по настоящее время',
        'срок_по': срок_по or 'до полного исполнения обязательств',
        'срок_работы': срок_работы or 'уточняется сторонами',
        'проценты': проценты,
        'возврат': возврат,
        'неустойка': неустойка,
        'особые': особые,
        'перечень': список,
    }


def место_текст(место):
    """«г. Пример» без дублирования «г.» в значении от юриста."""
    место = (место or '').strip()
    return место if re.match(r'^(г\. ?|гор\. ?|с\. ?|пос\. ?|п\. ?|ст\. ?)', место) \
        else 'г. %s' % место


# ─────────────────────────────────── word ───────────────────────────────────

def new_doc():
    doc = Document()
    st = doc.styles['Normal']
    st.font.name = 'Times New Roman'
    st.font.size = Pt(11)
    st.paragraph_format.space_after = Pt(0)
    for s in doc.sections:
        s.left_margin, s.right_margin = Cm(2), Cm(1.5)
        s.top_margin, s.bottom_margin = Cm(2), Cm(2)
    return doc


_ALIGN = {'center': WD_ALIGN_PARAGRAPH.CENTER,
          'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
          'left': WD_ALIGN_PARAGRAPH.LEFT,
          'right': WD_ALIGN_PARAGRAPH.RIGHT}


def add_par(doc, text, align='justify', indent=True, bold=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(2)
    if indent:
        pf.first_line_indent = Cm(1.25)
    p.alignment = _ALIGN[align]
    r = p.add_run(text)
    r.bold = bold
    return p


def подпись_block(знач, doc):
    doc.add_paragraph()
    add_par(doc, '%s: %s' % (знач['роль_а'], знач['имя_а']), 'left', indent=False)
    add_par(doc, 'реквизиты: %s' % знач['реквизиты_а'], 'left', indent=False)
    add_par(doc, '__________________ / подпись /', 'left', indent=False)
    doc.add_paragraph()
    add_par(doc, '%s: %s' % (знач['роль_б'], знач['имя_б']), 'left', indent=False)
    add_par(doc, 'реквизиты: %s' % знач['реквизиты_б'], 'left', indent=False)
    add_par(doc, '__________________ / подпись /', 'left', indent=False)


def build_docx(номер, вид, знач, перечень_lines, график_lines, out_path):
    шапка, основа, разделы = тексты.ШАБЛОНЫ[вид]
    doc = new_doc()

    add_par(doc, шапка + ' № %s' % знач['номер'], 'center', indent=False, bold=True)
    строка_шапки = место_текст(знач['место'])
    if знач['дата']:
        строка_шапки += '  —  %s' % знач['дата']
    add_par(doc, строка_шапки, 'center', indent=False)
    doc.add_paragraph()

    преамбула = '%s%s %s, именуемый(ая) в дальнейшем «%s», с одной стороны, ' \
        'и %s%s %s, именуемый(ый) в дальнейшем «%s», с другой стороны, ' \
        'заключили настоящий Договор о нижеследующем.' % (
            (знач['субъ_а'] + ' ') if знач['субъ_а'] else '',
            знач['имя_а'],
            _подписатель(знач['имя_а'], знач['реквизиты_а']),
            знач['роль_а'],
            (знач['субъ_б'] + ' ') if знач['субъ_б'] else '',
            знач['имя_б'],
            _подписатель(знач['имя_б'], знач['реквизиты_б']),
            знач['роль_б'])
    add_par(doc, преамбула)

    def подставить(т):
        for к, зн in знач.items():
            т = т.replace('{%s}' % к, str(зн))
        return т

    for i, (заголовок, абзацы) in enumerate(разделы, 1):
        add_par(doc, '%d. %s' % (i, подставить(заголовок).upper()),
                'center', indent=False, bold=True)
        doc.add_paragraph()
        for j, абзац in enumerate(абзацы, 1):
            t = подставить(абзац)
            if '{перечень}' in t:
                for_линия_выше = t.split('{перечень}')[0]
                if for_линия_выше:
                    add_par(doc, '%d.%d. %s' % (i, j, for_линия_выше))
                для_список = список_в_doc(doc, i, перечень_lines or график_lines, знач)
                for_линия_после = t.split('{перечень}')[1]
                if for_линия_после:
                    add_par(doc, '%d.%d. %s' % (i, j, for_линия_после))
                continue
            add_par(doc, '%d.%d. %s' % (i, j, t.replace('\n', '\n     ') if '\n' in t else t))

    doc.add_paragraph()
    add_par(doc, 'ПРАВОВАЯ ОСНОВА', 'center', indent=False, bold=True)
    add_par(doc, основа, 'left', indent=False)
    add_par(doc, 'Формулировки — черновик; обязательная вычитка юристом перед '
                 'подписанием.', 'left', indent=False)
    doc.add_paragraph()
    add_par(doc, 'РЕКВИЗИТЫ И ПОДПИСИ СТОРОН', 'center', indent=False, bold=True)
    doc.add_paragraph()
    подпись_block(знач, doc)
    doc.save(out_path)
    return doc


def build_blank(вид, out_path):
    """Бланк-шаблон договора в .docx: настоящий текст с полями [в квадратных скобках].

    Тот же текст, что и в собранном договоре, но без данных сторон —
    вместо них поля для заполнения. Нужен юристу, чтобы править формулировки
    в Word перед подписанием.
    """
    шапка, основа, разделы = тексты.ШАБЛОНЫ[вид]
    doc = new_doc()

    add_par(doc, шапка + ' № [номер]', 'center', indent=False, bold=True)
    add_par(doc, '[место]  —  [дата]', 'center', indent=False)
    doc.add_paragraph()

    add_par(doc, '[наименование/ФИО стороны 1], именуемый(ая) в дальнейшем «[роль 1]», '
                 'с одной стороны, и [наименование/ФИО стороны 2], именуемый(ый) в дальнейшем '
                 '«[роль 2]», с другой стороны, заключили настоящий Договор о нижеследующем.')
    doc.add_paragraph()

    for i, (заголовок, абзацы) in enumerate(разделы, 1):
        add_par(doc, '%d. %s' % (i, _чистые_скобки(заголовок).upper()),
                'center', indent=False, bold=True)
        doc.add_paragraph()
        for j, абзац in enumerate(абзацы, 1):
            add_par(doc, '%d.%d. %s' % (i, j, _чистые_скобки(абзац)))
    doc.add_paragraph()
    add_par(doc, 'ПРАВОВАЯ ОСНОВА', 'center', indent=False, bold=True)
    add_par(doc, _чистые_скобки(основа), 'left', indent=False)
    add_par(doc, 'Формулировки — черновик; обязательная вычитка юристом перед '
                 'подписанием.', 'left', indent=False)
    doc.add_paragraph()
    add_par(doc, 'РЕКВИЗИТЫ И ПОДПИСИ СТОРОН', 'center', indent=False, bold=True)
    doc.add_paragraph()
    add_par(doc, '[роль 1]: [ФИО / наименование]', 'left', indent=False)
    add_par(doc, 'реквизиты: [реквизиты]', 'left', indent=False)
    add_par(doc, '__________________ / подпись /', 'left', indent=False)
    doc.add_paragraph()
    add_par(doc, '[роль 2]: [ФИО / наименование]', 'left', indent=False)
    add_par(doc, 'реквизиты: [реквизиты]', 'left', indent=False)
    add_par(doc, '__________________ / подпись /', 'left', indent=False)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    return out_path


def _чистые_скобки(т):
    """«{поле}» → «[поле]»: в бланке видимые поля для заполнения."""
    return re.sub(r'\{([^{}]+)\}', r'[\1]', т)


def список_в_doc(doc, i, линии, знач):
    if линии:
        for k, line in enumerate(линии, 1):
            add_par(doc, '     %s) %s' % (k, line), 'left', indent=False)
    else:
        add_par(doc, '     — (перечень не заполнен; см. предмет Договора)',
                'left', indent=False)
    doc.add_paragraph()


def _подписатель(имя, реквизиты):
    """«(паспорт: …)» для физлиц; для фирм — «в лице …» если есть р/с."""
    r = _str(реквизиты)
    m = re.search(r'паспорт[^;,]{0,80}', r, flags=re.IGNORECASE)
    if m:
        return '(%s)' % m.group(0)
    return ''


def build_акт(вид, знач, out_path):
    шапка = 'АКТ приёма-передачи к Договору № %s' % знач['номер']
    акт = new_doc()
    add_par(акт, шапка, 'center', indent=False, bold=True)
    add_par(акт, '%s — %s' % (место_текст(знач['место']), знач['дата']),
            'center', indent=False)
    акт.add_paragraph()
    add_par(акт, 'К Договору № %s от %s между %s (%s) и %s (%s).' % (
        знач['номер'], знач['дата'], знач['имя_а'], знач['роль_а'],
        знач['имя_б'], знач['роль_б']))
    add_par(акт, 'Передающее лицо передаёт, а Принимающее лицо принимает '
                 'следующее имущество (результат работ / услуг): '
                 '%s.' % знач['предмет'])
    add_par(акт, 'Состояние передаваемого имущества: ____________________; '
                 'комплектность: ____________________; недостатки: '
                 '____________________.')
    add_par(акт, 'Справедливо по взаимному согласию; подписывается в двух '
                 'экземплярах равной силы.')
    акт.add_paragraph()
    add_par(акт, 'РЕКВИЗИТЫ И ПОДПИСИ СТОРОН', 'center', indent=False, bold=True)
    акт.add_paragraph()
    подпись_block(знач, акт)
    акт.save(out_path)


# ─────────────────────────────────── сборка ─────────────────────────────────

def собери(book_path, демо=False):
    """Выполняет сборку; возвращает путь к папке выпуска."""
    договоры, перечень_raw, график_raw = read_book(book_path)
    штамп = dt.datetime.now().strftime('%Y-%m-%d_%H-%M')
    out = OUT_ROOT / штамп
    out.mkdir(parents=True, exist_ok=True)

    отчёт = ['ОТЧЁТ-СВЕРКИ (ДОГОВОРЫ)',
             'Анкета: %s' % Path(book_path).name,
             'Договоров в анкете: %d' % len(договоры),
             '']
    собрано, ошибок = 0, 0

    for i, d in enumerate(договоры, 1):
        вид = _str(d.get('Вид договора') or '')
        n = d['№']
        строки_перечня = перечень_raw.get(n, [])
        строки_графика = график_raw.get(n, [])
        ошибки, предупреждения = норма.проверить(d, строки_перечня)
        print('[%d/%d] Договор № %s (%s): %s' % (
            i, len(договоры), n, вид or '?',
            'ошибки: %d' % len(ошибки) if ошибки else 'проверка пройдена'), flush=True)
        if ошибки:
            ошибок += 1
            отчёт.append('ОШИБКИ — договор не сформирован:')
            отчёт.append('  ✗ № %s (%s):' % (n, вид or 'вид не указан'))
            for ош in ошибки:
                отчёт.append('      %s' % ош)
            if предупреждения:
                for пр in предупреждения:
                    отчёт.append('      ! %s' % пр)
            отчёт.append('')
            continue
        if предупреждения:
            for пр in предупреждения:
                отчёт.append('  ⚠ № %s (%s): %s' % (n, вид, пр))

        знач = значения(d, строки_перечня, строки_графика)
        имя = platform_utils.ascii_filename('contract-%02d-%s.docx' % (n, вид))
        build_docx(n, вид, знач, строки_перечня or строки_графика,
                   строки_графика, out / имя)
        собрано += 1
        if _str(d.get('Акт приёма-передачи (да/нет)')).lower() in ('да', 'yes', 'да.'):
            акт = 'handover-%02d.docx' % n
            try:
                build_акт(вид, знач, out / акт)
            except Exception as exc:                     # pragma: no cover
                отчёт.append('  ⚠ № %s: акт не собран — %s' % (n, exc))

    if собрано == 0 and ошибок:
        отчёт.insert(2, '')
        отчёт.insert(3, 'НИ ОДНОГО договора не собрано: исправьте ошибки в анкете.')

    отчёт_path = out / 'validation-report-contracts.txt'
    отчёт.append('')
    отчёт.append('Сформировано договоров: %d в папку: %s' % (собрано, out))
    отчёт_path.write_text('\n'.join(отчёт), encoding='utf-8')
    (BASE / 'validation-report-contracts.txt').write_text(
        '\n'.join(отчёт), encoding='utf-8')   # дубль для окна журнала ГУИ
    print(отчёт_path.read_text(encoding='utf-8'))
    return out


def main():
    argv = [a for a in sys.argv[1:]]
    демо = '--демо' in argv
    пути = [a for a in argv if not a.startswith('--')]

    # --шаблон <вид> — бланк-шаблона в .docx (для просмотра и правки в Word)
    if '--шаблон' in argv:
        вид = argv[argv.index('--шаблон') + 1] if len(argv) > argv.index('--шаблон') + 1 else ''
        if вид not in тексты.ШАБЛОНЫ:
            print('Неизвестный вид договора: %s' % (вид or '—'))
            print('Доступные: %s' % '; '.join(тексты.ШАБЛОНЫ))
            sys.exit(1)
        out = BASE / 'contract-templates' / platform_utils.ascii_filename('template-%s.docx' % вид)
        build_blank(вид, out)
        print(out)
        return
    if демо:
        book = BASE / 'contracts-demo.xlsx'
        if not book.exists():
            import create_contract_workbook as создать
            sys.argv = [sys.argv[0], str(book), '--демо']
            создать.main()
        собери(book, демо=True)
        return
    book = Path(пути[0]) if пути else DEFAULT_BOOK
    if not book.exists():
        print('Книга не найдена: %s' % book)
        print('Создайте её первым шагом или запустите с --демо для примера.')
        sys.exit(1)
    собери(book)


if __name__ == '__main__':
    main()
