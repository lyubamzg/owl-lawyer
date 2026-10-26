#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Рабочее место юриста: таблица → редактор → документы.

Боковая навигация: ЖКУ, семейное право, защита прав потребителя и договоры.
В каждом разделе — три карточки с последовательными этапами работы.
  1. Загрузить данные (Портянка + сканы Ф-9)
  2. Открыть таблицу в LibreOffice для редактирования
  3. Собрать документы
После каждого успешного этапа появляется кнопка «Следующий шаг» — она выполняет
следующий этап раздела; можно идти пошагово или пользоваться кнопками самому.
"""
import re
import subprocess
import threading
import queue
import time
import sys
from pathlib import Path

import platform_utils

import tkinter as tk
from tkinter import messagebox, filedialog, font as tkfont
from tkinter import ttk
from ui_theme import (Button, button_colors, BG, SIDEBAR, SURFACE,
                      SURFACE_HOVER, BORDER, FG, MUTED, ACCENT,
                      PRIMARY, PRIMARY_HOVER, SUCCESS)

BASE = Path(__file__).resolve().parent
ICON = platform_utils.icon_path()
BOOKS_JKU = ['housing-input.xlsx', 'input-workbook.xlsx']
BOOKS_SEM = ['family-questionnaire.xlsx']
BOOKS_POTR = ['consumer-questionnaire.xlsx']
BOOKS_DOG = ['contracts-workbook.xlsx']

root = tk.Tk()
root.title('СОВА — рабочее место юриста')
root.geometry('1180x780')
root.minsize(1040, 700)
root.configure(bg=BG)

try:
    if ICON.exists():
        _img = tk.PhotoImage(file=str(ICON))
        root.iconphoto(True, _img)
except tk.TclError:
    pass

family = platform_utils.pick_family(tkfont, ('SF Pro Text', 'Helvetica Neue', *platform_utils.UI_FONT))
big = tkfont.Font(family=family, size=12, weight='bold')
mid = tkfont.Font(family=family, size=11)
small = tkfont.Font(family=family, size=10)
heading = tkfont.Font(family=family, size=25, weight='bold')
logf = tkfont.Font(family=platform_utils.pick_family(tkfont, platform_utils.UI_FONT_MONO), size=10)

sidebar = tk.Frame(root, bg=SIDEBAR, width=224)
sidebar.pack(side='left', fill='y')
sidebar.pack_propagate(False)
brand = tk.Frame(sidebar, bg=SIDEBAR)
brand.pack(fill='x', padx=24, pady=(30, 32))
brand_lockup = tk.Frame(brand, bg=SIDEBAR)
brand_lockup.pack(anchor='w')
logo_path = ICON.with_name('owl-lawyer-256.png')
if not logo_path.is_file():
    logo_path = ICON
if logo_path.is_file():
    logo_source = tk.PhotoImage(file=str(logo_path))
    logo_scale = max(1, (logo_source.width() + 43) // 44)
    logo_image = logo_source.subsample(logo_scale, logo_scale)
    brand_icon = tk.Label(brand_lockup, image=logo_image, bg=SIDEBAR, borderwidth=0)
    brand_icon.pack(side='left', padx=(0, 10))
brand_wordmark = tk.Label(brand_lockup, text='СОВА', font=(family, 24, 'bold'),
                          bg=SIDEBAR, fg=FG, anchor='w')
brand_wordmark.pack(side='left')
tk.Label(brand, text='Рабочее место юриста', font=small, bg=SIDEBAR,
         fg=MUTED, anchor='w').pack(fill='x', pady=(4, 0))
tk.Label(sidebar, text='НАПРАВЛЕНИЯ', font=(family, 9, 'bold'), bg=SIDEBAR,
         fg='#65778F', anchor='w').pack(fill='x', padx=24, pady=(0, 12))
tab_row = tk.Frame(sidebar, bg=SIDEBAR)
tab_row.pack(fill='x', padx=12)
privacy = tk.Frame(sidebar, bg=SIDEBAR)
privacy.pack(side='bottom', fill='x', padx=24, pady=26)
tk.Frame(privacy, bg=BORDER, height=1).pack(fill='x', pady=(0, 18))
tk.Label(privacy, text='●  Локальное хранение', font=small, bg=SIDEBAR,
         fg=SUCCESS, anchor='w').pack(fill='x')
tk.Label(privacy, text='Документы остаются\nна вашем компьютере', font=small,
         bg=SIDEBAR, fg=MUTED, justify='left', anchor='w').pack(fill='x', pady=(8, 0))

workspace = tk.Frame(root, bg=BG)
workspace.pack(side='left', fill='both', expand=True, padx=32, pady=(30, 24))
tk.Label(workspace, text='ПОДГОТОВКА ДОКУМЕНТОВ', font=(family, 9, 'bold'),
         bg=BG, fg=ACCENT, anchor='w').pack(fill='x')
page_title = tk.Label(workspace, text='Жилищное право', font=heading,
                      bg=BG, fg=FG, anchor='w')
page_title.pack(fill='x', pady=(10, 8))
page_description = tk.Label(workspace, text='', font=mid, bg=BG, fg=MUTED,
                            justify='left', anchor='w', wraplength=700)
page_description.pack(fill='x')
status_panel = tk.Frame(workspace, bg=SURFACE, highlightbackground=BORDER,
                        highlightthickness=1)
status_panel.pack(fill='x', pady=(24, 24))
status_dot = tk.Label(status_panel, text='●', font=small, bg=SURFACE, fg=ACCENT)
status_dot.pack(side='left', padx=(16, 10), pady=13)
status_label = tk.Label(status_panel, text='', font=small, bg=SURFACE, fg=MUTED,
                        anchor='w', justify='left', wraplength=650)
status_label.pack(side='left', fill='x', expand=True, padx=(0, 16), pady=12)
main_content = tk.Frame(workspace, bg=BG)
main_content.pack(fill='both', expand=True)
current_section = {'books': BOOKS_JKU, 'word': 'таблица'}


def set_status(text, color=MUTED):
    status_label.config(text=text, fg=color)
    status_dot.config(fg=color)


def book_path(books):
    for book in books:
        if (BASE / book).exists():
            return book
    return None


def refresh_status(books=None, word=None):
    books = books if books is not None else current_section['books']
    word = word or current_section['word']
    book = book_path(books)
    text = ('Текущая ' + word + ': ' + book if book
            else 'Начните с загрузки документов — ' + word + ' пока не создана')
    set_status(text)


refresh_status()

_missing = platform_utils.missing_tools()
if _missing:
    messagebox.showwarning(
        'Не всё установлено',
        'Для полной работы Сова-юриста не хватает:\n\n  • '
        + '\n  • '.join(_missing)
        + '\n\nУстановите по README-файлам пакета Сова-юрист и перезапустите.'
          '\nБез этого кнопки, зависящие от программы, работать не будут.', parent=root)

buttons = {}
activity_panel = tk.Frame(main_content, bg=SURFACE, highlightbackground=BORDER,
                          highlightthickness=1)
activity_panel.pack(fill='both', expand=True, pady=(20, 0))
activity_heading = tk.Frame(activity_panel, bg=SURFACE)
activity_heading.pack(fill='x', padx=18, pady=(16, 12))
tk.Label(activity_heading, text='Ход работы', font=big, bg=SURFACE,
         fg=FG, anchor='w').pack(side='left')
tk.Label(activity_heading, text='Результаты и замечания', font=small,
         bg=SURFACE, fg=MUTED).pack(side='right')
tk.Frame(activity_panel, bg=BORDER, height=1).pack(fill='x')
log_body = tk.Frame(activity_panel, bg=SURFACE)
log_body.pack(fill='both', expand=True, padx=18, pady=14)
log_scroll = ttk.Scrollbar(log_body, orient='vertical')
log = tk.Text(log_body, bg=SURFACE, fg=MUTED, insertbackground=FG, font=logf,
              wrap='word', state='disabled', height=8, relief='flat',
              borderwidth=0, highlightthickness=0, padx=0, pady=0,
              spacing1=3, spacing3=3, yscrollcommand=log_scroll.set)
log.pack(side='left', fill='both', expand=True)
log_scroll.config(command=log.yview)
log_scroll.pack(side='right', fill='y', padx=(12, 0))
log_placeholder = tk.Label(log_body,
    text='Пока здесь тихо\n\nНачните с загрузки документов. Здесь появятся\nрезультаты обработки и замечания для проверки.',
    font=mid, bg=SURFACE, fg=MUTED, justify='center')
log_placeholder.place(relx=0.5, rely=0.5, anchor='center')

prog_frame = tk.Frame(main_content, bg=BG)
prog_label = tk.Label(prog_frame, font=mid, bg=BG, fg=FG, anchor='w')
prog_label.pack(fill='x')
prog_bar = ttk.Progressbar(prog_frame, orient='horizontal', length=560,
                           mode='indeterminate')
prog_bar.pack(fill='x', pady=(8, 4))


q = queue.Queue()


# ─── Визуал загрузки: полоса прогресса «файл X из Y», время и оценка остатка ───
prog_state = {'n': 0, 'total': 0, 't0': None, 'times': []}
PROG_RE = re.compile(r'^\[(\d+)/(\d+)\]')


def show_progress(title):
    """Показать полосу хода работы (создать при первом вызове)."""
    prog_state.update({'n': 0, 'total': 0, 't0': time.time(), 'times': []})
    prog_label.config(text=title, fg=FG)
    prog_bar.config(mode='indeterminate')
    prog_bar.start(10)
    set_status('Обработка документов…', color=ACCENT)
    prog_frame.pack(before=activity_panel, fill='x', pady=(16, 0))


def hide_progress():
    prog_frame.pack_forget()
    prog_bar.stop()
    prog_state['t0'] = None


def bump_progress(n, total, note=''):
    """Отметка «файл n из total готов»; показывает время и оценку остатка."""
    prog_state.update({'n': n, 'total': max(total, 1)})
    prog_state['times'].append(time.time())
    prog_bar.config(mode='determinate', maximum=max(total, 1), value=n)


def tick_progress(elapsed_only=False):
    t0 = prog_state.get('t0')
    if t0 is None:
        return
    el = time.time() - t0
    n, total = prog_state['n'], prog_state['total']
    txt = f'прошло {el:4.0f} с'
    if total and n:
        ts = prog_state['times']
        if len(ts) >= 2:
            avg = (ts[-1] - ts[0]) / (len(ts) - 1)
            txt += f'  ·  осталось ≈ {max(0, (total - n)) * avg:3.0f} с'
        else:
            txt += f'  ·  готово к этому моменту: {n} из {total}'
    elif elapsed_only:
        txt += '  ·  объём уточняется по ходу…'
    prog_label.config(text=prog_label.cget('text').split('   ·')[0] + '   ·   ' + txt
                      if '·' not in prog_label.cget('text') else
                      prog_label.cget('text').rsplit('   ·', 1)[0] + '   ·   ' + txt)


# ─── Пошаговый режим: после каждого этапа — кнопка «Следующий шаг» ───
pending_next = {'cmd': None}


def arm_continue(next_cmd, note=''):
    """Показать «Следующий шаг»; нажатие выполнит следующий этап раздела."""
    if not next_cmd:
        return
    pending_next['cmd'] = next_cmd
    cont_btn.config(state='normal')
    cont_btn.pack(before=activity_panel, anchor='e', pady=(16, 0))
    set_status(
                         text=(note or 'Этап выполнен') + '. Можно перейти к следующему шагу.',
                         color=ACCENT)
    q.put('')
    q.put('Этап выполнен. Для продолжения нажмите «Следующий шаг».')


def disarm_continue():
    """Убрать кнопку (новый этап запущен, или этап не удался)."""
    pending_next['cmd'] = None
    cont_btn.pack_forget()


def continue_next():
    nxt = pending_next['cmd']
    disarm_continue()
    if nxt:
        nxt()


def step_hook(steps, i):
    """Хук завершения этапа i: показать «Следующий шаг» для этапа i+1; после последнего — итог."""
    def hook(c=0):
        entry = current_section.get('entry')
        if entry:
            entry['markers'][i].config(text='✓', fg=SUCCESS)
        nxt = steps[i + 1] if i + 1 < len(steps) else None
        if nxt:
            arm_continue(nxt, f'Этап {i + 1}-й из {len(steps)} выполнен')
        else:
            set_status(
                                 text='✔  Все этапы раздела выполнены — документы готовы',
                                 color=SUCCESS)
            q.put('')
            q.put('✔  ВСЕ ЭТАПЫ РАЗДЕЛА ВЫПОЛНЕНЫ — документы готовы. '
                  'Проверьте документы в открывшейся папке.')
    return hook


cont_btn = Button(main_content, text='Следующий шаг  →', font=big,
                     bg=PRIMARY, fg=FG, activebackground=PRIMARY_HOVER,
                     relief='flat', cursor='hand2', padx=22, pady=10,
                     command=continue_next)
# сразу не упакован — появится после успешного завершения очередного этапа


LOG_MAX_LINES = 2000  # предел журнала — чтобы окно за день не копило память


def pump():
    global _last_tick
    try:
        while True:
            line = q.get_nowait()
            m = PROG_RE.match(line)
            if m:
                bump_progress(int(m.group(1)), int(m.group(2)))
                tick_progress()
            if line.strip():
                log_placeholder.place_forget()
            log.config(state='normal')
            log.insert('end', line + '\n')
            total = int(log.index('end-1c').split('.')[0])
            if total > LOG_MAX_LINES:
                log.delete('1.0', f'{total - LOG_MAX_LINES}.0')
            log.see('end')
            log.config(state='disabled')
    except queue.Empty:
        pass
    if time.time() - _last_tick > 1:
        _last_tick = time.time()
        tick_progress(elapsed_only=True)
    root.after(100, pump)


root.after(100, pump)
_last_tick = 0.0


def set_buttons(busy=True):
    for b in buttons.values():
        b.config(state='disabled' if busy else 'normal')
    cont_btn.config(state='disabled' if busy else 'normal')
    for entry in sections:
        entry['btn'].config(state='disabled' if busy else 'normal')
    contracts_nav.config(state='disabled' if busy else 'normal')
    if not busy:
        hide_progress()


def run_cmd(cmd, done_msg, after=lambda c: None):
    """Запуск команды в фоне с выводом в журнал."""
    def work():
        try:
            proc = platform_utils.popen(cmd, cwd=BASE, stdout=subprocess.PIPE,
                                   stderr=subprocess.STDOUT, text=True,
                                   bufsize=1)
            for line in proc.stdout:
                q.put(line.rstrip())
            proc.wait()
            proc.stdout.close()
            rc = proc.returncode
        except Exception as ex:  # noqa: BLE001
            q.put(f'! Ошибка запуска: {ex}')
            rc = -1

        def finish():
            set_buttons(False)
            if rc == 0:
                q.put('')
                q.put(done_msg)
                q.put('')
                refresh_status()
                after(rc)
            else:
                set_status('Не удалось завершить обработку. Проверьте замечания в журнале.', color='#D6A298')
                q.put('')
                q.put('!!! Команда завершилась с ошибкой — прочитайте журнал выше. '
                  'Запустите этап заново кнопкой раздела.')
        root.after(0, finish)

    threading.Thread(target=work, daemon=True).start()


def папка_загрузок():
    """Папка «Загрузка» (или Downloads на английской системе) — старт для выбора файлов."""
    home = Path.home()
    for nm in ('Загрузка', 'Downloads'):
        p = home / nm
        if p.is_dir():
            return p
    return home


def ask_documents(initial=None):
    """Окно выбора документов со своим списком: Ctrl+A выделяет всё.

    Системное окно выбора файлов на некоторых системах не понимает Ctrl+A,
    поэтому список файлов в папке рисуем сами. Показываем ЛЮБЫЕ файлы —
    импортёр разберётся с форматами сам.
    """
    dlg = tk.Toplevel(root)
    dlg.title('Какие документы загрузить?')
    dlg.configure(bg=BG)
    dlg.transient(root)
    dlg.resizable(True, True)

    out = {'picked': None}
    folder = {'path': Path(initial) if initial else Path.home()}
    rows = []       # (тип строки, путь): 'dir' — папка, 'file' — документ

    tk.Label(dlg, text='Какие документы загрузить?', font=big, bg=BG, fg=FG,
             justify='center').pack(padx=18, pady=(14, 2))
    tk.Label(dlg, text='Папки — двойной клик · Ctrl+A — все документы '
                       '· Ctrl/Shift+клик — по одному',
             font=mid, bg=BG, fg=MUTED).pack(padx=18, pady=(0, 8))

    dirlabel = tk.Label(dlg, text='', font=mid, bg=BG, fg=SUCCESS)
    dirlabel.pack(fill='x', padx=18)

    body = tk.Frame(dlg, bg=BG)
    body.pack(fill='both', expand=True, padx=18, pady=6)
    lb = tk.Listbox(body, selectmode='extended', exportselection=False,
                    font=mid, bg=SURFACE, fg=FG, activestyle='dotbox',
                    highlightthickness=1, highlightbackground=SURFACE_HOVER)
    sb = tk.Scrollbar(body, command=lb.yview)
    lb.config(yscrollcommand=sb.set)
    sb.pack(side='right', fill='y')
    lb.pack(side='left', fill='both', expand=True)

    def refill():
        lb.delete(0, 'end')
        rows.clear()
        dirs, files = [], []
        for p in sorted(folder['path'].iterdir(), key=lambda p: str(p).lower()):
            if p.name.startswith('.'):
                continue
            if p.is_dir():
                dirs.append(p)
            elif p.is_file():
                files.append(p)
        for d in dirs:
            lb.insert('end', '📁  ' + d.name)
            rows.append(('dir', d))
        for f in files:
            lb.insert('end', '      ' + f.name)
            rows.append(('file', f))
        dirlabel.config(text='Папка: ' + str(folder['path'])
                        + ('   —  документов: ' + str(len(files)) if files
                           else '   —  документов нет'))
        if not rows:
            lb.insert('end', '(папка пуста)')

    def open_row(idx):
        kind, path = rows[idx]
        if kind == 'dir':
            folder['path'] = path
            clear_sel()
            refill()
        elif kind == 'file':
            lb.selection_clear(0, 'end')
            lb.selection_set(idx)
            lb.selection_anchor(idx)
            ok()

    def prune_to_files(e=None):
        """Папки из выделения убираем: Ctrl+A и прямоугольник — только документы."""
        sel = lb.curselection()
        files_sel = [i for i in sel if rows[i][0] == 'file']
        if len(files_sel) != len(sel):
            lb.selection_clear(0, 'end')
            for i in files_sel:
                lb.selection_set(i)
        return None

    def select_all(_e=None):
        lb.selection_clear(0, 'end')
        for i, (kind, _p) in enumerate(rows):
            if kind == 'file':
                lb.selection_set(i)
        lb.see(0)
        return 'break'

    def clear_sel():
        lb.selection_clear(0, 'end')

    def ok(_e=None):
        picked = [rows[i][1] for i in lb.curselection() if rows[i][0] == 'file']
        if not picked:
            messagebox.showwarning('Ничего не выбрано',
                                   'Выделите хотя бы один документ (или нажмите Ctrl+A).',
                                   parent=dlg)
            return
        out['picked'] = [str(p) for p in picked]
        dlg.destroy()

    def cancel(_e=None):
        dlg.destroy()

    def switch_dir(_e=None):
        nd = filedialog.askdirectory(parent=dlg, initialdir=str(folder['path']),
                                     title='Папка с документами')
        if nd:
            folder['path'] = Path(nd)
            refill()

    def up_dir(_e=None):
        parent = folder['path'].parent
        if parent != folder['path']:
            folder['path'] = parent
            refill()

    # Ctrl+A при любой раскладке (латинице и кириллице).
    def ctrl_key(e):
        if e.keycode in (38, 65) or e.keysym.lower() in ('a', 'cyrillic_a'):
            return select_all()
    dlg.bind('<Control-KeyPress>', ctrl_key)
    lb.bind('<Control-KeyPress>', ctrl_key)
    for seq in ('<Control-a>', '<Control-A>', '<Control-Cyrillic_a>',
                '<Control-Cyrillic_A>'):
        lb.bind(seq, select_all)
    dlg.bind('<Escape>', cancel)
    dlg.bind('<Return>', ok)
    lb.bind('<Double-Button-1>',
            lambda e: (open_row(lb.nearest(e.y)), 'break')[1])
    # Прямоугольник мышью / обычные клики — папки из выделения убираем после отпускания кнопки
    lb.bind('<ButtonRelease-1>', prune_to_files)

    nav = tk.Frame(dlg, bg=BG)
    nav.pack(fill='x', padx=18)
    Button(nav, text='↑  Папка выше', font=mid, bg=SURFACE_HOVER, fg=FG,
              activebackground=SURFACE_HOVER, relief='flat', cursor='hand2',
              padx=12, pady=5, command=up_dir).pack(side='left', pady=4)
    Button(nav, text='📁  Другая папка…', font=mid, bg=SURFACE_HOVER, fg=FG,
              activebackground=SURFACE_HOVER, relief='flat', cursor='hand2',
              padx=12, pady=5, command=switch_dir).pack(side='left', padx=(8, 0), pady=4)

    sel = tk.Frame(dlg, bg=BG)
    sel.pack(fill='x', padx=18)
    Button(sel, text='Выделить все', font=mid, bg=SURFACE_HOVER, fg=FG,
              activebackground=SURFACE_HOVER, relief='flat', cursor='hand2',
              padx=12, pady=5, command=select_all).pack(side='left', pady=4)
    Button(sel, text='Снять выделение', font=mid, bg=SURFACE_HOVER, fg=FG,
              activebackground=SURFACE_HOVER, relief='flat', cursor='hand2',
              padx=12, pady=5, command=clear_sel).pack(side='left', padx=(8, 0), pady=4)

    Button(dlg, text='Загрузить выбранные', font=big, bg=PRIMARY, fg='#FFFFFF',
              activebackground=PRIMARY_HOVER, relief='flat', cursor='hand2',
              padx=16, pady=8, command=ok).pack(fill='x', padx=18, pady=(10, 4))
    Button(dlg, text='Отмена', font=big, bg=SURFACE_HOVER, fg='#FFFFFF',
              activebackground=SURFACE_HOVER, relief='flat', cursor='hand2',
              padx=16, pady=8, command=cancel).pack(fill='x', padx=18, pady=(0, 14))

    refill()
    dlg.update_idletasks()
    dlg.geometry('600x520+%d+%d' % (
        max(0, root.winfo_rootx() + (root.winfo_width() - 600) // 2),
        max(0, root.winfo_rooty() + (root.winfo_height() - 520) // 2)))
    dlg.grab_set()
    dlg.focus_set()
    root.wait_window(dlg)
    return out['picked']


def make_load_data(script, header, done_msg, tips=(), src_dir=None,
                   on_done=None):
    def cmd():
        # Спрашиваем, какие документы загрузить (видны все форматы, выбор нескольких).
        # Документы выбираем начиная с папки «Загрузка».
        initial = папка_загрузок()
        try:
            picked = ask_documents(initial)
        except Exception as ex:  # noqa: BLE001 — своё окно не открылось → системное
            q.put('! Своё окно выбора не открылось (%s) — открываю системное.' % ex)
            picked = filedialog.askopenfilenames(
                parent=root, initialdir=str(initial), filetypes=FILETYPES,
                title='Какие документы загрузить?')
        picked = [p for p in (picked or [])]
        if not picked:
            q.put('')
            q.put('(загрузка отменена — файлы не выбраны)')
            where_hint = Path(src_dir) if src_dir else BASE
            q.put(f'Положите документы в «{where_hint.name}» рядом с приложением или выберите их в любой папке:')
            q.put('  pdf, jpg/png, Word (doc/docx/rtf/odt), Excel (xlsx/xls/ods), txt/csv.')
            refresh_status()
            return
        others = ask_extra_scripts(script) if script in EXTRA_LOADS else []
        set_buttons(True)
        show_progress('Загрузка документов… (файлы обрабатываются по очереди)')
        q.put('')
        q.put(header)
        q.put('Выбрано документов: ' + str(len(picked)) + ' — из «' + str(Path(picked[0]).parent) + '»')
        if len(picked) > 1:
            q.put('несколько сразу: Ctrl/Shift + клик, Ctrl+A — все файлы')
        if others:
            q.put('Эти же документы пойдут и в разделы: '
                  + ', '.join(EXTRA_LOADS[sc] for sc in others))
        for t in tips:
            q.put(t)

        def finish_all(last_ok):
            if others:
                q.put('── Все выбранные разделы отработали эти же документы ──')
            if on_done:
                on_done(last_ok)

        def run_others(i=0):
            if i >= len(others):
                finish_all(0)
                return
            sc = others[i]
            q.put('')
            q.put('── ТЕ ЖЕ ДОКУМЕНТЫ — ЗАГРУЗКА В РАЗДЕЛ «' + EXTRA_LOADS[sc] + '» ──')
            set_buttons(True)   # между скриптами кнопки снова блокируем
            show_progress('Загрузка тех же документов в «' + EXTRA_LOADS[sc] + '»…')
            run_cmd([sys.executable, sc, *picked],
                    '── раздел «' + EXTRA_LOADS[sc] + '» обработал эти же документы ──',
                    lambda c, i=i: run_others(i + 1))

        run_cmd([sys.executable, script, *picked], done_msg,
                run_others if others else (on_done if on_done else (lambda c: None)))
    return cmd


def make_open_editor(books, word='таблица', on_done=None):
    def cmd():
        b = book_path(books)
        if not b:
            messagebox.showwarning('Нет ' + ('таблицы' if word == 'таблица' else 'анкеты'),
                                   'Сначала загрузите данные — файл создастся автоматически.')
            return
        platform_utils.open_in_libreoffice(BASE / b)
        q.put(f'Открываю «{b}» в LibreOffice…')
        if on_done:
            on_done(0)
    return cmd


def make_build_docs(books, script, report_name, out_glob, header,
                    on_done=None):
    def cmd():
        b = book_path(books)
        if not b:
            messagebox.showwarning('Нет данных', 'Сначала загрузите данные.')
            return
        set_buttons(True)
        show_progress('Сбор документов…')
        q.put('')
        q.put(f'── {header} («{b}») ──')
        q.put('Файл должен быть сохранён в LibreOffice (Ctrl+S)!')

        def done(c):
            rpt = BASE / report_name
            if rpt.exists():
                q.put('──── ' + report_name.replace('.txt', '').upper() + ' ────')
                q.put(rpt.read_text(encoding='utf-8', errors='replace'))
                q.put('───────────────────────')
            out = sorted(BASE.glob(out_glob))
            if c == 0 and out:
                platform_utils.reveal(out[-1])
                q.put('')
                q.put('Готово. Папка с документами: ' + out[-1].name)
            if on_done:
                on_done(c)

        run_cmd([sys.executable, script, b], '', done)
    return cmd


# ─── Выбор входящих файлов и дополнительные разделы ───
# фильтры выбора файлов: все поддержанные форматы документов
FILETYPES = [('Документы (pdf, jpg, Word, Excel, txt)',
              '*.pdf *.jpg *.jpeg *.png *.tif *.tiff *.bmp '
              '*.docx *.doc *.rtf *.odt *.xlsx *.xls *.ods *.txt *.csv'),
             ('PDF', '*.pdf'),
             ('Сканы (jpg, png, tif, bmp)', '*.jpg *.jpeg *.png *.tif *.tiff *.bmp'),
             ('Word (doc, docx, rtf, odt)', '*.docx *.doc *.rtf *.odt'),
             ('Excel (xlsx, xls, ods)', '*.xlsx *.xls *.ods'),
             ('Текст (txt, csv)', '*.txt *.csv'),
             ('Все файлы', '*')]

# скан-разделы: те же документы можно загрузить сразу в несколько из них
EXTRA_LOADS = {'import_family.py': 'СЕМЕЙНОЕ ПРАВО',
               'import_consumer.py': 'ЗАЩИТА ПРАВ ПОТРЕБИТЕЛЯ'}


def ask_extra_scripts(current):
    """Диалог после выбора файлов: загрузить их же и в другие разделы? → список скриптов."""
    dlg = tk.Toplevel(root)
    dlg.title('Загрузить и в другие разделы?')
    dlg.configure(bg=BG)
    dlg.transient(root)
    dlg.resizable(False, False)
    tk.Label(dlg, text='Эти же документы загрузить ещё и в:', font=big, bg=BG, fg=FG,
             justify='center').pack(padx=22, pady=(16, 10))
    chosen, order = {}, []
    for sc, nm in EXTRA_LOADS.items():
        if sc == current:
            continue
        v = tk.BooleanVar(value=True)
        chosen[sc] = v
        order.append(sc)
        tk.Checkbutton(dlg, text=nm, variable=v, font=big, bg=BG, fg=FG,
                       selectcolor=SURFACE_HOVER, activebackground=BG,
                       highlightthickness=0).pack(anchor='w', padx=34, pady=2)
    out = {'scripts': []}
    def ok_run():
        out['scripts'] = [sc for sc in order if chosen[sc].get()]
        dlg.destroy()
    def ok_skip():
        dlg.destroy()
    Button(dlg, text='Загрузить выбранные разделы тоже', font=big, bg=PRIMARY,
              fg='#FFFFFF', activebackground=PRIMARY_HOVER, relief='flat', cursor='hand2',
              padx=16, pady=8, command=ok_run).pack(fill='x', padx=22, pady=(12, 4))
    Button(dlg, text='Только текущий раздел', font=big, bg=SURFACE_HOVER, fg='#FFFFFF',
              activebackground=SURFACE_HOVER, relief='flat', cursor='hand2',
              padx=16, pady=8, command=ok_skip).pack(fill='x', padx=22, pady=(0, 16))
    dlg.grab_set()
    root.wait_window(dlg)
    return out['scripts']
def договоры_список():
    """Строки анкеты договоров для окна: [(текст строки, ключ поиска)] или (None, [])."""
    path = BASE / 'contracts-workbook.xlsx'
    if not path.exists():
        return None, []
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        rows = list(wb['Договоры'].values)
        wb.close()
    except Exception as ex:   # noqa: BLE001
        return None, [('Анкета повреждена, не открылась: %s' % ex, '')]
    if not rows:
        return None, []
    заг = [str(h).strip() if h else '' for h in rows[0]]
    out = []
    for row in rows[1:]:
        if row[0] is None and (len(row) < 2 or row[1] in (None, '')):
            continue
        d = dict(zip(заг, row))
        n = d.get('№')
        вид = d.get('Вид договора') or '?'
        кто = '%s ↔ %s' % (str(d.get('Сторона 1 (кто)') or '…')[:22],
                           str(d.get('Сторона 2 (кто)') or '…')[:22])
        цена = d.get('Цена (сумма), руб.')
        текст = '№ %s  ·  %s  ·  %s%s' % (n, вид, кто,
                                          ('  ·  %s ₽' % цена) if цена not in (None, '') else '')
        ключ = ' '.join(str(v) for v in row if v is not None).lower()
        out.append((текст, ключ))
    return path, out


def окно_договоров():
    """Раздел ДОГОВОРЫ: окно всех договоров анкеты + поисковая строка по названию."""
    dlg = tk.Toplevel(root)
    dlg.title('СОВА — договоры')
    dlg.configure(bg=BG)
    dlg.transient(root)
    dlg.resizable(True, True)
    dlg.geometry('900x660')
    dlg.minsize(780, 600)

    tk.Label(dlg, text='Договоры', font=heading, bg=BG, fg=FG,
             anchor='w').pack(fill='x', padx=24, pady=(24, 8))
    tk.Label(dlg, text='Найдите договор по виду, сторонам или предмету. '
                       'Двойной клик откроет шаблон в Word.',
             font=mid, bg=BG, fg=MUTED, justify='left', anchor='w',
             wraplength=740).pack(fill='x', padx=24, pady=(0, 18))

    top = tk.Frame(dlg, bg=BG)
    top.pack(fill='x', padx=24, pady=(0, 16))
    tk.Label(top, text='Поиск', font=mid, bg=BG, fg=MUTED).pack(side='left')
    поиск = tk.StringVar()
    ent = tk.Entry(top, textvariable=поиск, font=big, bg=SURFACE, fg=FG,
                   insertbackground=FG, relief='flat', highlightthickness=1,
                   highlightbackground=BORDER, highlightcolor=ACCENT)
    ent.pack(side='left', fill='x', expand=True, padx=(8, 0), ipady=5)
    btn_найти = Button(top, text='Найти', font=mid, bg=PRIMARY, fg=FG,
                          activebackground=PRIMARY_HOVER, relief='flat', cursor='hand2',
                          padx=16, pady=6, command=lambda: найти())
    btn_найти.pack(side='left', padx=(8, 0))

    midf = tk.Frame(dlg, bg=BG)
    midf.pack(fill='both', expand=True, padx=24, pady=0)
    sb = ttk.Scrollbar(midf, orient='vertical')
    lb = tk.Listbox(midf, font=big, bg=SURFACE, fg=FG, selectmode='browse',
                    exportselection=False, activestyle='none', relief='flat',
                    highlightthickness=1, highlightbackground=BORDER,
                    highlightcolor=ACCENT, selectbackground=PRIMARY, selectforeground=FG,
                    yscrollcommand=sb.set)
    lb.pack(side='left', fill='both', expand=True)
    sb.config(command=lb.yview)
    sb.pack(side='right', fill='y')

    счёт = tk.Label(dlg, text='', font=mid, bg=BG, fg=SUCCESS)
    счёт.pack(fill='x', padx=24, pady=(10, 0))

    def refill(*_):
        lb.delete(0, 'end')
        path, rows = договоры_список()
        запрос = поиск.get().strip().lower()
        if path is None:
            lb.insert('end', '  Здесь появятся ваши договоры')
            lb.insert('end', '  Начните с кнопки «Создать анкету»')
            счёт.config(text='Создайте анкету, чтобы добавить первый договор', fg=MUTED)
            return
        найдено = 0
        for текст, ключ in rows:
            if not запрос or запрос in ключ:
                lb.insert('end', текст)
                найдено += 1
        счёт.config(text='Найдено договоров: %d из %d' % (найдено, len(rows)),
                    fg=SUCCESS)

    ищем = {'идёт': False}

    def показать_результат():
        refill()
        btn_найти.config(state='normal')
        lb.config(bg=SURFACE)
        dlg.configure(cursor='')
        ищем['идёт'] = False

    def найти(*_):
        if ищем['идёт']:
            return
        ищем['идёт'] = True
        btn_найти.config(state='disabled')
        lb.config(bg=SURFACE_HOVER)          # список затемняется — видно, что идёт поиск
        dlg.configure(cursor='watch')    # курсор «часы»
        счёт.config(text='⌛  Ищу договоры в анкете…', fg=ACCENT)
        dlg.after(320, показать_результат)

    поиск.trace_add('write', refill)
    ent.bind('<Return>', найти)

    def открыть_анкету(_=None):
        b = book_path(BOOKS_DOG)
        if not b:
            messagebox.showinfo('Нет анкеты', 'Анкета ещё не создана — нажмите «➕ Создать анкету».')
            return
        platform_utils.open_in_libreoffice(BASE / b)
        q.put('Открываю анкету договоров («%s») в LibreOffice…' % b)
        dlg.destroy()

    def собрать():
        b = book_path(BOOKS_DOG)
        if not b:
            messagebox.showinfo('Нет анкеты', 'Сначала создайте и заполните анкету договоров.')
            return
        set_buttons(True)
        show_progress('Сбор договоров…')
        q.put('')
        q.put('── СБОР ДОГОВОРОВ (окно «Договоры») ──')
        q.put('Файл должен быть сохранён в LibreOffice (Ctrl+S)!')

        def done(c):
            out = sorted(BASE.glob('output-contracts/202*'))
            rpt = BASE / 'validation-report-contracts.txt'
            if rpt.exists():
                q.put('──── ОТЧЁТ-СВЕРКИ (ДОГОВОРЫ) ────')
                q.put(rpt.read_text(encoding='utf-8', errors='replace'))
                q.put('───────────────────────')
            if c == 0 and out:
                platform_utils.reveal(out[-1])
                q.put('Готово. Папка с договорами: ' + out[-1].name)
        run_cmd([sys.executable, 'generate_contracts.py', b], '', done)

    def создать():
        set_buttons(True)
        show_progress('Создание анкеты договоров…')
        q.put('')
        q.put('── СОЗДАНИЕ АНКЕТЫ ДОГОВОРОВ ──')

        def done(c):
            if (BASE / 'contracts-workbook.xlsx').exists():
                q.put('Анкета создана. Заполните лист «Договоры» (строка = договор).')
            refill()
            if c == 0:
                открыть_анкету()
        run_cmd([sys.executable, 'create_contract_workbook.py'], 'Анкета создана.', done)

    def последний_выпуск():
        out = sorted(BASE.glob('output-contracts/202*'))
        if out:
            platform_utils.reveal(out[-1])
            q.put('Папка последнего сбора: %s' % out[-1].name)
        else:
            messagebox.showinfo('Пока пусто', 'Договоры ещё не собирались — нажмите «⚖️ Собрать договоры».')

    def открыть_шаблон(вид):
        """Собрать бланк-шаблона договора в .docx и открыть его в Word."""
        set_buttons(True)
        show_progress('Сборка шаблона: %s…' % вид)
        q.put('')
        q.put('── ШАБЛОН ДОГОВОРА: %s ──' % вид)

        def done(c):
            if c == 0:
                out = BASE / 'contract-templates' / platform_utils.ascii_filename('template-%s.docx' % вид)
                if out.exists():
                    platform_utils.open_in_libreoffice(out)
                    q.put('Шаблон открыт в Word: %s' % out.name)
                    q.put('Поля [в квадратных скобках] заполняются при сборе договора из анкеты.')
                else:
                    q.put('! Файл шаблона не найден: %s' % out)
        run_cmd([sys.executable, 'generate_contracts.py', '--шаблон', вид], '', done)

    def посмотреть_шаблон(_=None):
        """Шаблон договора — в формате Word (.docx), до заполнения анкеты."""
        try:
            import contract_texts as тд
        except ImportError:
            messagebox.showerror('Нет шаблонов',
                                 'Файл «contract_texts.py» не найден рядом с программой.')
            return

        sel = lb.curselection()
        вид = next((v for v in тд.ШАБЛОНЫ if v in lb.get(sel[0])), None) if sel else None
        if вид is not None:
            открыть_шаблон(вид)
            return
        # ничего не выделено (или строка-подсказка) — предложить выбрать вид
        ch = tk.Toplevel(dlg)
        ch.title('Шаблон какого договора показать?')
        ch.configure(bg=BG)
        ch.transient(dlg)
        ch.geometry('480x430')
        tk.Label(ch, text='Шаблон какого договора показать?', font=big,
                 bg=BG, fg=FG).pack(pady=(14, 8))
        tk.Label(ch, text='Шаблон откроется в Word — файлом .docx.',
                 font=mid, bg=BG, fg=MUTED, justify='center').pack(pady=(0, 6))
        lbv = tk.Listbox(ch, font=big, bg=SURFACE, fg=FG,
                         selectmode='browse', activestyle='none',
                         exportselection=False)
        lbv.pack(fill='both', expand=True, padx=16)
        for в in тд.ШАБЛОНЫ:
            lbv.insert('end', '  ' + в)
        lbv.selection_set(0)

        def открыть(_=None):
            выб = lbv.curselection()
            if выб:
                ch.destroy()
                открыть_шаблон(list(тд.ШАБЛОНЫ)[выб[0]])
        Button(ch, text='👁  Показать', font=big, bg=PRIMARY, fg='#FFFFFF',
                  activebackground=PRIMARY_HOVER, relief='flat', cursor='hand2',
                  padx=14, pady=7, command=открыть).pack(pady=(8, 12))
        lbv.bind('<Double-Button-1>', открыть)
        lbv.bind('<Return>', открыть)
        lbv.focus_set()
        ch.grab_set()

    btn_row = tk.Frame(dlg, bg=BG)
    btn_row.pack(fill='x', padx=20, pady=(16, 20))
    for index, (текст, ком, цвет) in enumerate((
        ('Создать анкету', создать, SURFACE_HOVER),
        ('Открыть анкету', открыть_анкету, SURFACE_HOVER),
        ('Шаблон в Word', посмотреть_шаблон, SURFACE_HOVER),
        ('Собрать договоры', собрать, PRIMARY),
        ('Последний сбор', последний_выпуск, SURFACE_HOVER),
        ('Обновить список', refill, SURFACE_HOVER),
    )):
        btn_row.columnconfigure(index % 3, weight=1, uniform='contract-actions')
        Button(btn_row, text=текст, font=mid, bg=цвет, fg='#FFFFFF',
                  activebackground=PRIMARY_HOVER if цвет == PRIMARY else '#30435B',
                  relief='flat', cursor='hand2', width=0,
                  padx=14, pady=10, command=ком).grid(row=index // 3, column=index % 3,
                                                    sticky='ew', padx=4, pady=4)
    lb.bind('<Double-Button-1>', посмотреть_шаблон)
    ent.focus_set()
    refill()
    dlg.grab_set()


sections = []
SECTION_INFO = {
    'РАЗДЕЛ ЖКУ': ('Жилищное право',
        'Судебные приказы и иски о взыскании задолженности за ЖКУ.',
        ('Выберите расчёт бухгалтерии и справки Ф-9. Данные будут перенесены в таблицу.',
         'Сверьте должников, реквизиты и суммы. Сохраните изменения в редакторе.',
         'Получите заявления и отчёт сверки. Проверьте документы перед подачей.')),
    'СЕМЕЙНОЕ ПРАВО': ('Семейное право',
        'Иски о расторжении брака, алиментах и месте жительства детей.',
        ('Загрузите паспорт, свидетельство о браке и документы детей.',
         'Проверьте распознанные данные и заполните обстоятельства дела.',
         'Сформируйте исковое заявление и проверьте итоговый документ.')),
    'ЗАЩИТА ПРАВ ПОТРЕБИТЕЛЯ': ('Права потребителей',
        'Иски о возврате денег, замене товара и устранении недостатков.',
        ('Выберите паспорт, претензию, чеки и другие документы по делу.',
         'Заполните обстоятельства спора и требования. Сверьте суммы и даты.',
         'Создайте исковое заявление с требованиями и расчётами.')),
}


def create_section(title, rows, default_open=False, word='таблица', books=BOOKS_JKU):
    display, description, details = SECTION_INFO[title]
    entry = {'state': {'open': False}, 'markers': []}

    def close_panel():
        entry['state']['open'] = False
        entry['frame'].pack_forget()
        button_colors(entry['btn'], SIDEBAR, SURFACE, MUTED)

    def open_panel():
        if entry['state']['open']:
            return
        for other in sections:
            if other is not entry:
                other['close_panel']()
        disarm_continue()
        entry['state']['open'] = True
        current_section.update(books=books, word=word, entry=entry)
        entry['frame'].pack(before=activity_panel, fill='x')
        button_colors(entry['btn'], '#20334D', '#293F5C', FG)
        page_title.config(text=display)
        page_description.config(text=description)
        refresh_status()

    btn = Button(tab_row, text=display, font=mid, bg=SIDEBAR, fg=MUTED,
                 activebackground=SURFACE, anchor='w', width=0,
                 cursor='hand2', padx=14, pady=13, command=open_panel)
    btn.pack(fill='x', pady=3)
    frame = tk.Frame(main_content, bg=BG)
    entry.update(btn=btn, frame=frame, close_panel=close_panel,
                 open_panel=open_panel, title=title)
    step_titles = ('Загрузить документы', 'Проверить данные', 'Собрать документы')
    action_titles = ('Выбрать файлы', 'Открыть ' + ('таблицу' if word == 'таблица' else 'анкету'),
                     'Создать документы')
    for index, (text, cmd, _) in enumerate(rows):
        frame.columnconfigure(index, weight=1, uniform='steps')
        card = tk.Frame(frame, bg=SURFACE, highlightbackground=BORDER,
                         highlightthickness=1, height=252)
        card.grid(row=0, column=index, sticky='nsew',
                  padx=(0 if index == 0 else 6, 0 if index == 2 else 6))
        card.grid_propagate(False)
        card.columnconfigure(0, weight=1)
        card.rowconfigure(2, weight=1)
        marker = tk.Label(card, text=f'{index + 1:02d}', font=(family, 11, 'bold'),
                           bg=SURFACE_HOVER, fg=ACCENT, padx=9, pady=5)
        marker.grid(row=0, column=0, sticky='w', padx=18, pady=(18, 14))
        entry['markers'].append(marker)
        label = tk.Label(card, text=step_titles[index], font=big, bg=SURFACE,
                          fg=FG, anchor='w', justify='left', wraplength=210)
        label.grid(row=1, column=0, sticky='ew', padx=18)
        detail = tk.Label(card, text=details[index], font=small, bg=SURFACE,
                           fg=MUTED, anchor='nw', justify='left', wraplength=210)
        detail.grid(row=2, column=0, sticky='new', padx=18, pady=(10, 12))

        def resize_card(event, labels=(label, detail)):
            for widget in labels:
                widget.config(wraplength=max(100, event.width - 38))
        card.bind('<Configure>', resize_card)

        def guarded(c=cmd):
            disarm_continue()
            c()
        action = Button(card, text=action_titles[index], font=mid,
                         bg=PRIMARY if index == 0 else SURFACE_HOVER, fg=FG,
                         activebackground=PRIMARY_HOVER if index == 0 else '#30435B',
                         cursor='hand2', padx=12, pady=10, width=0, command=guarded)
        action.grid(row=3, column=0, sticky='ew', padx=18, pady=(0, 18))
        buttons[text] = action
    sections.append(entry)
    if default_open:
        open_panel()
    return btn, frame


jku_steps = []    # цепочка этапов для кнопок «Следующий шаг»

jku_rows = [
    ('📋  1. Загрузить данные (Портянка + Ф-9)',
     make_load_data(
         'import_housing.py',
         '── ЗАГРУЗКА ДАННЫХ (Портянка + Ф-9) ──',
         'Данные импортированы. Дочистите таблицу (лист «Сверка импорта»), затем «Собрать документы».',
         ('ЛС распознаётся автоматически из текста каждой справки: один скан со многими справками — не проблема.',
          'Справки привязываются к своим ЛС (лист «Ф-9 справки» в книге) — в документ они не вставляются, их прикладывают вручную.',
          'Если автопривязка не удалась: впишите ЛС в «f9-index.xlsx» и повторите загрузку.'),
         on_done=step_hook(jku_steps, 0)),
     PRIMARY),
    ('✏️  2. Открыть таблицу в редакторе',
     make_open_editor(BOOKS_JKU, on_done=step_hook(jku_steps, 1)), SURFACE_HOVER),
    ('⚖️  3. Собрать документы',
     make_build_docs(BOOKS_JKU, 'generate.py', 'validation-report.txt', 'output/202*',
                     'СБОР ДОКУМЕНТОВ', on_done=step_hook(jku_steps, 2)),
     PRIMARY),
]
jku_steps.extend(c for _, c, _ in jku_rows)
create_section('РАЗДЕЛ ЖКУ', jku_rows, default_open=True, word='таблица', books=BOOKS_JKU)

sem_steps = []    # цепочка этапов для кнопок «Следующий шаг»

sem_rows = [
    ('📋  1. Загрузить сканы документов',
     make_load_data(
         'import_family.py',
         '── ЗАГРУЗКА ДОКУМЕНТОВ (паспорт, свидетельства) ──',
         'Данные распознаны и вписаны в анкету (только пустые ячейки). Сверьте с оригиналами документов.',
         ('Форматы: pdf, jpg/png, Word (doc/docx/rtf/odt), Excel (xlsx/xls/ods), txt/csv.',
          'Ожидаемые имена файлов: «паспорт истца», «свидетельство о браке», «рождение-1/2/3».',
          'Лист «Сканы» в анкете показывает, что извлечено из каждого файла.'),
         src_dir=BASE / 'scans-family',
         on_done=step_hook(sem_steps, 0)),
     PRIMARY),
    ('✏️  2. Открыть анкету в редакторе',
     make_open_editor(BOOKS_SEM, 'анкета', on_done=step_hook(sem_steps, 1)), SURFACE_HOVER),
    ('⚖️  3. Собрать исковое заявление',
     make_build_docs(BOOKS_SEM, 'generate_family_claim.py', 'validation-report-family.txt',
                     'output-family/202*', 'СБОР СЕМЕЙНОГО ИСКА',
                     on_done=step_hook(sem_steps, 2)),
     PRIMARY),
]
sem_steps.extend(c for _, c, _ in sem_rows)
create_section('СЕМЕЙНОЕ ПРАВО', sem_rows, default_open=False, word='анкета', books=BOOKS_SEM)

potr_steps = []   # цепочка этапов для кнопок «Следующий шаг»

potr_rows = [
    ('📋  1. Загрузить сканы документов',
     make_load_data(
         'import_consumer.py',
         '── ЗАГРУЗКА ДОКУМЕНТОВ (паспорт, претензии, чеки) ──',
         'Паспорт распознан и вписан в анкету (только пустые ячейки); тексты претензий и чеков — на листе «Сканы».',
         ('Форматы: pdf, jpg/png, Word (doc/docx/rtf/odt), Excel (xlsx/xls/ods), txt/csv.',
          'Ожидаемые имена файлов: «паспорт истца», «претензия…», «чек…», договор.',
          'Описательную часть иска заполняйте на листе «Проблематика»: строка = абзац.'),
         src_dir=BASE / 'scans-consumer',
         on_done=step_hook(potr_steps, 0)),
     PRIMARY),
    ('✏️  2. Открыть анкету в редакторе',
     make_open_editor(BOOKS_POTR, 'анкета', on_done=step_hook(potr_steps, 1)), SURFACE_HOVER),
    ('⚖️  3. Собрать исковое заявление',
     make_build_docs(BOOKS_POTR, 'generate_consumer_claim.py',
                     'validation-report-consumer.txt',
                     'output-consumer/202*', 'СБОР ИСКА (ПОТРЕБИТЕЛЬ)',
                     on_done=step_hook(potr_steps, 2)),
     PRIMARY),
]
potr_steps.extend(c for _, c, _ in potr_rows)
create_section('ЗАЩИТА ПРАВ ПОТРЕБИТЕЛЯ', potr_rows, default_open=False,
               word='анкета потребителя', books=BOOKS_POTR)

contracts_nav = Button(tab_row, text='Договоры  ↗', font=mid, bg=SIDEBAR, fg=MUTED,
                       activebackground=SURFACE, anchor='w', width=0,
                       cursor='hand2', padx=14, pady=13, command=окно_договоров)
contracts_nav.pack(fill='x', pady=3)


root.mainloop()
