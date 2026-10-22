#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Рабочее место юриста: таблица → редактор → документы.

Вкладки-разделы в один ряд (ЖКУ, Семейное право, Защита прав потребителя);
панель шагов скрыта до клика и выпадает вниз под вкладками.
  1. Загрузить данные (Портянка + сканы Ф-9)
  2. Открыть таблицу в LibreOffice для редактирования
  3. Собрать документы
После каждого успешного этапа появляется кнопка «ПРОДОЛЖИТЬ» — она выполняет
следующий этап раздела; можно идти пошагово или пользоваться кнопками самому.
"""
import re
import subprocess
import threading
import queue
import time
import sys
from pathlib import Path

import платформа

import tkinter as tk
from tkinter import messagebox, filedialog, font as tkfont
from tkinter import ttk

BASE = Path(__file__).resolve().parent
ICON = платформа.icon_path()
BOOKS_JKU = ['Книга-ввода (портянка).xlsx', 'Книга-ввода.xlsx']
BOOKS_SEM = ['Анкета-бракоразводные.xlsx']
BOOKS_POTR = ['Анкета-потребительские.xlsx']
BOOKS_DOG = ['Книга-договоры.xlsx']

BG = '#2B3A55'      # тёмно-синий фон
FG = '#ECEFF4'      # светлый текст
ACC = '#8A6D3B'     # совиный коричневый
BTN_BG = '#4C5B7A'

root = tk.Tk()
root.title('Рабочее место юриста')
root.geometry('1100x700')
root.configure(bg=BG)

# ─── Фон с орнаментом из совиных голов (иконка приложения) ───
ORNAMENT = BASE / 'иконки' / 'фон-орнамент.png'
bg_canvas = tk.Canvas(root, bg=BG, highlightthickness=0)
bg_canvas.place(x=0, y=0, relwidth=1, relheight=1)


def paint_ornament(canvas, w, h):
    """Выстилает фон плиткой с совами (вызывается при изменении окна)."""
    canvas.delete('orn')
    if w < 2 or h < 2 or not ORNAMENT.exists():
        return
    tile = tk.PhotoImage(master=canvas, file=str(ORNAMENT))
    canvas.orn_tile = tile  # ссылка — иначе фотография соберётся сборщиком мусора
    tw, th = tile.width(), tile.height()
    for y in range(0, h + th, th):
        for x in range(0, w + tw, tw):
            canvas.create_image(x, y, anchor='nw', image=tile, tags='orn')

try:
    root.tk.call('tk', 'scaling', 1.2)
except Exception:
    pass

try:
    if ICON.exists():
        _img = tk.PhotoImage(file=str(ICON))
        root.iconphoto(True, _img)
except Exception:
    pass

big = tkfont.Font(family=платформа.pick_family(tkfont, платформа.UI_FONT), size=13, weight='bold')
mid = tkfont.Font(family=платформа.pick_family(tkfont, платформа.UI_FONT), size=11)
logf = tkfont.Font(family=платформа.pick_family(tkfont, платформа.UI_FONT_MONO), size=10)

# Шапку (заголовок, подзаголовок, статус) и подпись рисуем текстом прямо
# на фоновом canvas — надписи лежат на орнаменте, без собственных рамок.
t_item = bg_canvas.create_text(
    0, 0, text='РАБОЧЕЕ МЕСТО ЮРИСТА', justify='center', fill=FG,
    font=tkfont.Font(family=big.actual('family'), size=18, weight='bold'), tags='chrome')
status_item = bg_canvas.create_text(0, 0, text='', justify='center',
                                    fill='#B9E28C', font=mid, tags='chrome')
foot_item = bg_canvas.create_text(0, 0, justify='center', fill='#76819A',
                                  font=mid, tags='chrome')


def layout_chrome(w, h):
    """Расставляет шапку и подпись по центру при изменении окна."""
    cx = w / 2
    y = 26
    for item, dy in ((t_item, 34), (status_item, 26)):
        bg_canvas.coords(item, cx, y)
        y += dy
    bg_canvas.coords(foot_item, cx, h - 16)


def book_path(books):
    for b in books:
        if (BASE / b).exists():
            return b
    return None


def refresh_status(books=BOOKS_JKU, word='таблица'):
    b = book_path(books)
    msg = ('Текущая ' + word + ': ' + b if b
           else word.capitalize() + ' не найдена — нажмите «Загрузить данные»')
    bg_canvas.itemconfig(status_item, text=msg)


refresh_status()

# Раскладка шапки + орнамент при первом показе и любом изменении окна.
bg_canvas.bind('<Configure>', lambda e: (paint_ornament(bg_canvas, e.width, e.height),
                                         layout_chrome(e.width, e.height)))

# Что-то из необходимого не установлено? Сообщаем сразу, а не при нажатии кнопки.
_missing = платформа.missing_tools()
if _missing:
    tk.Tk().withdraw()
    messagebox.showwarning(
        'Не всё установлено',
        'Для полной работы Сова-юриста не хватает:\n\n  • '
        + '\n  • '.join(_missing)
        + '\n\nУстановите по README-файлам пакета Сова-юрист и перезапустите.'
          '\nБез этого кнопки, зависящие от программы, работать не будут.')

buttons = {}
log = tk.Text(root, bg='#1B2233', fg=FG, insertbackground=FG, font=logf,
              wrap='word', state='disabled', height=14)
log.pack(fill='both', expand=True, padx=18, pady=(6, 8))

prog_frame = tk.Frame(root, bg=BG)
prog_label = tk.Label(prog_frame, font=big, bg=BG, fg=FG, anchor='w')
prog_label.pack(fill='x')
prog_bar = ttk.Progressbar(prog_frame, orient='horizontal', length=560,
                           mode='indeterminate')
prog_bar.pack(fill='x', pady=(2, 4))

# ─── Ряд вкладок-разделов: кнопки в одну строку, панели выпадают вниз ───
tab_row = tk.Frame(root, bg=BG)
tab_row.pack(before=log, fill='x', padx=26, pady=(78, 0))

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
    prog_frame.pack(before=log, fill='x', padx=26, pady=(14, 4))


def hide_progress():
    prog_frame.pack_forget()
    prog_bar.stop()


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


# ─── Пошаговый режим: после каждого этапа — кнопка «ПРОДОЛЖИТЬ» ───
pending_next = {'cmd': None}


def arm_continue(next_cmd, note=''):
    """Показать «ПРОДОЛЖИТЬ»; нажатие выполнит следующий этап раздела."""
    if not next_cmd:
        return
    pending_next['cmd'] = next_cmd
    cont_btn.config(state='normal')
    cont_btn.pack(before=log, padx=26, pady=(0, 6))
    bg_canvas.itemconfig(status_item,
                         text='✔  ' + (note or 'Этап выполнен') + ' — жмите зелёную «ПРОДОЛЖИТЬ»',
                         fill='#E7C46A')
    q.put('')
    q.put('── ЭТАП ВЫПОЛНЕН. Внизу, над журналом, зелёная кнопка «ПРОДОЛЖИТЬ» — нажмите её ──')
    q.put('── для перехода к следующему этапу этого раздела. ──')


def disarm_continue():
    """Убрать кнопку (новый этап запущен, или этап не удался)."""
    pending_next['cmd'] = None
    cont_btn.pack_forget()
    bg_canvas.itemconfig(status_item,
                         text='… выполняется этап — окно не закрывайте; результат придёт в журнал',
                         fill='#E7C46A')


def continue_next():
    nxt = pending_next['cmd']
    disarm_continue()
    if nxt:
        nxt()


def step_hook(steps, i):
    """Хук завершения этапа i: показать «ПРОДОЛЖИТЬ» для этапа i+1; после последнего — итог."""
    def hook(c=0):
        nxt = steps[i + 1] if i + 1 < len(steps) else None
        if nxt:
            arm_continue(nxt, f'Этап {i + 1}-й из {len(steps)} выполнен')
        else:
            bg_canvas.itemconfig(status_item,
                                 text='✔  Все этапы раздела выполнены — документы готовы',
                                 fill='#9FD98F')
            q.put('')
            q.put('✔  ВСЕ ЭТАПЫ РАЗДЕЛА ВЫПОЛНЕНЫ — документы готовы. '
                  'Проверьте их; папка с готовыми документами уже открыта справа.')
    return hook


cont_btn = tk.Button(root, text='▶   ПРОДОЛЖИТЬ', font=('DejaVu Sans', 13, 'bold'),
                     bg='#3F8A2B', fg='#FFFFFF', activebackground='#5EAD3A',
                     relief='groove', cursor='hand2', padx=34, pady=12,
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
    if not busy:
        hide_progress()


def run_cmd(cmd, done_msg, after=lambda c: None):
    """Запуск команды в фоне с выводом в журнал."""
    def work():
        try:
            proc = платформа.popen(cmd, cwd=BASE, stdout=subprocess.PIPE,
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
             font=mid, bg=BG, fg='#B9C2D8').pack(padx=18, pady=(0, 8))

    dirlabel = tk.Label(dlg, text='', font=mid, bg=BG, fg='#B9E28C')
    dirlabel.pack(fill='x', padx=18)

    body = tk.Frame(dlg, bg=BG)
    body.pack(fill='both', expand=True, padx=18, pady=6)
    lb = tk.Listbox(body, selectmode='extended', exportselection=False,
                    font=mid, bg='#22304A', fg=FG, activestyle='dotbox',
                    highlightthickness=1, highlightbackground='#4C5B7A')
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
    tk.Button(nav, text='↑  Папка выше', font=mid, bg='#3D4C6B', fg=FG,
              activebackground='#4C5B7A', relief='flat', cursor='hand2',
              padx=12, pady=5, command=up_dir).pack(side='left', pady=4)
    tk.Button(nav, text='📁  Другая папка…', font=mid, bg='#3D4C6B', fg=FG,
              activebackground='#4C5B7A', relief='flat', cursor='hand2',
              padx=12, pady=5, command=switch_dir).pack(side='left', padx=(8, 0), pady=4)

    sel = tk.Frame(dlg, bg=BG)
    sel.pack(fill='x', padx=18)
    tk.Button(sel, text='Выделить все', font=mid, bg='#3D4C6B', fg=FG,
              activebackground='#4C5B7A', relief='flat', cursor='hand2',
              padx=12, pady=5, command=select_all).pack(side='left', pady=4)
    tk.Button(sel, text='Снять выделение', font=mid, bg='#3D4C6B', fg=FG,
              activebackground='#4C5B7A', relief='flat', cursor='hand2',
              padx=12, pady=5, command=clear_sel).pack(side='left', padx=(8, 0), pady=4)

    tk.Button(dlg, text='Загрузить выбранные', font=big, bg='#2B8A47', fg='#FFFFFF',
              activebackground='#3AA45B', relief='flat', cursor='hand2',
              padx=16, pady=8, command=ok).pack(fill='x', padx=18, pady=(10, 4))
    tk.Button(dlg, text='Отмена', font=big, bg='#4C5B7A', fg='#FFFFFF',
              activebackground='#5A6C8E', relief='flat', cursor='hand2',
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
        платформа.open_in_libreoffice(BASE / b)
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
                платформа.reveal(out[-1])
                q.put('')
                q.put('Готово. Папка с документами: ' + out[-1].name)
            if on_done:
                on_done(c)

        run_cmd([sys.executable, script, b], '', done)
    return cmd


def make_btn(text, cmd, color, parent):
    b = tk.Button(parent, text=text, font=big, bg=color, fg='#FFFFFF',
                  activebackground=BTN_BG, relief='flat', cursor='hand2',
                  padx=18, pady=10, command=cmd)
    b.pack(fill='x', padx=10, pady=4)
    return b


# ─── Разделы: вкладки в один ряд; панель шагов выпадает вниз под вкладками ───
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

SEC_BG = '#33415C'   # цвет выпадающих панелей
# скан-разделы: те же документы можно загрузить сразу в несколько из них
EXTRA_LOADS = {'импорт_семейных.py': 'СЕМЕЙНОЕ ПРАВО',
               'импорт_потребительских.py': 'ЗАЩИТА ПРАВ ПОТРЕБИТЕЛЯ'}


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
                       selectcolor='#4C5B7A', activebackground=BG,
                       highlightthickness=0).pack(anchor='w', padx=34, pady=2)
    out = {'scripts': []}
    def ok_run():
        out['scripts'] = [sc for sc in order if chosen[sc].get()]
        dlg.destroy()
    def ok_skip():
        dlg.destroy()
    tk.Button(dlg, text='Загрузить выбранные разделы тоже', font=big, bg='#2B8A47',
              fg='#FFFFFF', activebackground='#3AA45B', relief='flat', cursor='hand2',
              padx=16, pady=8, command=ok_run).pack(fill='x', padx=22, pady=(12, 4))
    tk.Button(dlg, text='Только текущий раздел', font=big, bg='#4C5B7A', fg='#FFFFFF',
              activebackground='#5A6C8E', relief='flat', cursor='hand2',
              padx=16, pady=8, command=ok_skip).pack(fill='x', padx=22, pady=(0, 16))
    dlg.grab_set()
    root.wait_window(dlg)
    return out['scripts']
def договоры_список():
    """Строки анкеты договоров для окна: [(текст строки, ключ поиска)] или (None, [])."""
    path = BASE / 'Книга-договоры.xlsx'
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
    dlg.title('ДОГОВОРЫ — все договоры')
    dlg.configure(bg=BG)
    dlg.transient(root)
    dlg.resizable(True, True)
    dlg.geometry('840x600')

    tk.Label(dlg, text='Все договоры', font=big, bg=BG, fg=FG,
             justify='center').pack(padx=18, pady=(14, 0))
    tk.Label(dlg, text='Введите слово и нажмите «🔍 Найти» (или Enter) — список сожмётся до нужного '
                       'договора: ищет по виду, сторонам, предмету, №.\n'
                       'Двойной клик по договору — его шаблон в Word (.docx).',
             font=mid, bg=BG, fg='#B9C2D8', justify='center').pack(padx=18, pady=(2, 6))

    top = tk.Frame(dlg, bg=BG)
    top.pack(fill='x', padx=18, pady=(4, 6))
    tk.Label(top, text='🔍  Поиск:', font=big, bg=BG, fg=FG).pack(side='left')
    поиск = tk.StringVar()
    ent = tk.Entry(top, textvariable=поиск, font=big, bg='#1B2233', fg=FG,
                   insertbackground=FG, relief='flat')
    ent.pack(side='left', fill='x', expand=True, padx=(8, 0), ipady=5)
    btn_найти = tk.Button(top, text='🔍  Найти', font=big, bg='#3D7BD0', fg='#FFFFFF',
                          activebackground='#5A9AE8', relief='flat', cursor='hand2',
                          padx=16, pady=6, command=lambda: найти())
    btn_найти.pack(side='left', padx=(8, 0))

    midf = tk.Frame(dlg, bg=BG)
    midf.pack(fill='both', expand=True, padx=18, pady=4)
    sb = tk.Scrollbar(midf, orient='vertical')
    lb = tk.Listbox(midf, font=big, bg='#1B2233', fg=FG, selectmode='browse',
                    exportselection=False, activestyle='underline',
                    yscrollcommand=sb.set)
    lb.pack(side='left', fill='both', expand=True)
    sb.config(command=lb.yview)
    sb.pack(side='right', fill='y')

    счёт = tk.Label(dlg, text='', font=mid, bg=BG, fg='#B9E28C')
    счёт.pack(fill='x', padx=18)

    def refill(*_):
        lb.delete(0, 'end')
        path, rows = договоры_список()
        запрос = поиск.get().strip().lower()
        if path is None:
            lb.insert('end', '— Анкета «Книга-договоры.xlsx» ещё не создана —')
            lb.insert('end', '  создайте её кнопкой «➕ Создать анкету» ниже')
            счёт.config(text='Анкета не создана', fg='#E28C8C')
            return
        найдено = 0
        for текст, ключ in rows:
            if not запрос or запрос in ключ:
                lb.insert('end', текст)
                найдено += 1
        счёт.config(text='Найдено договоров: %d из %d' % (найдено, len(rows)),
                    fg='#B9E28C')

    ищем = {'идёт': False}

    def показать_результат():
        refill()
        btn_найти.config(state='normal')
        lb.config(bg='#1B2233')
        dlg.configure(cursor='')
        ищем['идёт'] = False

    def найти(*_):
        if ищем['идёт']:
            return
        ищем['идёт'] = True
        btn_найти.config(state='disabled')
        lb.config(bg='#232C42')          # список затемняется — видно, что идёт поиск
        dlg.configure(cursor='watch')    # курсор «часы»
        счёт.config(text='⌛  Ищу договоры в анкете…', fg='#E8C468')
        dlg.after(320, показать_результат)

    поиск.trace_add('write', refill)
    ent.bind('<Return>', найти)

    def открыть_анкету(_=None):
        b = book_path(BOOKS_DOG)
        if not b:
            messagebox.showinfo('Нет анкеты', 'Анкета ещё не создана — нажмите «➕ Создать анкету».')
            return
        платформа.open_in_libreoffice(BASE / b)
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
            out = sorted(BASE.glob('Выход-договоры/202*'))
            rpt = BASE / 'Отчёт-сверки (договоры).txt'
            if rpt.exists():
                q.put('──── ОТЧЁТ-СВЕРКИ (ДОГОВОРЫ) ────')
                q.put(rpt.read_text(encoding='utf-8', errors='replace'))
                q.put('───────────────────────')
            if c == 0 and out:
                платформа.reveal(out[-1])
                q.put('Готово. Папка с договорами: ' + out[-1].name)
        run_cmd([sys.executable, 'генерировать_договоры.py', b], '', done)

    def создать():
        set_buttons(True)
        show_progress('Создание анкеты договоров…')
        q.put('')
        q.put('── СОЗДАНИЕ АНКЕТЫ ДОГОВОРОВ ──')

        def done(c):
            if (BASE / 'Книга-договоры.xlsx').exists():
                q.put('Анкета создана. Заполните лист «Договоры» (строка = договор).')
            refill()
            if c == 0:
                открыть_анкету()
        run_cmd([sys.executable, 'создать_книгу_договоров.py'], 'Анкета создана.', done)

    def последний_выпуск():
        out = sorted(BASE.glob('Выход-договоры/202*'))
        if out:
            платформа.reveal(out[-1])
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
                out = BASE / 'Шаблоны-договоров' / ('Шаблон — %s.docx' % вид)
                if out.exists():
                    платформа.open_in_libreoffice(out)
                    q.put('Шаблон открыт в Word: %s' % out.name)
                    q.put('Поля [в квадратных скобках] заполняются при сборе договора из анкеты.')
                else:
                    q.put('! Файл шаблона не найден: %s' % out)
        run_cmd([sys.executable, 'генерировать_договоры.py', '--шаблон', вид], '', done)

    def посмотреть_шаблон(_=None):
        """Шаблон договора — в формате Word (.docx), до заполнения анкеты."""
        try:
            import договоры_тексты as тд
        except ImportError:
            messagebox.showerror('Нет шаблонов',
                                 'Файл «договоры_тексты.py» не найден рядом с программой.')
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
                 font=mid, bg=BG, fg='#B9C2D8', justify='center').pack(pady=(0, 6))
        lbv = tk.Listbox(ch, font=big, bg='#1B2233', fg=FG,
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
        tk.Button(ch, text='👁  Показать', font=big, bg='#3D7BD0', fg='#FFFFFF',
                  activebackground='#5A9AE8', relief='flat', cursor='hand2',
                  padx=14, pady=7, command=открыть).pack(pady=(8, 12))
        lbv.bind('<Double-Button-1>', открыть)
        lbv.bind('<Return>', открыть)
        lbv.focus_set()
        ch.grab_set()

    btn_row = tk.Frame(dlg, bg=BG)
    btn_row.pack(fill='x', padx=18, pady=(10, 14))
    for текст, ком, цвет in (
        ('➕ Создать анкету', создать, '#2B8A47'),
        ('✏️ Открыть анкету', открыть_анкету, '#6A4CB7'),
        ('👁 Шаблон в Word', посмотреть_шаблон, '#B07A2A'),
        ('⚖️ Собрать договоры', собрать, '#2B8A47'),
        ('📁 Последний сбор', последний_выпуск, '#2B68B7'),
        ('🔄 Обновить', refill, '#4C5B7A'),
    ):
        tk.Button(btn_row, text=текст, font=mid, bg=цвет, fg='#FFFFFF',
                  activebackground='#5A6C8E', relief='flat', cursor='hand2',
                  padx=14, pady=8, command=ком).pack(side='left', expand=True,
                                                     fill='x', padx=4)
    lb.bind('<Double-Button-1>', посмотреть_шаблон)
    ent.focus_set()
    refill()
    dlg.grab_set()


sections = []          # [запись раздела]; раскрыт максимум один (гармошка)


def create_section(title, rows, default_open=False, word='таблица', books=BOOKS_JKU):
    entry = {}
    state = entry.setdefault('state', {'open': False})

    def close_panel():
        state['open'] = False
        entry['frame'].pack_forget()
        entry['btn'].config(text='▶   ' + title)

    def open_panel():
        disarm_continue()   # открыт другой раздел — прошлая кнопка «ПРОДОЛЖИТЬ» уходит
        state['open'] = True
        # before=log: панель всегда встаёт прямо под рядом вкладок, журнал остаётся внизу
        entry['frame'].pack(before=log, fill='x', padx=26, pady=(4, 6))
        entry['btn'].config(text='▼   ' + title)
        refresh_status(books, word)

    def toggle():
        if state['open']:
            close_panel()
        else:
            for other in sections:   # вкладочный стиль: другое раскрытое закрывается
                if other is not entry and other['state']['open']:
                    other['close_panel']()
            open_panel()

    btn = tk.Button(tab_row, text=('▼   ' if default_open else '▶   ') + title,
                    font=big, bg=ACC, fg='#FFFFFF',
                    activebackground='#6A5430', relief='flat', cursor='hand2',
                    padx=10, pady=8, command=toggle)
    btn.pack(side='left', padx=(0, 8))
    frame = tk.Frame(root, bg=SEC_BG)
    entry.update({'btn': btn, 'frame': frame, 'close_panel': close_panel, 'title': title})
    for text, cmd, color in rows:
        def guarded(c=cmd):
            disarm_continue()    # запущен новый этап — прежняя кнопка не нужна
            c()
        buttons[text] = make_btn(text, guarded, color, frame)
    sections.append(entry)
    if default_open:
        open_panel()
    return btn, frame


jku_steps = []    # цепочка этапов для кнопок «ПРОДОЛЖИТЬ»

jku_rows = [
    ('📋  1. Загрузить данные (Портянка + Ф-9)',
     make_load_data(
         'импорт_портянки.py',
         '── ЗАГРУЗКА ДАННЫХ (Портянка + Ф-9) ──',
         'Данные импортированы. Дочистите таблицу (лист «Сверка импорта»), затем «Собрать документы».',
         ('ЛС распознаётся автоматически из текста каждой справки: один скан со многими справками — не проблема.',
          'Справки привязываются к своим ЛС (лист «Ф-9 справки» в книге) — в документ они не вставляются, их прикладывают вручную.',
          'Если автопривязка не удалась: впишите ЛС в «Ф-9 индекс.xlsx» и повторите загрузку.'),
         on_done=step_hook(jku_steps, 0)),
     '#2B68B7'),
    ('✏️  2. Открыть таблицу в редакторе',
     make_open_editor(BOOKS_JKU, on_done=step_hook(jku_steps, 1)), '#6A4CB7'),
    ('⚖️  3. Собрать документы',
     make_build_docs(BOOKS_JKU, 'generate.py', 'Отчёт-сверки.txt', 'Выход/202*',
                     'СБОР ДОКУМЕНТОВ', on_done=step_hook(jku_steps, 2)),
     '#2B8A47'),
]
jku_steps.extend(c for _, c, _ in jku_rows)
create_section('РАЗДЕЛ ЖКУ', jku_rows, default_open=False, word='таблица', books=BOOKS_JKU)

sem_steps = []    # цепочка этапов для кнопок «ПРОДОЛЖИТЬ»

sem_rows = [
    ('📋  1. Загрузить сканы документов',
     make_load_data(
         'импорт_семейных.py',
         '── ЗАГРУЗКА ДОКУМЕНТОВ (паспорт, свидетельства) ──',
         'Данные распознаны и вписаны в анкету (только пустые ячейки). Сверьте с оригиналами документов.',
         ('Форматы: pdf, jpg/png, Word (doc/docx/rtf/odt), Excel (xlsx/xls/ods), txt/csv.',
          'Ожидаемые имена файлов: «паспорт истца», «свидетельство о браке», «рождение-1/2/3».',
          'Лист «Сканы» в анкете показывает, что извлечено из каждого файла.'),
         src_dir=BASE / 'Сканы-семейные',
         on_done=step_hook(sem_steps, 0)),
     '#2B68B7'),
    ('✏️  2. Открыть анкету в редакторе',
     make_open_editor(BOOKS_SEM, 'анкета', on_done=step_hook(sem_steps, 1)), '#6A4CB7'),
    ('⚖️  3. Собрать исковое заявление',
     make_build_docs(BOOKS_SEM, 'генерировать_иск.py', 'Отчёт-сверки (семейные).txt',
                     'Выход-семейные/202*', 'СБОР СЕМЕЙНОГО ИСКА',
                     on_done=step_hook(sem_steps, 2)),
     '#2B8A47'),
]
sem_steps.extend(c for _, c, _ in sem_rows)
create_section('СЕМЕЙНОЕ ПРАВО', sem_rows, default_open=False, word='анкета', books=BOOKS_SEM)

potr_steps = []   # цепочка этапов для кнопок «ПРОДОЛЖИТЬ»

potr_rows = [
    ('📋  1. Загрузить сканы документов',
     make_load_data(
         'импорт_потребительских.py',
         '── ЗАГРУЗКА ДОКУМЕНТОВ (паспорт, претензии, чеки) ──',
         'Паспорт распознан и вписан в анкету (только пустые ячейки); тексты претензий и чеков — на листе «Сканы».',
         ('Форматы: pdf, jpg/png, Word (doc/docx/rtf/odt), Excel (xlsx/xls/ods), txt/csv.',
          'Ожидаемые имена файлов: «паспорт истца», «претензия…», «чек…», договор.',
          'Описательную часть иска заполняйте на листе «Проблематика»: строка = абзац.'),
         src_dir=BASE / 'Сканы-потребительские',
         on_done=step_hook(potr_steps, 0)),
     '#2B68B7'),
    ('✏️  2. Открыть анкету в редакторе',
     make_open_editor(BOOKS_POTR, 'анкета', on_done=step_hook(potr_steps, 1)), '#6A4CB7'),
    ('⚖️  3. Собрать исковое заявление',
     make_build_docs(BOOKS_POTR, 'генерировать_иск_потребительский.py',
                     'Отчёт-сверки (потребительские).txt',
                     'Выход-потребительские/202*', 'СБОР ИСКА (ПОТРЕБИТЕЛЬ)',
                     on_done=step_hook(potr_steps, 2)),
     '#2B8A47'),
]
potr_steps.extend(c for _, c, _ in potr_rows)
create_section('ЗАЩИТА ПРАВ ПОТРЕБИТЕЛЯ', potr_rows, default_open=False,
               word='анкета потребителя', books=BOOKS_POTR)

# Кнопка-раздел «📁 ДОГОВОРЫ» в ряду вкладок — того же размера, что и остальные
# разделы; открывает окно со всеми договорами и поиском по названию.
tk.Button(tab_row, text='📁   ДОГОВОРЫ',
          font=big, bg='#3D7BD0', fg='#FFFFFF',
          activebackground='#5A9AE8', relief='flat', cursor='hand2',
          padx=10, pady=8, command=окно_договоров).pack(side='left', padx=(0, 8))

bg_canvas.itemconfig(foot_item, text='Все данные обрабатываются только на этом компьютере.')


root.mainloop()