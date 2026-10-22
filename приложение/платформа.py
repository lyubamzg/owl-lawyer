# -*- coding: utf-8 -*-
"""Платформенный слой «Сова-юрист»: пути, открытие файлов, кодировки, шрифты.

Все отличия Linux/Windows собраны здесь — остальным скриптам достаточно
импортировать этот модуль. Только стандартная библиотека (python 3.9+).
"""
import os
import shutil
import subprocess
import sys
from pathlib import Path

IS_WIN = sys.platform.startswith('win')
BASE = Path(__file__).resolve().parent


# ---------------------------------------------------------------- иконка приложения
def icon_path():
    """Иконка окна: сначала рядом с приложением (переносимый пакет),
    затем старое место ~/.local/share/icons (прошлые установки)."""
    new = BASE / 'иконки' / 'сова-юрист.png'
    if new.is_file():
        return new
    return Path.home() / '.local/share/icons/сова-юрист.png'


# ---------------------------------------------------------------- LibreOffice
def soffice_cmd():
    """Полный путь к LibreOffice на текущей системе."""
    if not IS_WIN:
        for c in ('soffice', 'libreoffice'):
            p = shutil.which(c)
            if p:
                return p
        return 'soffice'  # прежнее поведение: пусть ошибка будет привычной
    for p in (r'C:\Program Files\LibreOffice\program\soffice.exe',
              r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'):
        if Path(p).exists():
            return p
    return 'soffice.exe'  # может сработать, если LibreOffice добавлен в PATH


def open_in_libreoffice(path):
    subprocess.Popen([soffice_cmd(), str(path)], cwd=str(BASE),
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


# ---------------------------------------------------------------- открыть папку/файл руками ОС
def reveal(path):
    if IS_WIN:
        os.startfile(str(path))  # noqa: S606
    else:
        subprocess.Popen(['xdg-open', str(path)],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


# ---------------------------------------------------------------- дочерние процессы
def child_env():
    """Окружение для дочерних python-процессов: на Windows — UTF-8 вывод."""
    env = dict(os.environ)
    if IS_WIN:
        env.setdefault('PYTHONUTF8', '1')
        env.setdefault('PYTHONIOENCODING', 'utf-8')
    return env


def popen(cmd, **kw):
    """subprocess.Popen с правильной кодировкой вывода на обеих системах."""
    if kw.get('text') and 'encoding' not in kw:
        kw['encoding'] = 'utf-8'
        kw['errors'] = 'replace'
    kw.setdefault('env', child_env())
    return subprocess.Popen(cmd, **kw)  # noqa: S603


# ---------------------------------------------------------------- GUI-шрифты
UI_FONT = ['DejaVu Sans', 'Segoe UI', 'Arial', 'Helvetica']
UI_FONT_MONO = ['DejaVu Sans Mono', 'Consolas', 'Courier New', 'Courier']


def pick_family(tkfont, candidates):
    """Первый из кандидатов, реально установленный в системе."""
    try:
        have = set(tkfont.families())
    except Exception:
        return candidates[0]
    for f in candidates:
        if f in have:
            return f
    return candidates[0]


# ---------------------------------------------------------------- предпроверка инструментов
TOOLS = {
    'pdftoppm': 'poppler-utils (Linux) или poppler-windows (Windows)',
    'tesseract': 'tesseract-ocr + языковой пакет Russian',
    'soffice': 'LibreOffice',
}


def missing_tools():
    """Список неустановленных внешних программ (для понятной ошибки)."""
    missing = []
    for name, hint in TOOLS.items():
        if name == 'soffice':
            ok = bool(shutil.which('soffice') or shutil.which('libreoffice'))
            if IS_WIN and not ok:
                ok = Path(soffice_cmd()).exists()
        else:
            ok = bool(shutil.which(name))
        if not ok:
            missing.append(f'{name} ({hint})')
    return missing


def fix_console():
    """На Windows консоль часто cp866 — переводим вывод python в UTF-8."""
    if not IS_WIN:
        return
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding='utf-8', errors='replace')
        except Exception:  # noqa: BLE001
            pass