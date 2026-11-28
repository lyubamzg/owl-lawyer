#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Флаги функций «Совы-юрист».

Разделы рабочего места можно включать и выключать, не правя код.
Флаги читаются в таком порядке (поздний перекрывает ранний):

  1. значения по умолчанию — DEFAULTS ниже;
  2. файл рядом с программой — features.json;
  3. пользовательский файл — ~/.config/sova-yurist/features.json;
  4. переменная окружения SOVA_FEATURES — JSON-строка, например
     SOVA_FEATURES='{"семейное_право": false}'.

Значение флага — true/false, а также «да/нет», «вкл/выкл», 1/0.
"""
import json
import os
from pathlib import Path

BASE = Path(__file__).resolve().parent

DEFAULTS = {
    'семейное_право': True,
    'защита_прав_потребителя': True,
    'договоры': True,
    'досудебные_претензии': True,
    'исполнительное_производство': True,
}

APP_FILE = BASE / 'features.json'
USER_FILE = (Path(os.environ.get('XDG_CONFIG_HOME') or Path.home() / '.config')
              / 'sova-yurist' / 'features.json')

_FALSE = {'нет', 'no', 'off', '0', 'выкл', 'n', 'false', 'выкл.'}
_TRUE = {'да', 'yes', 'on', '1', 'вкл', 'истина', 'true', 'y'}


def _as_bool(value, default=True):
    """«да/нет» → True/False; неизвестное значение — default."""
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    v = str(value).strip().lower()
    if not v:
        return default
    if v in _FALSE:
        return False
    if v in _TRUE:
        return True
    return default


def _read_json(path):
    """Прочитать JSON-объект; битый или отсутствующий файл — пустой словарь."""
    try:
        data = json.loads(Path(path).read_text(encoding='utf-8'))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def _load():
    flags = dict(DEFAULTS)
    flags.update(_read_json(APP_FILE))
    flags.update(_read_json(USER_FILE))
    env = os.environ.get('SOVA_FEATURES', '').strip()
    if env:
        try:
            data = json.loads(env)
        except ValueError:
            data = {}
        if isinstance(data, dict):
            flags.update(data)
    return flags


FLAGS = _load()


def enabled(name, default=True):
    """Включён ли флаг. Неизвестный флаг — default."""
    key = str(name or '').strip().lower()
    if key not in FLAGS:
        return default
    return _as_bool(FLAGS[key], default)


def write_user_flags(values):
    """Записать флаги в пользовательский файл (для галочки в настройках)."""
    USER_FILE.parent.mkdir(parents=True, exist_ok=True)
    USER_FILE.write_text(json.dumps(values, ensure_ascii=False, indent=2) + '\n',
                         encoding='utf-8')
