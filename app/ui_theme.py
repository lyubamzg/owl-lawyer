# -*- coding: utf-8 -*-
"""Shared desktop palette and portable, focus-stable controls."""
from tkinter import ttk

BG = '#111923'
SIDEBAR = '#0D141E'
SURFACE = '#192433'
SURFACE_HOVER = '#243449'
BORDER = '#2B394B'
FG = '#EDF2F8'
MUTED = '#96A6BA'
ACCENT = '#759FEA'
PRIMARY = '#365F9B'
PRIMARY_HOVER = '#426FAA'
SUCCESS = '#8CBDA5'


def Button(master, **options):
    """Render the same muted palette on macOS, Windows, and Linux."""

    style = ttk.Style(master)
    # Native Aqua buttons ignore background colours. Use a separate theme so
    # all buttons (including those in dialogs) retain the application's palette.
    theme = 'sova'
    if theme not in style.theme_names():
        style.theme_create(theme, parent='clam', settings={
            'TScrollbar': {'configure': {
                'background': SURFACE_HOVER, 'troughcolor': SURFACE,
                'bordercolor': SURFACE, 'lightcolor': SURFACE_HOVER,
                'darkcolor': SURFACE_HOVER, 'arrowcolor': MUTED,
                'arrowsize': 12,
            }, 'map': {'background': [('active', BORDER), ('!active', SURFACE_HOVER)]}},
            'Horizontal.TProgressbar': {'configure': {
                'background': ACCENT, 'troughcolor': SURFACE,
                'bordercolor': SURFACE, 'lightcolor': ACCENT,
                'darkcolor': ACCENT, 'thickness': 5,
            }},
        })
    if style.theme_use() != theme:
        style.theme_use(theme)

    background = options.pop('bg', SURFACE_HOVER)
    foreground = options.pop('fg', FG)
    active = options.pop('activebackground', background)
    active_foreground = options.pop('activeforeground', foreground)
    font = options.pop('font', None)
    anchor = options.pop('anchor', 'center')
    padding = (options.pop('padx', 0), options.pop('pady', 0))
    options.pop('relief', None)
    button = ttk.Button(master, **options)
    name = f'Sova{str(button).replace(".", "_")}.TButton'
    style.configure(name, background=background, foreground=foreground,
                    font=font, padding=padding, borderwidth=0,
                    focusthickness=1, focuscolor=ACCENT, anchor=anchor)
    style.map(name,
              background=[('disabled', background), ('pressed', active),
                          ('active', active), ('!active', background)],
              foreground=[('disabled', '#64758A'),
                          ('pressed', active_foreground),
                          ('active', active_foreground),
                          ('!active', foreground)],
              relief=[('pressed', 'flat'), ('!pressed', 'flat')])
    button.configure(style=name)
    return button


def button_colors(button, background, active=None, foreground=FG):
    """Update navigation selection without changing the widget's command."""
    style = ttk.Style(button)
    name = button.cget('style')
    style.configure(name, background=background, foreground=foreground)
    style.map(name,
              background=[('disabled', background),
                          ('active', active or background),
                          ('!active', background)],
              foreground=[('disabled', '#64758A'), ('!disabled', foreground)])
