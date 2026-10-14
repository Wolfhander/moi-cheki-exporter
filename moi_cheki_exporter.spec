# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec для «Мои Чеки — Экспорт и аналитика».

Собирает всё в один .exe (onefile). В бандл включаются:
  - food_template.xlsx  — шаблон листа «Продукты»

Личные базы НЕ включаются: на новом компьютере они создаются заново.
"""
import os
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# Папка со .spec'ом = корень проекта
PROJECT_DIR = os.path.abspath(os.path.dirname(SPEC))  # noqa: F821

datas = [
    (os.path.join(PROJECT_DIR, "food_template.xlsx"), "."),
]

# Личные базы и настройки никогда не включаются в распространяемую сборку.

datas += collect_data_files("customtkinter")

hidden_imports = [
    "customtkinter",
    "openpyxl",
    "openpyxl.chart",
    "openpyxl.chart.bar_chart",
    "openpyxl.chart.pie_chart",
    "openpyxl.chart.line_chart",
    "openpyxl.chart.label",
    "openpyxl.chart.plotarea",
    "openpyxl.styles",
    "openpyxl.utils",
    "requests",
    "sqlite3",
]

a = Analysis(
    ["main.py"],
    pathex=[PROJECT_DIR],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "matplotlib",
        "numpy",
        "pandas",
        "scipy",
        "PyQt5",
        "PyQt6",
        "PySide2",
        "PySide6",
        "tkinter.test",
        "test",
        "unittest",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="MoiCheki",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # GUI-режим, без консольного окна
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,               # можно указать .ico-файл при желании
)
