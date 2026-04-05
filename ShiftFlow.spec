# -*- mode: python ; coding: utf-8 -*-
# WAŻNE: entry point to bootstrap.py, a NIE main.py.
# main.py jest dołączony jako plik danych — bootstrap exec()uje go.
# Przebudowy .exe wymagają TYLKO zmian w zależnościach Python/PySide6.
# Każdą zmianę kodu wystarczy wgrać na GitHub (version.txt + main.py).
import os

block_cipher = None

a = Analysis(
    ['bootstrap.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('main.py',          '.'),   # bundlowany fallback gdy brak sieci
        ('harmonogram.json', '.'),
        ('config.json',      '.'),
        ('icon.ico',         '.'),
        ('version.txt',      '.'),
    ],
    hiddenimports=[
        'google.genai',
        'google.generativeai',
        'qdarkstyle',
        'openpyxl',
        'docx',
        'dotenv',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ShiftFlow',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)
