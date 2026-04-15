# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for VN Pathfinder

import sys
from pathlib import Path

block_cipher = None

a = Analysis(
    ['vn_pathfinder.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('assets', 'assets'),          # logo.png + logo.ico
    ],
    hiddenimports=[
        'PIL._tkinter_finder',
        'PIL.Image',
        'PIL.ImageTk',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        # Metadata scraping — bundled but gracefully absent at runtime if missing
        'curl_cffi',
        'curl_cffi.requests',
        'bs4',
        'lxml',
        'lxml.etree',
        'lxml._elementpath',
        # pywebview — in-app browser for itch.io login
        'webview',
        'webview.platforms.winforms',
        'webview.http',
        'bottle',
        # pythonnet / clr — needed by pywebview edgechromium backend
        'clr',
        'clr_loader',
        'pythonnet',
        # multiprocessing support for pywebview subprocess
        'multiprocessing',
        'multiprocessing.process',
        'multiprocessing.queues',
        'multiprocessing.reduction',
        'multiprocessing.popen_spawn_win32',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'numpy', 'scipy', 'matplotlib', 'pandas',
        'PyQt5', 'PyQt6', 'PySide2', 'PySide6',
        'IPython', 'jupyter',
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
    name='VNPathfinder',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/logo.ico',
    version_file=None,
)
