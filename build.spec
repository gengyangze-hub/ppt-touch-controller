# -*- mode: python ; coding: utf-8 -*-
# PyInstaller build spec for PPT Touch Controller

import sys
from pathlib import Path

# Project root
PROJECT_DIR = Path(SPECPATH)

a = Analysis(
    [str(PROJECT_DIR / 'src' / 'main.py')],
    pathex=[str(PROJECT_DIR / 'src')],
    binaries=[],
    datas=[
        (str(PROJECT_DIR / 'src' / 'resources'), 'resources'),
    ],
    hiddenimports=[
        'win32com',
        'win32com.client',
        'win32com.client.dynamic',
        'pythoncom',
        'win32gui',
        'win32con',
        'win32api',
        'winreg',
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
        'PySide6.QtNetwork',
        'pptx',
        'pptx.parts',
        'PIL',
        'PIL.Image',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'numpy',
        'scipy',
        'PySide6.QtWebEngine',
        'PySide6.QtWebEngineCore',
        'PySide6.QtWebEngineWidgets',
        'PySide6.QtWebChannel',
        'PySide6.QtMultimedia',
        'PySide6.QtMultimediaWidgets',
        'PySide6.Qt3DAnimation',
        'PySide6.Qt3DCore',
        'PySide6.Qt3DExtras',
        'PySide6.Qt3DInput',
        'PySide6.Qt3DLogic',
        'PySide6.Qt3DRender',
        'PySide6.QtBluetooth',
        'PySide6.QtNfc',
        'PySide6.QtPositioning',
        'PySide6.QtQuick',
        'PySide6.QtQuickWidgets',
        'PySide6.QtSensors',
        'PySide6.QtSerialPort',
        'PySide6.QtSql',
        'PySide6.QtSvg',
        'PySide6.QtSvgWidgets',
        'PySide6.QtTest',
        'PySide6.QtXml',
    ],
    noarchive=False,
    optimize=2,
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='PPTTouchController',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # Windowed app (no console)
    disable_windowed_traceback=False,
    # icon=[str(PROJECT_DIR / 'src' / 'resources' / 'icons' / 'app.ico')],
    uac_admin=False,
)
