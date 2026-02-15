# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs

datas = []
datas += collect_data_files('certifi')

# Collect pywin32 DLLs needed for COM automation (win32com)
binaries = []
venv_site = os.path.join('venv', 'Lib', 'site-packages')
pywin32_sys32 = os.path.join(venv_site, 'pywin32_system32')
if os.path.isdir(pywin32_sys32):
    for f in os.listdir(pywin32_sys32):
        if f.endswith('.dll'):
            binaries.append((os.path.join(pywin32_sys32, f), '.'))


a = Analysis(
    ['scraper_enhanced.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=[
        'win32com',
        'win32com.client',
        'win32com.client.dynamic',
        'win32com.client.gencache',
        'pythoncom',
        'pywintypes',
        'win32api',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='EnergyPlanScraper',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
