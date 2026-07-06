# -*- mode: python ; coding: utf-8 -*-
"""
TallyBridge PyInstaller spec file.

Builds a single Windows executable with all dependencies bundled.
Usage: pyinstaller tallybridge.spec
"""

import sys
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# Collect CustomTkinter assets (themes, images)
ctk_data = collect_data_files('customtkinter')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=ctk_data,
    hiddenimports=[
        'openpyxl',
        'customtkinter',
        'src',
        'src.parser',
        'src.generator',
        'src.gui',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='TallyBridge',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window (GUI app)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon path here if you have one: icon='assets/icon.ico'
)
