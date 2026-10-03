# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, copy_metadata

block_cipher = None

# --- Сбор данных, метаданных и библиотек для PyTorch и зависимостей ---
datas = []
binaries = []

# Собираем все DLL-библиотеки из torch и его зависимостей (mkl, intel-openmp)
# Это ключевой момент для решения проблемы с shm.dll
binaries += collect_dynamic_libs('torch')
binaries += collect_dynamic_libs('mkl')
binaries += collect_dynamic_libs('intel_openmp')

# Собираем данные и метаданные для torch, transformers и sentence_transformers
datas += collect_data_files('torch')
datas += copy_metadata('torch')
datas += collect_data_files('transformers')
datas += copy_metadata('transformers')
datas += collect_data_files('sentence_transformers')
datas += copy_metadata('sentence_transformers')


datas += [
    ('Image', 'Image'),
    ('Image_White', 'Image_White'),
    ('Image_settings', 'Image_settings'),
    ('Image_Context_menu', 'Image_Context_menu'),
    ('Image_Context_menu_White', 'Image_Context_menu_White'),
    ('Icon', 'Icon'),
    ('Sound', 'Sound'),
    ('Documentation', 'Documentation'),
    ('Screenshots', 'Screenshots'),
    ('plugins', 'plugins'),
    ('vendor', 'vendor'),
    ('LICENSE', '.'),
    ('LICENSE.ru', '.'),
    ('dist/Echos_reader.exe', '.'),
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,  # <-- Собранные DLL от torch, mkl, intel-openmp
    datas=datas,        # <-- Собранные данные и метаданные
    hiddenimports=[
        'PySide6.QtPrintSupport',
        'PySide6.QtSvg',
        'PySide6.QtNetwork',
        'PySide6.QtMultimedia',
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
        'torch',
        'torch._C',
        'torch.backends.cudnn',
        'transformers',
        'sentence_transformers',
        'mkl',
        'intel_openmp',
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
    [],
    exclude_binaries=True,
    name='Echos',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='Icon/Icon.ico',
    version='version_info.txt',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='Echos',
)
