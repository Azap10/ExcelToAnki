# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, collect_submodules, copy_metadata

datas = [('data/cedict_ts.u8', 'data')]
binaries = []
hiddenimports = ['paddleocr', 'pymupdf']
datas += collect_data_files('paddleocr')
datas += collect_data_files('paddlex')
datas += collect_data_files('paddle')
datas += copy_metadata('paddleocr')
datas += copy_metadata('paddlex')
datas += copy_metadata('paddlepaddle')
binaries += collect_dynamic_libs('paddle')
binaries += collect_dynamic_libs('pymupdf')
hiddenimports += collect_submodules('paddlex.inference.pipelines.ocr')
hiddenimports += collect_submodules('paddlex.inference.models.text_detection')
hiddenimports += collect_submodules('paddlex.inference.models.text_recognition')
hiddenimports += collect_submodules('paddlex.inference.models.common')


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='FlashcardFactory',
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
)
