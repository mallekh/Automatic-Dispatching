# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files, collect_submodules
import importlib.util


hiddenimports = [
    # Core runtime used by converter + UI.
    'pandas',
    'numpy',
    'openpyxl',
    'pdfplumber',
    'customtkinter',
    # Keep ML/runtime dependencies bundled for future logic growth.
    'scipy',
    'sklearn',
]
if importlib.util.find_spec('tkinterdnd2') is not None:
    hiddenimports.append('tkinterdnd2')
hiddenimports += collect_submodules('sklearn')
hiddenimports += collect_submodules('scipy')

datas = []
datas += collect_data_files('pandas')
datas += collect_data_files('openpyxl')
datas += collect_data_files('customtkinter')
if importlib.util.find_spec('tkinterdnd2') is not None:
    datas += collect_data_files('tkinterdnd2')


a = Analysis(
    ['app.py'],
    pathex=['.'],
    binaries=[],
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
    name='app',
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
