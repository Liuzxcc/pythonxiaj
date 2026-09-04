# -*- mode: python ; coding: utf-8 -*-
# Windows onefile 打包 spec
# hiddenimports 必须与 macOS spec（进度跟踪报表工具.spec）保持一致：
#   tkinter 全套 + threading + openpyxl + xlutils/xlutils.copy + --collect-submodules core
# 缺 xlutils 会让 exe 在调用 write_inplace（from xlutils.copy import copy）时崩溃。
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = [
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'threading',
    'openpyxl',
    'xlutils',
    'xlutils.copy',
]
hiddenimports += collect_submodules('core')


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
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
    name='进度跟踪报表工具',
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
    icon=['assets/app.ico'],
)
