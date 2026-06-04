# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, copy_metadata

datas = [('C:\\Users\\5999\\Downloads\\files\\icon.ico', '.')]
binaries = []
hiddenimports = ['win32timezone']

# Bundle pywin32 + Pillow, and the AI Preparation stack (markitdown and the
# pieces PyInstaller can't auto-detect: magika models, onnxruntime, extract_msg,
# pdfminer data). markitdown is imported lazily, so these must be collected
# explicitly or "Convert to Markdown" would fail only at runtime.
for _pkg in ('win32com', 'win32api', 'pywintypes', 'PIL',
             'markitdown', 'magika', 'onnxruntime', 'extract_msg', 'pdfminer'):
    tmp_ret = collect_all(_pkg)
    datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# markitdown reads its own metadata at import time.
datas += copy_metadata('markitdown')


a = Analysis(
    ['File.py'],
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
    name='AllPowerfulFilePrep',
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
    icon=['C:\\Users\\5999\\Downloads\\files\\icon.ico'],
)
