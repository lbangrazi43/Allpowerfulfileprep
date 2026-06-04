# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, copy_metadata

datas = [
    ('C:\\Users\\5999\\Downloads\\files\\icon.ico', '.'),
    ('owl_source.png', '.'),   # splash: owl sprite
    ('logo_b.png', '.'),       # splash: 'b' logo sprite
]
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

# Heavy scientific / ML / GUI-toolkit packages get dragged in transitively
# (e.g. torch hard-depends on sympy) but a file-conversion app never uses them.
# Excluding them shrinks the one-file archive by well over 100 MB, which is the
# main thing that has to be unpacked to temp on every launch — so the app opens
# noticeably faster. (numpy/pandas are kept: markitdown needs them for Excel/CSV.)
EXCLUDES = [
    'torch', 'torchvision', 'torchaudio', 'sympy', 'pygame',
    'matplotlib', 'scipy', 'IPython', 'notebook', 'jupyter',
    'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'tkinter.test', 'test',
]


a = Analysis(
    ['File.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

# Built-in bootloader splash: shown by the loader *during* the one-file unpack,
# before any Python runs — so the user sees the brand immediately instead of a
# delay. The app closes it (pyi_splash.close()) once its own splash is on screen.
splash = Splash(
    'splash_static.png',
    binaries=a.binaries,
    datas=a.datas,
    text_pos=None,
)

exe = EXE(
    pyz,
    a.scripts,
    splash,
    splash.binaries,
    a.binaries,
    a.datas,
    [],
    name='AllPowerfulFilePrep',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    # UPX compresses the bundled DLLs, but they must be decompressed on launch —
    # that adds startup time. Off = slightly larger file, faster to open.
    upx=False,
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
