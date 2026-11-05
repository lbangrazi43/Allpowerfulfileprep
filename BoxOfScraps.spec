# -*- mode: python ; coding: utf-8 -*-
import base64
import io
import os

from PIL import Image, ImageOps
from PyInstaller.utils.hooks import collect_all, copy_metadata

datas = [
    # All build inputs are repo-relative on purpose: an absolute path here (this
    # used to point into a Downloads folder) builds fine on one machine and
    # fails on every other one, including a fresh clone.
    ('icon.ico', '.'),
]
binaries = []
hiddenimports = ['win32timezone']

# Bundle pywin32 + Pillow, and the AI Preparation stack (markitdown and the
# pieces PyInstaller can't auto-detect: magika models, onnxruntime, extract_msg,
# pdfminer data). markitdown is imported lazily, so these must be collected
# explicitly or "Convert to Markdown" would fail only at runtime. The same
# applies to py7zr (.7z extraction), rarfile (.rar) and pillow_heif (.heic/.heif
# images) — all three are lazily imported behind an "install X" error, so
# without collecting them here the feature would only fail on a user's machine.
for _pkg in ('win32com', 'win32api', 'pywintypes', 'PIL',
             'markitdown', 'magika', 'onnxruntime', 'extract_msg', 'pdfminer',
             'msoffcrypto', 'olefile',    # Password Removal (open-password decryption)
             'py7zr', 'rarfile',          # Folder Unzipping: .7z / .rar archives
             'pillow_heif',               # HEIC/HEIF images (iPhone photos)
             # Updater TLS. _update_ssl_context imports it inside the function
             # and falls back to the stock context when it is missing — so a
             # build that quietly dropped it would still run, and the certificate
             # failures on TLS-inspecting networks it exists to fix would come
             # back with nothing to show they had. Collected explicitly so that
             # can't happen silently; its platform backends are chosen by a
             # sys.platform branch, which collect_all follows and a bare
             # hiddenimport would not.
             'truststore'):
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


# ── Splash ───────────────────────────────────────────────────────────────────
# The bootloader's splash is the app's only loading screen: it appears before
# any Python runs and stays up until the main window is painted. PyInstaller's
# script only shows a still image, so splash_anim.tcl is appended to it. That
# file holds the explanation; what's here just fills in what it needs.

# How long one more file costs to unpack, expressed as bytes of file content, so
# that the bar tracks time rather than file count: 55% of the bundle's bytes sit
# in 10 of its 3,400 files. An estimate until it is measured from a real launch.
SPLASH_BYTES_PER_FILE = 64 * 1024

# Share of the bar for unpacking, and where it may drift to while Python starts.
# File.py's _splash_stage calls carry on from these numbers. Also estimates until
# they are measured from a real launch.
SPLASH_EXTRACT_SHARE = 60
SPLASH_BOOT_CEILING = 75


def _png_base64(image):
    buf = io.BytesIO()
    image.save(buf, format='PNG')
    return base64.b64encode(buf.getvalue()).decode('ascii')


def _owl_sprite(height):
    """The owl as a navy silhouette, trimmed of its white margin."""
    gray = Image.open(os.path.join(SPECPATH, 'owl_source.png')).convert('L')
    gray = gray.crop(ImageOps.invert(gray).getbbox())
    owl = Image.new('RGBA', gray.size, '#073b5e')
    owl.putalpha(ImageOps.invert(gray))
    return owl.resize((max(1, owl.width * height // owl.height), height), Image.LANCZOS)


def _logo_sprite(height):
    logo = Image.open(os.path.join(SPECPATH, 'logo_b.png')).convert('RGBA')
    return logo.resize((max(1, logo.width * height // logo.height), height), Image.LANCZOS)


def _extraction_costs(toc):
    """Return (cost of every file the bootloader unpacks, as a total) and the
    Tcl list of the files that cost more than the default of 1."""
    total, expensive = 0, []
    for dest, src, _typecode in toc:
        cost = 1 + os.path.getsize(src) // SPLASH_BYTES_PER_FILE
        total += cost
        if cost > 1:
            expensive.append('{%s} %d' % (dest, cost))
    return total, ' '.join(expensive)


class AnimatedSplash(Splash):
    def generate_script(self):
        total, expensive = _extraction_costs(a.binaries + a.datas)
        with open(os.path.join(SPECPATH, 'splash_anim.tcl'), encoding='utf-8') as f:
            anim = f.read()
        for key, value in {
            '@EXTRACT_SHARE@': SPLASH_EXTRACT_SHARE,
            '@BOOT_CEILING@': SPLASH_BOOT_CEILING,
            '@FILE_COSTS@': expensive,
            '@COST_TOTAL@': total,
            '@OWL_PNG@': _png_base64(_owl_sprite(70)),
            '@LOGO_PNG@': _png_base64(_logo_sprite(60)),
        }.items():
            anim = anim.replace(key, str(value))

        script = super().generate_script() + '\n' + anim
        # PyInstaller keeps a copy on disk so the script that ran can be read.
        with open(self.script_name, 'w', encoding='utf-8') as f:
            f.write(script)
        return script


splash = AnimatedSplash(
    'splash_static.png',       # the blank frame the sprites run across
    binaries=a.binaries,
    datas=a.datas,
    text_pos=None,
    always_on_top=False,       # other windows may cover it while it loads
)

exe = EXE(
    pyz,
    a.scripts,
    splash,
    splash.binaries,
    a.binaries,
    a.datas,
    [],
    name='BoxOfScraps',
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
    icon=['icon.ico'],
)
