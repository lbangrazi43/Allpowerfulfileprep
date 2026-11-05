# -*- mode: python ; coding: utf-8 -*-
import base64
import io
import os

from PIL import Image, ImageOps
from PyInstaller.utils.hooks import collect_all, copy_metadata
from tzlocal.windows_tz import win_tz

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
    # Removes the pkg_resources and setuptools runtime hooks from every launch.
    # Don't add 'distutils' to this list: the build dies inside PyInstaller's own
    # hook for it.
    'pkg_resources', 'setuptools', '_distutils_hack',
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

# ── Trim ─────────────────────────────────────────────────────────────────────
# A one-file build unpacks every bundled file on every launch while Defender scans
# each one, and that was most of the launch time. File count matters more than
# size (measured at about 0.9 ms per file and 9 ms per MB), so most of what comes
# out here is small. This must run before PYZ and the splash, which read these lists.
_DROP_PREFIXES = (
    # Named-zone data and message catalogues for Tcl's `clock`, used only by Tk's
    # demos. The `clock milliseconds` the splash calls doesn't need them.
    '_tcl_data/tzdata/', '_tcl_data/msgs/',
    # pytz's zone database. Only timezone-aware pandas would reach it, and reading
    # a spreadsheet produces plain timestamps.
    'pytz/zoneinfo/',
    # The offline speech engine, never called: markitdown transcribes through
    # recognize_google and flac-win32.exe, which stays. The non-Windows flac goes too.
    'speech_recognition/pocketsphinx-data/',
    'speech_recognition/flac-linux', 'speech_recognition/flac-mac',
    # AVIF isn't in IMAGE_EXTS, and Pillow copes with the module being absent.
    'PIL/_avif',
    # win32ui and friends only serve makepy's progress window and COM browsers.
    # The app drives Office through DispatchEx.
    'Pythonwin/', 'win32com/test/', 'win32com/HTML/', 'win32com/demos/',
    'olefile/doc/',
)
# Source copies collect_all puts on disk beside modules already compiled into the
# PYZ, plus build leftovers.
_DROP_SUFFIXES = ('.py', '.pyi', '.h', '.lib', '.md', 'py.typed')
# tzdata itself has to stay, because extract_msg reads .msg dates through tzlocal,
# but only the zones Windows can actually report are needed. A machine with a TZ
# variable naming any other zone would get an error converting a .msg to Markdown.
_KEEP_ZONES = {'tzdata/zoneinfo/' + zone for zone in win_tz.values()}


def _keep_in_bundle(dest):
    name = dest.replace('\\', '/')
    if name.startswith(_DROP_PREFIXES) or name.endswith(_DROP_SUFFIXES):
        return False
    return not name.startswith('tzdata/zoneinfo/') or name in _KEEP_ZONES


a.binaries = [entry for entry in a.binaries if _keep_in_bundle(entry[0])]
a.datas = [entry for entry in a.datas if _keep_in_bundle(entry[0])]

pyz = PYZ(a.pure)


# ── Splash ───────────────────────────────────────────────────────────────────
# The bootloader's splash is the app's only loading screen: it appears before
# any Python runs and stays up until the main window is painted. PyInstaller's
# script only shows a still image, so splash_anim.tcl is appended to it. That
# file holds the explanation; what's here just fills in what it needs.

# What one more file costs to unpack, expressed as bytes of file content, so that
# the bar tracks time rather than file count: the bundle's bytes sit in very few of
# its files. Regressed from logged launches, which agreed at 176-214 KB.
SPLASH_BYTES_PER_FILE = 180 * 1024

# Share of the bar for unpacking, and where it may drift to while Python starts,
# from a typical logged launch of this bundle: unpacking took 2.0 s of 3.9 s and
# Python's startup another 0.7 s. File.py's _splash_stage calls carry on from these
# numbers. If the bundle changes much, these shares change with it.
SPLASH_EXTRACT_SHARE = 52
SPLASH_BOOT_CEILING = 70


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
