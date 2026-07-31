# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

"Box of Scraps" — a Windows desktop file-preparation toolkit. Single-file
Python/Tkinter GUI shipped as a one-file PyInstaller `.exe`. Windows-only: most
features drive Microsoft Office through COM automation (`win32com`) and use Win32
APIs via `ctypes`. The entire application lives in **`File.py`** (~4,000 lines);
there is no package structure.

## Commands

```powershell
# Run in development (needs the runtime deps below; repo uses a local .venv)
python File.py

# Syntax-check after edits (there is no linter configured)
python -m py_compile File.py

# Build the one-file exe. Use the GLOBAL Python 3.12 — it has PyInstaller AND the
# bundled deps; the .venv has the runtime deps but NOT PyInstaller.
py -3.12 -m PyInstaller --noconfirm BoxOfScraps.spec
# -> dist/BoxOfScraps.exe
# If `build/` triggers a OneDrive PermissionError, delete the build/ folder first
# rather than passing --clean.

# Regenerate icon.ico from owl_source.png (rarely needed)
python make_icon.py
```

Runtime dependencies (import-time or lazily imported): `tkinterdnd2`, `pywin32`
(`win32com`/`pythoncom`), `Pillow`, `markitdown`, `msoffcrypto-tool`. Several are
imported lazily with a clear "install X" error so the rest of the app still runs
when one is missing.

**There is no automated test suite.** Verification is done by running the app, or
by writing throwaway COM scripts (load `File.py` via `importlib`, call the target
function against real Office files, then delete the script). COM apartments are
thread-bound, so such scripts must keep a single `pythoncom.CoInitialize()` for
their whole run and avoid `CoUninitialize()` between Office launches.

## Releases

Releases are published manually (not via CI):
1. Build the exe with the command above (confirm `PYI_EXIT=0`).
2. Commit on `main` with a `Co-Authored-By:` trailer; tag `vX.Y.Z`; push both.
3. Create the GitHub release and upload `dist/BoxOfScraps.exe` as an asset
   via the REST API, using a token obtained from `git credential fill`
   (host `github.com`). `gh` is not installed in this environment.
4. **Auto-archive the previous stable release.** Immediately after publishing a new
   stable (non-prerelease) release, set the previously-visible stable release to
   `draft: true` via the REST API, so only the newest stable release stays public.
   Invariant: at most one published stable release at a time (prereleases/betas are
   left visible). Drafting preserves the tag + asset and is reversible. This does
   **not** apply when cutting a prerelease/beta — those don't archive anything.

## Architecture

**Shell + lazy pages.** `ConverterApp` owns all UI. `_build_shell()` renders the
sidebar nav buttons; `_show_page(key)` builds each page lazily on first visit by
calling its `_build_<key>_page(parent)`. Every tool is one page and follows the
same layout: title + description, a drop zone wrapping a `Listbox`,
Add/Remove/Clear buttons, an output-folder row, tool-specific options, a
progress bar, a status label, an `_OpTimer`, and a Run + Cancel button row.
**All per-page state is prefixed** to avoid collisions (`self._pwd_*`,
`self._office_*`, `self._ai_*`, `self._uz_*`, …). To add a tool: add a nav entry
in `_build_shell`, a branch in `_show_page`, a `_build_<key>_page`, prefixed
state in `__init__`, and a worker.

**Threading model.** Each operation runs on a `daemon=True` thread. Workers never
touch widgets directly — they marshal UI updates back with
`self.root.after(0, lambda: ...)`. Cancellation is a `threading.Event` per page.
When a worker needs a modal dialog (e.g. a password or rename prompt), it calls
`_run_on_main(fn)`, which schedules `fn` on the Tk thread via `root.after` and
blocks the worker on an `Event` until the dialog returns. Such dialogs do **not**
`grab_set` (that would freeze the main-window Cancel button); instead the open
dialog is stored on `self` so Cancel can `destroy()` it.

**Office COM automation (the load-bearing, fragile part).** For each Office family
the tools launch a **dedicated, isolated instance** via
`_launch_office_isolated(family)` (uses `DispatchEx` and diffs process IDs before
/after so it learns the PID of *only the instance it started*). This is what lets
Cancel and the per-file watchdog force-kill a stuck conversion with
`_terminate_pid(pid)` **without ever touching the user's own open Office
windows** — a hard requirement throughout. `_launch_office_isolated` also disables
macros, pre-answers the modal prompts (update-links, convert, overwrite) that
otherwise hang a hidden instance, and turns off the work that is pointless in an
invisible app (`ScreenUpdating`, Word background `Pagination`, Excel
`EnableEvents`). PowerPoint can attach to the user's existing instance (no new
PID) and is then deliberately never killed; Outlook is the user's mail client and
is never killed either.

**Every worker that drives Office must go through `_launch_office_isolated` —
never a bare `Dispatch`.** `Dispatch` *attaches to the user's already-open Office
window*; the worker then hides it, drives it, and quits it at the end with alerts
suppressed, silently discarding their unsaved work. (Auto-PDF did exactly this
until it was fixed to use isolated instances.) A helper that launches Office
without returning PIDs to track is a bug, not a shortcut.

**`_FileWatchdog` bounds a single file** so one stuck conversion can't hang a
batch. Used as a context manager around each file by all three conversion workers
(`_convert_worker`, `_office_worker`, `_auto_pdf_scan_and_convert`); after
`OFFICE_FILE_TIMEOUT` (300s) it force-kills the tracked PIDs, which makes the
blocked COM call raise so the loop records a failure and continues. `fired`
distinguishes a timeout from a real conversion error. With no PIDs to kill (a
PowerPoint that attached to the user's instance) it stays inert rather than risk
a process we didn't start. Each worker replaces the killed instance before the
next file.

**Password Removal uses pure-Python `msoffcrypto`, NOT COM — on purpose.** Excel's
COM `Workbooks.Open` pops a *blocking interactive* password dialog on a wrong
password (Word/PowerPoint raise cleanly, Excel does not), which would hang a
hidden instance. `msoffcrypto` decrypts the standard encrypted container with no
Office process, no prompt, and an instant clean error on a bad password. The other
tools call `_is_password_protected()` to skip locked files and point the user at
Password Removal. That check short-circuits on extension first
(`ENCRYPTABLE_EXTS` = `PWD_EXTS` + `.zip`), so batches of images/text/email don't
pay to open and parse files that can never be encrypted.

**Text report conversion (`.rep`/`.rpt`/`.prn`/`.lst`).** Plain-text report dumps
from legacy/ERP/mainframe systems. Word selects an importer by extension and has
none registered for these, so `_text_report_to_pdf` writes a temp `.txt` copy
(UTF-8 with BOM, in its own `mkdtemp`) and opens *that* — never the original, and
never a temp file next to the user's source. `_read_report_text` decodes through
utf-8/cp1252/latin-1 and calls `_looks_binary` first, so a binary file wearing a
report extension (a compiled Crystal Reports `.rpt`) fails with a clear message
instead of emitting mojibake. `_fit_report_to_page` then forces Consolas, zeroes
the paragraph spacing Word's Normal style would otherwise double-space the report
with, and goes landscape past `_REPORT_LANDSCAPE_COLS` before scaling the type to
the widest line. Wired into the PDF page (mode `"Text Report"` → `mode_key
"text_report"`) and into `AUTO_PDF_EXTS` for Folder Unzipping's auto-PDF. The same
function now also handles auto-PDF's **extensionless** files, which previously
inlined their own temp-copy logic.

**Cross-cutting helpers worth reusing:**
- `_unique_output_path` / `_unique_pdf_path` — never overwrite; suffix `_2`, `_3`…
- `_delete_original_file(src, allowed)` / `_resolved_set(paths)` — the HARD
  safeguard behind every "delete originals after operation" feature. Deletion only
  happens if `src`'s resolved path is in `allowed`, a set built from the
  operation's own upload list (`_resolved_set(self._ai_files)` etc.) — **never from
  a directory scan**. This guarantees an unrelated file sitting in a chosen output
  folder can never be deleted. Used by `_aiprep_worker`, `_office_worker`, and
  `_auto_pdf_scan_and_convert`. Other `unlink` calls only remove uniquely-named
  output the app just wrote (failure cleanup) or program-created temp files.
- `_OpTimer` / `_build_timer` — the stacked Total + per-file elapsed-time labels.
- `_set_cancel_state` — the grey/red Cancel-button states.
- `_safe_filename`, `_is_password_protected`, `_get_markitdown` (markitdown powers
  Markdown Conversion and document text extraction; lazy, shared instance).
- **Custom modern controls (Pillow-rendered, anti-aliased) — prefer these over the
  native Tk widgets for a consistent look:**
  - `_ToggleSwitch` — pill on/off switch replacing `tk.Checkbutton`. Bound to a
    `BooleanVar`; supports `command`, `"disabled"` state, `.config(state=…/text=…)`,
    and animates the knob slide.
  - `_ModernRadio` — Windows-11-style radio replacing `tk.Radiobutton`; instances
    sharing one variable form a group (selecting one re-renders all via a trace).
  - `_ModernDropdown` — readonly combobox replacement (replaced the PDF mode
    `ttk.Combobox`). Popup animates open downward and gets slightly rounded corners
    via `_apply_round_corners` (Windows DWM `DWMWCP_ROUNDSMALL`, best-effort/no-op
    elsewhere). Calls `command` only when the selection actually changes (mirrors
    `<<ComboboxSelected>>`). Animation `after` ticks are scheduled on the widget
    (not the transient popup) and cancelled on close.
  - Shared color helpers: `_hex_to_rgb`, `_lerp_rgb`; palette `TOGGLE_ON/OFF/DISABLED`.

**Packaging specifics (`BoxOfScraps.spec`).** One-file build that
`collect_all`s pywin32/PIL/markitdown/msoffcrypto and excludes heavy ML/GUI libs
(torch, scipy, PyQt, …) to shrink the archive and speed startup. It bundles a
bootloader `Splash` shown during one-file unpack; the app closes it with
`pyi_splash.close()`. **`import pyi_splash` MUST stay a static import** — a dynamic
import makes PyInstaller omit the module, so the bootloader splash never closes
and resurfaces over the app. A separate animated `_SplashScreen` runs during UI
construction. Asset paths resolve through `_resource_base()` /
`sys._MEIPASS` when frozen.

**Document Structuring tool.** Input is **ZIP archives only** (drop zone and Add
dialog reject non-`.zip`). The worker `_ds_worker` extracts each archive to a
temp dir, expands nested zips at any depth via the shared
`_recursive_unzip_in_place`, then runs the selected structuring on every
extracted file (`_ds_structure_file`): optional AI rename, optional AI
date-prefix (`_ai_infer_date` → `YYYY-MM-DD`, applied as `Date_Name`; with rename
on it becomes `Date_AISuggestedName`), and/or organize into filetype subfolders
(`FILETYPE_FOLDERS` / `_filetype_folder`). Output folder is **optional** (defaults
to each zip's own folder, "Same as source") and **need not be empty** — only the
extracted files are written/organized there: existing files are never moved, swept
into the created subfolders, or overwritten (`_unique_output_path` suffixes
collisions; files are copied, not moved). Password-protected archives are skipped;
bad/corrupt zips are reported. (Note: Folder Unzipping still *does* require an
empty output folder — that constraint was removed only from Document Structuring.)

**AI naming backend — credentials come from the environment, never the build.**
No secret is hardcoded or baked into the source/`.exe`; `_load_ai_backend()`
reads the endpoint, key, and deployment from environment variables at call time
(`APFP_AI_ENDPOINT`, `APFP_AI_API_KEY`, `APFP_AI_DEPLOYMENT`, optional
`APFP_AI_API_VERSION` defaulting to `2024-06-01`). This keeps the shipped binary
safe to distribute — a malicious reader of the source/binary finds no key to
exploit. When any required variable is unset, `_ai_backend_ready()` is False,
both AI checkboxes (rename and date-prefix) render disabled with "(coming soon)",
and `_start_ds` refuses any AI path (organize-by-filetype still works). Renames
apply AI suggestions automatically — there is deliberately no per-file approval
dialog. The API is called with stdlib `urllib` (no SDK dependency; request shape
is Azure-OpenAI-style chat completions) and every step degrades gracefully when
text extraction yields nothing or the API is unreachable.
