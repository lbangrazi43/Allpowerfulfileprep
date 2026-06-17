# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

"All Powerful File Prep" — a Windows desktop file-preparation toolkit. Single-file
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
py -3.12 -m PyInstaller --noconfirm AllPowerfulFilePrep.spec
# -> dist/AllPowerfulFilePrep.exe
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
3. Create the GitHub release and upload `dist/AllPowerfulFilePrep.exe` as an asset
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
windows** — a hard requirement throughout. `OFFICE_FILE_TIMEOUT` (300s) bounds a
single hung file. `_launch_office_isolated` also disables macros and pre-answers
the modal prompts (update-links, convert, overwrite) that otherwise hang a hidden
instance. PowerPoint can attach to the user's existing instance (no new PID) and
is then deliberately never killed; Outlook is the user's mail client and is never
killed either.

**Password Removal uses pure-Python `msoffcrypto`, NOT COM — on purpose.** Excel's
COM `Workbooks.Open` pops a *blocking interactive* password dialog on a wrong
password (Word/PowerPoint raise cleanly, Excel does not), which would hang a
hidden instance. `msoffcrypto` decrypts the standard encrypted container with no
Office process, no prompt, and an instant clean error on a bad password. The other
tools call `_is_password_protected()` to skip locked files and point the user at
Password Removal.

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

**Packaging specifics (`AllPowerfulFilePrep.spec`).** One-file build that
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

**AI naming backend — currently a SHELL.** The service details live in the
module-level `_AI_BACKEND` dict and will be baked into the build by the
developer; **end users never see or configure anything**. While any `PLACEHOLDER`
value remains, `_ai_backend_ready()` is False, both AI checkboxes (rename and
date-prefix) render disabled with "(coming soon)", and `_start_ds` refuses any AI
path (organize-by-filetype still works). Renames apply AI suggestions automatically —
there is deliberately no per-file approval dialog. The API is called with stdlib
`urllib` (no SDK dependency; request shape is Azure-OpenAI-style chat
completions, adjustable when real details arrive) and every step degrades
gracefully when text extraction yields nothing or the API is unreachable.
