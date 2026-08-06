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

# Cut a release (build + tag + push + publish + archive). See Releases below.
python release.py 1.7.2 --notes notes.md

# Ask a built exe what version it is (the build is windowed, so it answers
# through a file; --version writes to an attached parent console instead)
dist\BoxOfScraps.exe --version-file v.txt
```

Runtime dependencies (import-time or lazily imported): `tkinterdnd2`, `pywin32`
(`win32com`/`pythoncom`), `Pillow`, `markitdown`, `msoffcrypto-tool`, `py7zr`
(.7z), `rarfile` (.rar), `pillow-heif` (.heic/.heif). Several are imported lazily
with a clear "install X" error so the rest of the app still runs when one is
missing.

**There is no automated test suite.** Verification is done by running the app, or
by writing throwaway COM scripts (load `File.py` via `importlib`, call the target
function against real Office files, then delete the script). COM apartments are
thread-bound, so such scripts must keep a single `pythoncom.CoInitialize()` for
their whole run and avoid `CoUninitialize()` between Office launches.

## Releases

**Use `release.py` — it performs every step below in order, so none can be
skipped:**

```powershell
python release.py 1.7.2 --notes notes.md              # cut a stable release
python release.py 1.8.0-beta.1 --notes notes.md --prerelease
python release.py 1.7.2 --notes notes.md --dry-run    # preflight only, no changes
```

**Add the `CHANGELOG.md` section before running it** — preflight fails without a
`## <version>` entry, since that file is what the in-app update screen shows.
Keep it to one terse line per change; `--notes` still carries the long-form
write-up that becomes the GitHub release body.

It refuses to run unless it is on `main` with a clean tree, the version is
well-formed and newer than the current `APP_VERSION`, the tag is unused both
locally and on the remote, and the token can actually push to the repo — checked
in preflight so a permissions problem surfaces before the ~10-minute build rather
than at the upload, with a commit already pushed.

**Release credentials.** Set `APFP_RELEASE_TOKEN` to a fine-grained GitHub token
limited to this repository with *Contents: read and write*; `GITHUB_TOKEN` is
honoured next, and `git credential fill` is the last resort. That fallback is
normally an OAuth token carrying the classic `repo` scope — read and write over
*every* repository on the account — so preflight prints which source it used and
warns when the broad one is in play, since a silent fallback defeats the point.
Tokens bypass 2FA regardless (that is what makes them scriptable), so narrowing
the token is what limits the blast radius of a leak, not the account's 2FA. After building it runs the new exe with
`--version-file` and aborts unless the binary reports the version it was just
built from — catching a build that silently didn't pick up the bump. Nothing is
pushed or published until you confirm, and declining reverts the `APP_VERSION`
edit. It then uploads both assets and archives the previous stable release,
finishing by printing what is publicly visible and what `/releases/latest`
returns.

The steps it automates, for when it has to be done by hand:
1. **Bump `APP_VERSION` in `File.py` first.** It is the single source of truth
   for the in-app updater, which compares it against the newest release tag.
   Ship a stale value and the new exe still believes it is the old version, so
   it offers the same update forever. It must match the tag you are about to
   push. **`APP_RELEASE_DATE` sits directly beneath it and is bumped in the same
   breath** — the About page shows it under the version number, and `release.py`
   takes the value from that version's `CHANGELOG.md` heading so the date on the
   About page and the date in the changelog can never disagree. It is the same
   class of trap as a stale `APP_VERSION`: nothing errors, it just tells every
   user the wrong thing forever.
2. Build the exe with the command above (confirm `PYI_EXIT=0`).
3. Commit on `main` with a `Co-Authored-By:` trailer; tag `vX.Y.Z`; push both.
4. Create the GitHub release and upload `dist/BoxOfScraps.exe` as an asset
   via the REST API, using a token obtained from `git credential fill`
   (host `github.com`). `gh` is not installed in this environment.
   The asset **must** be named exactly `BoxOfScraps.exe` — that is the name the
   updater looks for — and must be the bare `.exe`, never a zip (files extracted
   from a zip inherit its Mark-of-the-Web tag, which reintroduces the manual
   Unblock step the updater exists to avoid).
   Optionally also upload `BoxOfScraps.exe.sha256` (a `sha256sum`-style line);
   when present the updater verifies the download against it, and when absent it
   falls back to its PE-header validation.
5. **Auto-archive the previous stable release.** Immediately after publishing a new
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

**The PDF page is table-driven — `PDF_MODES`.** One `_PdfMode` row per dropdown
entry carries the mode key `convert_file` dispatches on, the label, the noun used
in dialogs, the accepted extensions, the Office `families` the mode needs, whether
it also needs Outlook, the drop-zone hint, and an `auto` flag. `_allowed_extensions`,
`_add_files`, `_update_drop_label`, `_start_convert` and `_convert_worker` all read
from it, and `AUTO_PDF_EXTS` is *derived* from it — so a new file type is one row,
and the page and Folder Unzipping's auto-PDF cannot drift apart. `families` is a
tuple because a mode may drive two apps (Access needs Excel to lay the tables out);
the worker launches, tracks, watchdog-kills and replaces every family in it, and
`convert_file` takes a `{family: app}` dict rather than one parameter per
application. `auto=False` keeps a type off auto-PDF — set for Access and OneNote,
because auto-PDF **deletes the original** once the PDF is written and a table dump
loses queries, forms and relationships in a way a Word→PDF never does.

The table is built in two pieces. `_PDF_TYPE_MODES` holds the single-type modes;
`PDF_MODES` prepends the **`"all"` mixed-batch mode**, which is therefore first in
the dropdown and — since `self._mode` defaults to `PDF_MODES[0].label` — the page's
default. It owns no converter: `PDF_MODE_BY_EXT` (the *complete* extension map,
including the Access and OneNote types `AUTO_PDF_EXTS` leaves out) routes each file
to its own mode, and extensionless files go to `text_report`, the same rule auto-PDF
uses. Its `families` is empty and `outlook` False deliberately — which applications
a batch needs isn't knowable until the files are seen, so `_convert_worker`'s `_app`
helper starts them on first use rather than launching Word, Excel, PowerPoint,
Visio, Access and OneNote up front. A single-type batch still launches up front and
aborts with one dialog if its app is missing; in a mixed batch a missing application
fails only the files that needed it. Outlook is acquired lazily and **only for
`.msg`** — `.eml` is parsed in Python and rendered through Word, so a machine
without Outlook can still convert it.

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
`EnableEvents`, Access `DoCmd.SetWarnings`/`Echo` — Access has no `DisplayAlerts`).
PowerPoint can attach to the user's existing instance (no new PID) and is then
deliberately never killed; Outlook is the user's mail client and is never killed
either. OneNote is in `NEVER_QUIT_FAMILIES` for the same reason: single-instance,
holding the user's synced notebooks, so the object we get is *their* app.

**A successful `DispatchEx` is not proof the app can be driven.** Office leaves
ProgIDs and CLSIDs registered for components that can't actually be activated, so
the object comes back hollow and only fails on first use — which would be one
cryptic error *per file* for a whole batch. Two mechanisms handle this:
`OFFICE_INFO` maps each family to a **tuple of ProgIDs tried in order**, and
`OFFICE_REQUIRED_MEMBERS` + `_office_app_usable` reject a candidate whose methods
don't resolve, so `_launch_office_isolated` returns the first ProgID that actually
answers. OneNote needs all of this: on 64-bit Python against 32-bit Office the
type library behind `OneNote.Application` is registered only in the 32-bit
registry view, so every method lookup fails with "Library not registered" while
the older `OneNote.Application.12` ProgID resolves against a library that loads.
Separately, `COM_SERVER_UNREACHABLE_HRESULTS` (`0x80080005` server execution
failed, `0x800706BE` RPC call failed, `0x800706BA` RPC server unavailable) marks
"installed but Windows can't start it" — reported as a repair suggestion rather
than "please install X", which would send the user to fix something that isn't
broken.

**Every worker that drives Office must go through `_launch_office_isolated` —
never a bare `Dispatch`.** `Dispatch` *attaches to the user's already-open Office
window*; the worker then hides it, drives it, and quits it at the end with alerts
suppressed, silently discarding their unsaved work. (Auto-PDF did exactly this
until it was fixed to use isolated instances.) A helper that launches Office
without returning PIDs to track is a bug, not a shortcut.

**Email attachments, and auto-PDF's work queue.** The email converters have
always extracted attachments beside the PDF — `_save_msg_attachments` (Outlook
COM, skipping `olEmbeddeditem`/`olOLE`) and `_save_eml_attachments` (pure Python,
skipping body parts and inline images by Content-ID) write into
`_attachments_dir(out_path)`, i.e. a folder named after the PDF. AI Preparation
has its own parallel pair (`_extract_*_attachments` → `_email_attachments_to_markdown`,
into `<stem>_attachments`) which uses pure-Python `extract_msg` instead of COM and
drops image attachments, since images are noise in a text extraction.

What was missing was the loop: `_auto_pdf_scan_and_convert` built its work list
*once*, so attachments written during the run were never converted — the
spreadsheet inside the email, which is usually the thing the user actually wanted,
was the one file left raw. It is now a **queue**: converters report what they
wrote through `convert_file(..., attachments_out=[])`, those paths are appended,
and an attachment that is itself an archive is unpacked by
`_expand_attached_archive` with its contents queued too, bounded by
`AUTO_PDF_MAX_CONTAINER_DEPTH` and a `seen` set so nothing is processed twice.

**Deletion was deliberately not extended to any of it.** `originals` is still
built only from `candidate_files`, so `_delete_original_file`'s hard safeguard is
untouched and only upload files are ever deleted. Attachments and unpacked
contents are always kept — the email they came from *is* deleted once its PDF
exists, so keeping them is what stops a bulk run leaving a picture of a
spreadsheet as the only copy. Known gap: an email attached to an email is still
not extracted by either backend (`olEmbeddeditem` is skipped, and a
`message/rfc822` part yields no payload bytes).

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
(`ENCRYPTABLE_EXTS` = `PWD_EXTS` + `ARCHIVE_EXTS`), so batches of images/text/email
don't pay to open and parse files that can never be encrypted.

**Archive extraction (`.zip`/`.7z`/`.rar`).** `ARCHIVE_EXTS` is what Folder
Unzipping accepts and what `_recursive_unzip_in_place` looks for at any depth, in
any mix. `_extract_archive` dispatches: stdlib `zipfile`, `py7zr` (pure Python, so
it bundles cleanly), and `rarfile` for RAR. **RAR needs an external extractor** —
its decompression is proprietary and `rarfile` is only a parser — so `_extract_rar`
tries `rarfile` (which finds an installed unrar/unar/bsdtar itself) and then falls
back to invoking Windows' own `tar.exe` directly. That fallback is the point: bsdtar
on libarchive has shipped in Windows since 10 1803 and reads RAR4 and RAR5, so
`.rar` works on a stock machine **without bundling a third-party binary** into the
exe. The subprocess runs with `_CREATE_NO_WINDOW` (the app is windowed; a child
console would be a visible pop-up) and a timeout. `_ArchiveProtected` is raised for
encrypted archives so the UI lists them under "password-protected" rather than
"corrupt", and `_archive_needs_password` extends `_is_password_protected` to
`.7z`/`.rar` — deliberately reporting *not* protected when no extractor is
available, so the user gets the real "install X" message instead of a wrong one.
Document Structuring accepts every `ARCHIVE_EXTS` format at the top level and
expands whatever is nested inside them, since it shares `_extract_archive` and
`_recursive_unzip_in_place`.

**Image conversion.** `IMAGE_EXTS` splits into `_WORD_NATIVE_IMAGE_EXTS`, which
Word's `AddPicture` filter reads directly, and `_PILLOW_IMAGE_EXTS`
(`.webp`/`.heic`/`.heif`/`.ico`), which it cannot. `_image_pages_for_word` bridges
both gaps by rewriting what Word can't read into temp PNGs in the caller's own
`mkdtemp` — and by splitting a **multi-page TIFF** into one PNG per page, where
`AddPicture` would silently take only the first. Multi-frame splitting is limited
to TIFF on purpose: an animated GIF is one picture, not a 200-page document. HEIC
needs `pillow-heif`, registered lazily by `_register_heif`. Two Word-side details
are load-bearing: `_insert_picture_fitted` **downscales to the printable area**
(an iPhone photo carries no DPI and lands ~42 inches wide, straight off the page),
and pages are separated by setting `ParagraphFormat.PageBreakBefore` rather than
inserting a page-break character — `InsertBreak` puts the break in a paragraph of
its own whose paragraph mark claims a line and pushes a **blank page** out of the
end, which turned a 3-page TIFF into a 4-page PDF. `_flatten_paragraph_spacing`
zeroes Normal's 8pt space-after for the same class of reason.

**Access → PDF (`.accdb`/`.mdb`).** `_access_to_pdf` exports every user table onto
its own sheet of one temp `.xlsx` — `DoCmd.TransferSpreadsheet` appends a sheet per
call when the file name is reused — then hands that workbook to the existing
`_excel_to_pdf`, which already fits each sheet to one page wide and turns wide ones
landscape. Access's own `DoCmd.OutputTo` was rejected because it prints the
datasheet at 100% (slicing wide tables across pages) and writes one file per table,
which doesn't fit the one-input-one-PDF model. This is a dump of the *data*:
queries, forms, reports and relationships are not exported, the database is opened
non-exclusively and never modified, and `_ACCESS_SYSTEM_PREFIXES` filters the Jet
bookkeeping tables. A database-password prompt is interactive and can't be
pre-answered, so that case is left to the per-file watchdog. Note this is the one
mode needing two Office apps, which is why `families` is a tuple.

**OneNote → PDF (`.one`).** A `.one` file is a *section*, not a self-contained
document, so `OpenHierarchy` must add it to the OneNote hierarchy before `Publish`
can render it (`pfPDF = 3`); `_onenote_to_pdf` then closes it again so a converted
file isn't left in the user's notebook list. `_first_com_str` copes with the `[out]`
object-ID parameter not surfacing through late binding, and `_onenote_section_id`
falls back to matching the file path in the hierarchy XML. Only desktop OneNote
exposes this API — the Store "OneNote for Windows 10" app does not. **This path is
implemented to the documented API but has not been verified end to end**: the
development machine's OneNote runs by hand yet cannot activate its COM server at
all (see the ProgID/HRESULT notes above), so it exercises only the error paths.

**Basic Text conversion (the mode shown as "Basic Text"; key still
`text_report`).** `TEXT_REPORT_EXTS` covers everything that wants the same
monospace, fit-to-widest-line treatment: report dumps from legacy/ERP/mainframe
systems (`.rep`/`.rpt`/`.prn`/`.lst`), logs and program output
(`.log`/`.out`/`.err`/`.dat`/`.trace`), structured data
(`.json`/`.jsonl`/`.ndjson`/`.yaml`/`.sql`/`.tsv`/`.psv`/`.tab`), config
(`.ini`/`.cfg`/`.conf`/`.toml`/`.properties`/`.env`), prose formats Word can't
import (`.md`/`.text`/`.nfo`/`.asc`/`.rst`), diffs, and the scripts that turn up as
evidence in controls work (`.bat`/`.cmd`/`.ps1`/`.sh`/`.py`/`.js`/`.vbs`/`.css`).
The **key stays `text_report`** — it is what `convert_file` dispatches on and what
`_text_report_to_pdf` is named for; only the user-facing wording changed.

`TEXT_REPORT_ALSO_ACCEPTS` (`.txt`, `.xml`, `.csv`) is the mode's `shared` list:
extensions it takes when picked by hand but does **not** own. Word owns `.txt`/`.xml`
and Excel owns `.csv`, because their importers wrap prose and parse columns properly
— `_fit_report_to_page` would shrink one long unwrapped paragraph to 5pt. Picking
Basic Text is how the user asks for the raw text instead. `_mode_accepts(mode)`
returns owned + shared and is what the drop zone and file dialog filter on;
`PDF_MODE_BY_EXT` and `AUTO_PDF_EXTS` are built from `exts` alone, so automatic
routing stays unambiguous and no two modes may own the same extension.

Word selects an importer by extension and has
none registered for these, so `_text_report_to_pdf` writes a temp `.txt` copy
(UTF-8 with BOM, in its own `mkdtemp`) and opens *that* — never the original, and
never a temp file next to the user's source. `_read_report_text` decodes through
utf-8/cp1252/latin-1 and calls `_looks_binary` first, so a binary file wearing a
report extension (a compiled Crystal Reports `.rpt`) fails with a clear message
instead of emitting mojibake. Wired into the PDF page and into `AUTO_PDF_EXTS`,
and it also handles auto-PDF's **extensionless** files.

**How the page layout is chosen — three rules that between them keep output
readable.** `_fit_report_to_page` forces Consolas and zeroes the paragraph
spacing Word's Normal style would otherwise double-space everything with, then:

1. **Tabs are expanded to spaces first** (`_REPORT_TAB_WIDTH`, 8). Word's default
   tab stops are half-inch, which matches neither the monospace grid nor the
   one-character-per-tab a width estimate assumes — so a `.tsv` would drift out
   of alignment *and* be measured too narrow. Expanding in `_text_report_to_pdf`
   means the temp file Word opens and the width computed from it agree exactly.
2. **Width comes from `_report_fit_columns`, a 95th percentile of line lengths —
   not the maximum.** One 300-character stack trace in a log of 80-character
   lines used to shrink the entire document to 5pt to protect alignment on the
   one line that had none. Genuinely column-aligned content is unaffected,
   because there the percentile *is* the maximum.
3. **`_report_layout` will not shrink below `_REPORT_FONT_MIN` (7pt) to avoid
   wrapping.** If the content fits portrait it stays portrait; past
   `_REPORT_LANDSCAPE_COLS` it goes landscape; and if even the floor can't fit
   it, that means wrapping is unavoidable — at which point the alignment this
   whole layout exists to protect is already gone, so it returns to portrait at
   `_REPORT_FONT_WRAP` (9pt) and lets it wrap. Short lines read better than a
   wide wall of tiny text. Wherever wrapping does occur, continuations get a
   hanging indent so they read as part of the line above.

**`_inches()` replaced `Application.InchesToPoints`.** On a hidden, automated
Word instance that COM helper raises "Unspecified error" — and because every call
site wrapped it in a try/except, the failure was *silent*: the margins it was
there to compute were never applied, so pages kept the Normal template's defaults
and the fitting maths was computed against a width the document didn't have. The
conversion is exactly 72 points to the inch, so it belongs in Python. This
affected `_fit_word_doc_to_page` (Word/HTML/email output) as well as the text
path.

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
- `_is_within(child, parent)` / `_files_under_folder(folder)` — the containment
  safeguards for tools that take a **folder** as input. `_is_within` resolves both
  sides and compares path *components* (`parent in child.parents`), so a string
  prefix can't make `…\Reports2` read as inside `…\Reports` and a junction can't
  disguise a path. `_files_under_folder` walks with `os.walk(followlinks=False)`,
  prunes reparse-point directories, and re-checks every hit with `_is_within` — a
  symlink planted in an input folder therefore cannot pull in files from elsewhere
  on disk. Use both whenever a new tool accepts a folder.
- `_queue_add_path` / `_input_entry_label` / `_input_entry_name` — the shared
  mixed input queue (archive / folder / loose file) behind Document Structuring
  and AI Organization, so the two can't drift apart in what they accept. Paths
  that don't exist are never queued, which is what lets the queue stand as the
  literal definition of what the tool may read.
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
    `ttk.Combobox`). Popup animates open **downward, always**, and gets slightly
    rounded corners via `_apply_round_corners` (Windows DWM `DWMWCP_ROUNDSMALL`,
    best-effort/no-op elsewhere). Calls `command` only when the selection actually
    changes (mirrors `<<ComboboxSelected>>`). Animation `after` ticks are scheduled
    on the widget (not the transient popup) and cancelled on close.
    It shows at most `MAX_VISIBLE_ROWS` (10) — fewer if there isn't room below on
    screen — and **scrolls** for the rest, so a growing list can never push options
    off the bottom. Scrolling is deliberately not a Canvas: rows are `place`d inside
    a clipping frame sized to the visible height, Tk clips children to their parent,
    and scrolling just re-places them. Mouse wheel is bound on the Toplevel (which
    is in the bindtag chain of every widget inside it, so the wheel works over the
    rows), and the slim thumb is draggable. The popup opens with the current
    selection already in view. `_close_popup` drops the row/thumb references so a
    late wheel or drag event can't touch destroyed widgets.
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

**Document Structuring tool — organize by filetype, and nothing else.** It has
**no options and no toggle**: the one thing it does is sort files into
`<out>/<Type>/` (`FILETYPE_FOLDERS` / `_filetype_folder`). The AI rename and
date-prefix that used to share this page moved to AI Organization below.

Input is a **mixed queue** of three kinds, added by drop, "Add Files" or "Add
Folder" (`_queue_add_path`): an **archive** (any `ARCHIVE_EXTS`, extracted to a
temp dir with nested archives expanded via `_recursive_unzip_in_place`), a
**folder** (walked with `_files_under_folder`), or a **loose file** (taken as it
is). An output folder is **required** — the Run button stays disabled and the
label stays red until one is chosen — but it **need not be empty**.

The safeguards are the point of this tool, and they fire in this order:
1. `_queue_add_path` refuses to queue a path that doesn't exist, so the queue is
   an exact statement of what may be read.
2. Archives extract into `tempfile.mkdtemp()`, **never** beside the source — this
   is what makes `_recursive_unzip_in_place` safe here, since that function
   *deletes each archive it unpacks* and must therefore only ever see our copy.
   A `.zip` sitting inside a dropped *folder* is deliberately **not** expanded
   (it is filed under Archives); expanding it would mutate the user's own tree.
3. `_files_under_folder` won't follow a symlink/junction out of an input folder.
4. `_start_ds` rejects an output folder that is inside a queued input folder —
   that would write the run's output into the tree it is reading.
5. `_ds_copy_into_type_folder` re-checks both the type subfolder and the final
   path with `_is_within(out_base)` immediately before copying.
6. Files are **copied**, `_unique_output_path` suffixes collisions, and only the
   type subfolders are created — so originals, and anything already sitting in
   the output folder, are both untouched. **There is no `unlink` in this tool at
   all.**

Password-protected archives are skipped and listed; corrupt ones are reported.
(Note: Folder Unzipping still *does* require an empty output folder — that
constraint was removed only from Document Structuring.)

**AI Organization tool — present, but switched off.** `AI_ORGANIZATION_ENABLED`
(`False`) is the single master switch, and `_ai_organization_available()` is
`AI_ORGANIZATION_ENABLED and _ai_backend_ready()` — credentials appearing in the
environment is deliberately *not* on its own permission to run the feature, so
both must be true. While it is off: a "Coming soon" banner shows, both
`_ToggleSwitch`es render disabled with "(coming soon)" and are forced off, the
Run button is greyed out (not live-but-erroring), and `_start_aiorg` refuses with
a message that names which of the two gates failed.

The page is built right up to the point of doing work — it accepts the same mixed
archive/folder/file queue as Document Structuring — but it has **no worker**.
That is the missing piece: `_extract_document_text`, `_ai_suggest_filename` and
`_ai_infer_date` are complete and are what a worker will call, and are kept
rather than stubbed so enabling the feature is a flag flip plus a worker, not a
rewrite.

**Self-update from GitHub Releases (the "About & Updates" page).** `APP_VERSION`
(top of `File.py`) is compared against `/releases/latest`, which returns only the
newest *published, non-prerelease* release — so it dovetails with the auto-archive
invariant above and betas are never offered to users. The page has **one** accent
button: it reads "Check for Updates", and finding a newer release turns that same
button into "Download & Install X" in place (`_upd_set_mode`), so there is never a
dead Install button on screen. After each check the button locks for
`UPDATE_CHECK_COOLDOWN` (30s) and counts down in its own label — GitHub's quota is
60 calls an hour *per IP*, shared by everyone behind one office connection, so a
spammed button could exhaust it for the whole building. If the quota is hit anyway,
the error names the local time it resets, read from the `X-RateLimit-Reset` header
(a 403 only counts as a quota error when `X-RateLimit-Remaining` is 0). Three
design points carry the whole feature:

- **No "Unblock" step.** The Properties → Unblock checkbox exists because the
  *downloading application* writes the Mark-of-the-Web alternate data stream
  (`:Zone.Identifier`, `ZoneId=3`). Browsers do; `urllib` does not. Because
  `_download_release_asset` fetches the exe itself the file arrives untagged and
  runs immediately, and SmartScreen's shell prompt — which fires *on* that tag —
  never appears. `_strip_mark_of_the_web` deletes the stream anyway at every step
  so the guarantee never rests on that behaviour. (Verified empirically: a
  urllib-downloaded asset has no `:Zone.Identifier`.)
- **No helper `.bat`.** Windows locks a running image against deletion and
  overwriting but *permits renaming it*, so `_swap_in_new_exe` does two renames
  in-process — `target → target.old`, then `staged → target` — instead of a batch
  script that has to outlive the process and poll for it to exit. If the second
  rename fails the first is undone, so a failed update always leaves a working
  app. The `.old` backup can't be deleted by the update that created it (it's
  still the running image), so `_cleanup_previous_update` clears it at the next
  startup — on a background thread with ~11s of retries, because that next
  startup is *launched by the outgoing app* and routinely gets there while the
  `.old` image is still running and undeletable. One try was not enough; the
  symptom was a `.old` that never went away.
  **The new build MUST land at the same path under the same filename** — taskbar
  pins, desktop and Start Menu shortcuts are `.lnk` files that resolve by path, so
  keeping the path identical is what lets a pinned copy survive an update and
  launch the new version. Verified end-to-end, including that shell *link
  tracking* does not hijack the pin onto `BoxOfScraps.exe.old`: tracking only
  engages when the stored path fails to resolve, and here it always resolves.
  Anything that installs to a versioned folder or renamed exe would silently break
  every pin and shortcut the user has.
- **The restart must not inherit PyInstaller's one-file environment.** The
  bootloader unpacks the bundle to `%TEMP%\_MEInnnnnn` and tells the Python
  process it starts where that is via `_PYI_APPLICATION_HOME_DIR`,
  `_PYI_ARCHIVE_FILE`, `_PYI_PARENT_PROCESS_LEVEL`, `_PYI_SPLASH_IPC`
  (`PYI_ONEFILE_ENV_VARS`; `_MEIPASS2` pre-6.0) — and *that* process passes them
  to everything it spawns. A bare `Popen([exe])` therefore starts an app that
  believes it is already unpacked, skips extraction, and runs out of the outgoing
  app's temp folder, which the outgoing app deletes on its way out. The restart
  then dies mid-startup on whichever bundled file it needed next (seen in the
  wild: `pyi_rth_pkgres` → `pkg_resources` → `setuptools/_vendor/jaraco/text/`
  `Lorem ipsum.txt`, a real bundled file that had been deleted underneath it),
  and the exiting app shows "Failed to remove temporary directory" because the
  new one still holds the DLLs open. It is a race, so it passed on the build
  machine and failed elsewhere. `_clean_child_env()` strips the variables and
  `_relaunch` passes it as `env=`; `_relaunch_as_admin` takes them out of
  `os.environ` around the call instead, since `ShellExecuteW` has no `env`
  parameter. **Any future code that launches the app, or any other one-file
  exe, must go through `_clean_child_env`.** (Reproduced and verified with a
  throwaway one-file probe: without the scrub the child reports the parent's
  `sys._MEIPASS` and watches its bundled files vanish; with it the child gets
  its own `_MEI` directory and a complete bundle.)

**Cumulative update notes come from `CHANGELOG.md`, not the releases API.** The
update page shows every version between the one running and the one on offer, so
a user who skipped 1.7.1 and 1.7.2 sees all three entries rather than only
1.7.3's. It cannot read those from the API: auto-archive drafts each previous
stable release, and GitHub hides drafts from clients without push access — to an
ordinary user, only the newest release exists. So the notes live in the repo and
are fetched from `CHANGELOG_RAW_URL` (raw.githubusercontent.com, allowed by the
existing `.githubusercontent.com` suffix rule, and not the REST API — so it costs
nothing against the shared 60-calls-an-hour quota). Fetched only when an update
actually exists, and only from `_upd_check_worker`'s background thread.

`_changelog_since` bounds the selection at both ends (above the installed
version, no higher than the tag being installed) and drops prereleases, which the
updater never offers. `_fetch_changelog` returns `""` on *any* failure and
`_upd_check_done` falls back to the release's own body — enriching the page must
never break a working update check. **Entries are deliberately terse**, one line
per change: the long-form write-up stays in the GitHub release body, and this
file is the quick read shown in the app. `release.py` fails preflight when the
version being cut has no section, before the ~10-minute build.

Safeguards, in the order they fire: `_current_exe_path()` returns None unless
`sys.frozen` — from source `sys.executable` is python.exe, and every disk-touching
path refuses rather than swap *that* out. `_other_operation_running()` reads the
existing per-page Cancel-button states (enabled for exactly the duration of an
operation) to refuse an update that would close the app mid-conversion.
`_dir_is_writable` runs *before* the ~120 MB download so a Program Files install
fails fast and offers `_relaunch_as_admin` instead of dying at the swap;
`_has_room_for_update` checks for the download plus the backup. Nothing touches
the installed app until the download passes `_looks_like_windows_exe` (MZ + PE
signature — catches a captive-portal page served with a 200) plus a size floor and
the optional published SHA-256. **The published SHA-256 is an integrity check, not
an authenticity one** — it is served from the same release as the binary, so it
catches corruption and truncation but not an attacker who can alter both. Nothing
in the updater verifies *who* built the exe; a signing key would be needed for
that. Network traffic is confined by `_is_allowed_update_url` (HTTPS only, on a
GitHub host, via `urlsplit().hostname` so `https://github.com@evil.example` is
correctly read as `evil.example`) and by `_update_opener`, whose redirect handler
re-applies that rule to *every* hop — urllib's stock handler will otherwise follow
an https→http downgrade to any host, so checking only the first URL proves
nothing. Writing stops the moment a response exceeds the size GitHub declared
(`UPDATE_MAX_DOWNLOAD_BYTES` when it declares none), because comparing sizes after
the stream ends is too late: the disk is already full. The download lands in the
*system* temp dir, never
beside the exe, so a synced install doesn't push 120 MB through OneDrive and back;
`_stage_beside` then moves it onto the exe's own volume because `os.replace`
cannot rename across volumes. `_replace_with_retry` absorbs the transient locks a
sync client or AV scanner takes on a file it just watched change (~9s of backoff),
and `_cloud_sync_provider` detects OneDrive/Dropbox/Box/Google Drive — via
OneDrive's own env vars first, then whole-path-component matching so a folder
named "Box of Scraps" isn't mistaken for Box — to warn in the confirm dialog.
Every failure path leaves the current version untouched and offers the releases
page in a browser.

**AI naming backend — credentials come from the environment, never the build.**
No secret is hardcoded or baked into the source/`.exe`; `_load_ai_backend()`
reads the endpoint, key, and deployment from environment variables at call time
(`APFP_AI_ENDPOINT`, `APFP_AI_API_KEY`, `APFP_AI_DEPLOYMENT`, optional
`APFP_AI_API_VERSION` defaulting to `2024-06-01`). This keeps the shipped binary
safe to distribute — a malicious reader of the source/binary finds no key to
exploit. When any required variable is unset, `_ai_backend_ready()` is False,
which is one of the two gates on the AI Organization page (see above) — the other
being `AI_ORGANIZATION_ENABLED`, so credentials alone never switch the feature
on. The API is called with stdlib `urllib` (no SDK dependency; request shape is
Azure-OpenAI-style chat completions) and every step degrades gracefully when text
extraction yields nothing or the API is unreachable. When the feature does go
live, renames are intended to apply AI suggestions automatically — deliberately
no per-file approval dialog.
