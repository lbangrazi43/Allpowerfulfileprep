# Changelog

Short, plain-language notes shown **inside the app**, on the About & Updates
page. When an update is found, Box of Scraps displays every section between the
version you are running and the one you are installing — so skipping releases
never means missing what changed in between.

**Keep entries brief: one line per change.** The full write-up for a release
belongs in its GitHub release notes; this file is the quick read a user gets on
the update screen. `release.py` refuses to cut a release with no section here,
so a version can never ship without its line in the app.

Format is `## <version> — <YYYY-MM-DD>`, newest first. Prerelease sections are
kept for the record but never shown in the app, which is only ever offered
stable releases.

## 1.7.4 — 2026-08-05

- New "All Types" mode on the PDF page, now the default: drop in a mixed pile of
  files and each one is converted according to what it is, all into your chosen
  folder. The single-type modes are still there if you want to filter.
- `.eml` files no longer need Outlook installed to convert.
- Folder Unzipping now accepts `.7z` and `.rar` archives as well as `.zip`, and
  unpacks them nested inside each other in any combination.
- PDF conversion now reads OpenDocument files (`.odt`, `.ods`, `.odp`) — the
  format Google Docs and LibreOffice export.
- PDF conversion now covers the macro-enabled and template variants it was
  missing: `.docm`, `.dotx`, `.pptm`, `.potx`, `.xltx` and the rest.
- New "Basic Text" mode (was "Text Report") covers a much wider set of plain-text
  files — logs and traces, JSON/JSONL/YAML/SQL, TSV/PSV, TOML/INI/config, diffs,
  Markdown, and scripts like `.ps1`, `.bat`, `.py` and `.sh` — all laid out in
  monospace so columns stay aligned. It also accepts `.txt`, `.xml` and `.csv`
  when you pick it directly, for when you want the raw text.
- Text files are much more readable as PDFs. A single very long line — one stack
  trace in a log, one minified record — no longer shrinks the whole document to
  unreadable type; that line wraps instead and everything else stays full size.
- Tab-separated files now line up properly instead of drifting out of alignment.
- Wrapped lines are indented so you can see they continue the line above.
- Fixed: page margins were silently never applied on any Word-based conversion,
  so pages used Word's defaults instead of the intended narrow margins.
- Images now include `.heic`/`.heif` (the format iPhones save photos in),
  `.webp` and `.ico`.
- A multi-page TIFF now becomes a multi-page PDF instead of just its first page.
- Large photos are scaled to fit the page instead of running off the edge.
- New: Access databases (`.accdb`, `.mdb`) convert to a PDF with each table on
  its own sheet, sized to fit the page.
- New: OneNote sections (`.one`) convert to PDF.
- Email attachments are now converted to PDF too during a bulk unzip, instead of
  being extracted and left as-is. A zipped attachment is unpacked and its
  contents converted as well.
- Attachments are always kept alongside their PDF, so nothing is lost when the
  email they came from is converted away.
- About & Updates now shows the release date of the version you're running,
  underneath the version number.
- Clearer message when an Office app is installed but Windows can't start it,
  instead of being told to install something you already have.

## 1.7.3 — 2026-08-04

- Fixed: reopening after an update could fail with a "Failed to execute script"
  error. The update itself always succeeded — the restart is now clean.
- Fixed: the `BoxOfScraps.exe.old` backup of your previous version sometimes
  stayed on disk forever instead of being cleared on the next launch.
- New: update notes now cover every version you skipped, not just the newest —
  so going from 1.7.0 to 1.7.3 shows you 1.7.1 and 1.7.2 as well.

## 1.7.2 — 2026-08-04

- Downloads now stop at the size GitHub declared, instead of checking only after
  the whole file had already been written to disk.
- The download address, and any redirect it follows, must be an encrypted
  GitHub one.
- The app can report its own version from a command line (`--version`).
- Fixed a packaging bug that stopped the app being rebuilt from a clean copy of
  the source.

## 1.7.1 — 2026-08-03

- Version bump used to test the new updater end to end. No functional changes
  from 1.7.0.

## 1.7.0 — 2026-08-03

- New: the app updates itself. About & Updates shows your version, finds newer
  ones, and installs them — no Properties → Unblock step.
- A failed update always leaves your working version in place.
- Refuses to update while another tool is still running.
- Checks folder permissions and disk space before downloading, and offers to
  restart as administrator when installed somewhere restricted.
- Rejects a download that isn't a real Windows program.
- Works from OneDrive, Dropbox, Box, and Google Drive folders.

## 1.6.1 — 2026-07-31

- The app is now **Box of Scraps**; the download is `BoxOfScraps.exe`.
- New: Text Report → PDF for `.rep`, `.rpt`, `.prn`, and `.lst` dumps, with
  fixed-width layout, form-feed page breaks, and automatic landscape.
- Fixed: Folder Unzipping's auto-PDF could take over your open Office windows
  and discard unsaved work. It now uses its own isolated instances.
- Fixed: one stuck file could hang an entire PDF batch.
- Faster hidden-instance conversions, and password checks are skipped for file
  types that can never be encrypted.

## 1.6.0-beta.1 — 2026-06-17

- Beta side release: the Document Structuring tool (ZIP-only), organizing
  extracted files into per-filetype subfolders.

## 1.5.1 — 2026-06-09

- Excel → PDF now fits one page wide, so wide tables are no longer sliced across
  pages column by column.
- Password-protected files are skipped gracefully and listed by name when the
  run finishes, instead of failing it.

## 1.5.0 — 2026-06-09

- New tool: Password Removal, which strips the password from protected Word,
  Excel, and PowerPoint files without launching Office.
- Each tool page now has a short description under its title.

## 1.4.1 — 2026-06-09

- Fixed the loading splash reappearing and staying on top of the app.

## 1.4.0 — 2026-06-09

- New: Office Modernizer, converting old Office files to their current formats
  one-to-one by family.
- Cancel now force-stops a stuck conversion immediately, without touching your
  own open Office windows.
- Suppressed the Office dialogs that used to freeze conversions silently, and
  added a per-file watchdog so one bad file can't stop a batch.
- Live elapsed-time timers on every page.
- "AI Preparation" renamed to "Markdown Conversion".

## 1.3.3 — 2026-06-04

- A branded splash now appears immediately at launch, while the app unpacks.
- Fixed the white window that flashed before the main window appeared.

## 1.3.2 — 2026-06-04

- Faster startup and a smaller download (~116 MB, down from ~129 MB).
- Added the animated startup splash.
- Window renamed to "File Preparation Toolkit".

## 1.3.1 — 2026-06-04

- The title bar and taskbar now show the owl logo in place of the feather.

## 1.3.0 — 2026-06-04

- New: convert any file type to clean Markdown, ideal for feeding documents into
  another AI. Email attachments are saved alongside and converted too.

## 1.2.7 — 2026-06-04

- Unzipping now requires an empty destination folder, re-checked when you click
  Unzip.

## 1.2.6 — 2026-06-04

- New: a Cancel button beside Convert and Unzip, which stops the run once the
  current file finishes.

## 1.2.5 — 2026-06-03

- Email conversion reverted to the dependable Outlook/Word path. The
  browser-based renderer fitted pages better but was unreliable on real email.

## 1.2.4 — 2026-06-03

- Fixed formatted emails stalling on remote images and tracking pixels. Remote
  fetches are now blocked during rendering, matching Outlook's own default.

## 1.2.3 — 2026-06-03

- Large email batches are roughly 12x faster: one browser is reused for the
  whole batch instead of launched per file.

## 1.2.2 — 2026-06-03

- Fixed two rendering flags that intermittently froze each email for up to two
  minutes, plus a 45-second fallback to Word for any file that misbehaves.

## 1.2.1 — 2026-06-03

- Wide email tables and text are now forced to fit the page instead of running
  off the edge.

## 1.1.0 — 2026-06-03

- Fixed email content running off the edge of the page.
- Auto-convert after unzipping now only touches files actually extracted from
  the zip.
- Existing PDFs are never overwritten; collisions get a numeric suffix.

## 1.0.0 — 2026-06-02

- First release: convert various file types to PDF, and mass-unzip folders with
  optional PDF conversion.
