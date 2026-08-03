"""
PDF Conversion Tool
Converts .eml, .msg, and .html files to PDF.
Email conversion uses Microsoft Outlook/Word via COM automation.
HTML conversion uses Microsoft Word via COM automation.
Requires: Microsoft Office installed on the machine.
"""

import os
import sys
import math
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
import tempfile
import shutil
import html as htmllib
import re
import json
import zipfile

# The running app's version, and the single source of truth for the "Check for
# Updates" feature: it is compared against the newest GitHub release tag to
# decide whether an update exists. MUST be bumped before every release build —
# leaving it stale makes the freshly-installed exe still believe it is the old
# version, so it offers the same update forever.
APP_VERSION = "1.7.1"

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

try:
    import win32com.client
    import pythoncom
    HAS_COM = True
except ImportError:
    HAS_COM = False


# ─────────────────────────────────────────────
# Outlook COM conversion
# ─────────────────────────────────────────────

def _ensure_outlook():
    """Return an Outlook Application COM object, raising clearly if unavailable."""
    if not HAS_COM:
        raise RuntimeError(
            "pywin32 is not installed.\n"
            "Run: pip install pywin32"
        )
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        return outlook
    except Exception as e:
        raise RuntimeError(
            f"Could not launch Outlook.\n"
            f"Make sure Microsoft Outlook is installed and has been set up.\n\n"
            f"Detail: {e}"
        )


def _safe_filename(name: str) -> str:
    """Strip characters that are illegal in Windows filenames."""
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip() or "attachment"


def _attachments_dir(out_path: Path) -> Path:
    """Return the attachments subfolder path for a given PDF output path."""
    return out_path.parent / out_path.stem


def _save_eml_attachments(src_path: Path, out_path: Path):
    """
    Extract real (non-inline) attachments from a .eml file and save them
    into a subfolder named after the PDF. Returns the number saved.
    """
    import email as emaillib
    import email.policy
    from email.header import decode_header, make_header

    def decode_str(value):
        if not value:
            return ""
        try:
            return str(make_header(decode_header(value)))
        except Exception:
            return str(value)

    with open(src_path, "rb") as f:
        msg = emaillib.message_from_binary_file(f, policy=emaillib.policy.compat32)

    saved = 0
    att_dir = None

    for part in msg.walk():
        if part.is_multipart():
            continue
        ct  = part.get_content_type()
        cd  = str(part.get("Content-Disposition") or "")
        cid = part.get("Content-ID", "").strip("<>")

        # Skip inline images (embedded in the body) and body text parts
        is_inline_image = ct.startswith("image/") and cid and "attachment" not in cd
        is_body = ct in ("text/html", "text/plain") and "attachment" not in cd
        if is_inline_image or is_body:
            continue

        payload = part.get_payload(decode=True)
        if not payload:
            continue

        fname = decode_str(part.get_filename() or "")
        if not fname:
            # Build a sensible name from content-type
            ext = ct.split("/")[-1].split(";")[0].strip()
            fname = f"attachment_{saved + 1}.{ext}"

        if att_dir is None:
            att_dir = _attachments_dir(out_path)
            att_dir.mkdir(parents=True, exist_ok=True)

        dest = att_dir / _safe_filename(fname)
        # Avoid overwriting if two attachments share a name
        if dest.exists():
            dest = att_dir / f"{dest.stem}_{saved + 1}{dest.suffix}"
        dest.write_bytes(payload)
        saved += 1

    return saved


def _save_msg_attachments(mail, out_path: Path) -> int:
    """
    Extract attachments from an open Outlook COM mail item and save them
    into a subfolder named after the PDF. Returns the number saved.
    """
    try:
        attachments = mail.Attachments
        count = attachments.Count
    except Exception:
        return 0

    if count == 0:
        return 0

    att_dir = _attachments_dir(out_path)
    att_dir.mkdir(parents=True, exist_ok=True)
    saved = 0

    for i in range(1, count + 1):   # COM collections are 1-indexed
        try:
            att = attachments.Item(i)

            # Attachment type 5 = OLE object, 6 = embedded message.
            # Rarely useful as standalone files; skip them.
            try:
                att_type = att.Type
                if att_type in (5, 6):
                    continue
            except Exception:
                pass

            fname = _safe_filename(att.FileName or f"attachment_{i}")
            dest = att_dir / fname
            if dest.exists():
                dest = att_dir / f"{dest.stem}_{i}{dest.suffix}"

            att.SaveAsFile(str(dest.resolve()))
            saved += 1
        except Exception:
            pass   # skip attachments that can't be saved

    # Remove the folder if nothing was actually saved
    if saved == 0:
        try:
            att_dir.rmdir()
        except Exception:
            pass

    return saved


# ─────────────────────────────────────────────
# Make rendered email content fit the PDF page
# ─────────────────────────────────────────────

# Injected into the <head> of every email HTML document before Word renders it.
# Constrains content to the printable width and forces long content to wrap so
# nothing spills off the right edge of the page.
_PRINT_CSS = (
    "<style>"
    "@page { size: auto; margin: 0.5in; }"
    "html, body { margin: 0; padding: 0; }"
    "img { max-width: 100% !important; height: auto !important; }"
    "table { max-width: 100% !important; }"
    "td, th { word-wrap: break-word; overflow-wrap: break-word; }"
    "pre { white-space: pre-wrap !important; word-wrap: break-word !important; }"
    "* { overflow-wrap: break-word; }"
    "</style>"
)


def _inject_print_css(html: str) -> str:
    """Insert the print stylesheet into the document's <head> (creating one if needed)."""
    if re.search(r"<head[^>]*>", html, re.IGNORECASE):
        return re.sub(r"(<head[^>]*>)", r"\1" + _PRINT_CSS, html,
                      count=1, flags=re.IGNORECASE)
    if re.search(r"<html[^>]*>", html, re.IGNORECASE):
        return re.sub(r"(<html[^>]*>)", r"\1<head>" + _PRINT_CSS + "</head>", html,
                      count=1, flags=re.IGNORECASE)
    return "<html><head>" + _PRINT_CSS + "</head>" + html + "</html>"


def _fit_word_doc_to_page(doc, word):
    """
    Resize an opened Word document so nothing runs off the right edge of the page.

    HTML emails frequently use fixed-width (pixel) layout tables and oversized
    images that Word imports at their literal width; without this they spill past
    the printable area and get clipped in the exported PDF. We narrow the margins,
    shrink any table wider than the text column to fit the page, and scale down
    oversized inline images.
    """
    wdAutoFitWindow         = 2   # WdAutoFitBehavior: fit table to text column
    wdPreferredWidthPercent = 2   # WdPreferredWidthType

    # Narrow, uniform margins give content more room.
    try:
        for section in doc.Sections:
            ps = section.PageSetup
            ps.LeftMargin   = word.InchesToPoints(0.5)
            ps.RightMargin  = word.InchesToPoints(0.5)
            ps.TopMargin    = word.InchesToPoints(0.5)
            ps.BottomMargin = word.InchesToPoints(0.5)
    except Exception:
        pass

    # Printable width (points) of the first section, used to detect overflow.
    try:
        ps = doc.Sections(1).PageSetup
        printable = ps.PageWidth - ps.LeftMargin - ps.RightMargin
    except Exception:
        printable = None

    def fit_tables(tables):
        for tbl in tables:
            try:
                too_wide = True
                if printable is not None:
                    try:
                        total = 0.0
                        for col in tbl.Columns:
                            total += col.Width
                        # +1pt tolerance to avoid distorting tables that already fit
                        too_wide = total > printable + 1
                    except Exception:
                        # Merged/mixed-width cells can't be measured — fit anyway
                        too_wide = True
                if too_wide:
                    tbl.PreferredWidthType = wdPreferredWidthPercent
                    tbl.PreferredWidth = 100
                    tbl.AutoFitBehavior(wdAutoFitWindow)
            except Exception:
                pass
            # Recurse into nested tables (common in email layouts)
            try:
                fit_tables(tbl.Tables)
            except Exception:
                pass

    try:
        fit_tables(doc.Tables)
    except Exception:
        pass

    # Scale down images wider than the printable area, preserving aspect ratio.
    if printable:
        try:
            for shape in doc.InlineShapes:
                try:
                    if shape.Width > printable:
                        ratio = printable / float(shape.Width)
                        shape.Height = shape.Height * ratio
                        shape.Width = printable
                except Exception:
                    pass
        except Exception:
            pass


def _msg_to_pdf(src_path: Path, out_path: Path, outlook, word):
    """Convert a .msg file to PDF and extract attachments via Outlook COM."""
    mail = outlook.Session.OpenSharedItem(str(src_path.resolve()))
    try:
        _print_mail_to_pdf(mail, out_path, word)
        _save_msg_attachments(mail, out_path)
    finally:
        mail.Close(0)   # 0 = olDiscard


def _parse_eml_to_html(src_path: Path) -> str:
    """
    Parse a .eml file using Python's stdlib email module and return a
    fully self-contained HTML document (inline images embedded as data URIs).
    """
    import email as emaillib
    import email.policy
    from email.header import decode_header, make_header
    import base64

    def decode_str(value):
        if not value:
            return ""
        try:
            return str(make_header(decode_header(value)))
        except Exception:
            return str(value)

    with open(src_path, "rb") as f:
        msg = emaillib.message_from_binary_file(f, policy=emaillib.policy.compat32)

    subject = decode_str(msg.get("Subject", "(No Subject)"))
    from_   = decode_str(msg.get("From", ""))
    to_     = decode_str(msg.get("To", ""))
    cc_     = decode_str(msg.get("CC", ""))
    date_   = decode_str(msg.get("Date", ""))

    html_body  = None
    plain_body = None
    cid_map    = {}
    attachments = []

    for part in msg.walk():
        ct  = part.get_content_type()
        cd  = str(part.get("Content-Disposition") or "")
        cid = part.get("Content-ID", "").strip("<>")

        if part.is_multipart():
            continue

        def _decode(p):
            payload = p.get_payload(decode=True)
            if payload is None:
                return ""
            charset = p.get_content_charset() or "utf-8"
            return payload.decode(charset, errors="replace")

        if ct == "text/html" and "attachment" not in cd and html_body is None:
            html_body = _decode(part)
        elif ct == "text/plain" and "attachment" not in cd and plain_body is None:
            plain_body = _decode(part)
        elif ct.startswith("image/") and cid:
            payload = part.get_payload(decode=True)
            if payload:
                b64 = base64.b64encode(payload).decode()
                cid_map[cid] = f"data:{ct};base64,{b64}"
        elif "attachment" in cd:
            fname = part.get_filename() or "attachment"
            attachments.append(decode_str(fname))

    # Inline CID images
    if html_body and cid_map:
        for cid, data_uri in cid_map.items():
            html_body = re.sub(
                re.escape(f"cid:{cid}"), data_uri, html_body, flags=re.IGNORECASE
            )

    # Build header block
    def hrow(label, value):
        if not value:
            return ""
        return (
            f'<tr>'
            f'<td style="font-weight:bold;color:#444;white-space:nowrap;padding:2px 8px 2px 0;vertical-align:top">{label}:</td>'
            f'<td style="color:#111">{htmllib.escape(value)}</td>'
            f'</tr>'
        )

    header_html = (
        '<div style="border-bottom:2px solid #0078d4;padding-bottom:10px;margin-bottom:16px;">'
        '<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:10pt">'
        + hrow("From", from_)
        + hrow("To", to_)
        + hrow("CC", cc_)
        + hrow("Date", date_)
        + f'<tr><td style="font-weight:bold;color:#444;padding:2px 8px 2px 0;vertical-align:top">Subject:</td>'
          f'<td style="font-size:13pt;font-weight:bold;color:#0078d4">{htmllib.escape(subject)}</td></tr>'
        + '</table></div>'
    )

    if attachments:
        names = htmllib.escape(", ".join(attachments))
        header_html += (
            f'<div style="margin-bottom:12px;font-family:Arial,sans-serif;font-size:9pt;color:#555;">'
            f'📎 Attachments: {names}</div>'
        )

    if html_body:
        if re.search(r"<body[^>]*>", html_body, re.IGNORECASE):
            full_html = re.sub(
                r"(<body[^>]*>)", r"\1" + header_html,
                html_body, count=1, flags=re.IGNORECASE,
            )
        else:
            full_html = (
                "<html><head><meta charset='utf-8'></head><body>"
                + header_html + html_body + "</body></html>"
            )
    else:
        body_escaped = htmllib.escape(plain_body or "(No message body)")
        full_html = (
            "<html><head><meta charset='utf-8'></head><body>"
            + header_html
            + f"<pre style='font-family:Arial,sans-serif;white-space:pre-wrap'>{body_escaped}</pre>"
            + "</body></html>"
        )

    return _inject_print_css(full_html)


def _eml_to_pdf(src_path: Path, out_path: Path, word):
    """
    Convert a .eml file to PDF and extract attachments.
    Parses the .eml directly in Python (no Outlook involvement) then
    exports via Word — avoids the 'Invalid path or URL' COM error entirely.
    """
    if word is None:
        raise RuntimeError(
            "Microsoft Word is required to convert .eml files on this version of Outlook.\n"
            "Please ensure Word is installed."
        )

    html_doc = _parse_eml_to_html(src_path)
    tmp_dir = Path(tempfile.mkdtemp())
    doc = None
    try:
        tmp_html = tmp_dir / "email.html"
        tmp_html.write_text(html_doc, encoding="utf-8")
        out_path.parent.mkdir(parents=True, exist_ok=True)

        doc = word.Documents.Open(
            str(tmp_html.resolve()),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        _fit_word_doc_to_page(doc, word)
        doc.ExportAsFixedFormat(
            OutputFileName=str(out_path.resolve()),
            ExportFormat=17,        # wdExportFormatPDF
            OpenAfterExport=False,
            OptimizeFor=0,
            Range=0,
            Item=0,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=0,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        shutil.rmtree(tmp_dir, ignore_errors=True)

    # Extract attachments into a subfolder alongside the PDF
    _save_eml_attachments(src_path, out_path)



def _mail_to_pdf_via_outlook(mail, out_path: Path) -> bool:
    """
    Method 1: Outlook's native ExportAsFixedFormat (Outlook 2010+).
    Returns True on success, False if the method simply doesn't exist on this version.
    Raises on any other error so the caller can see the real problem.
    """
    try:
        mail.ExportAsFixedFormat(
            0,                          # olPDF
            str(out_path.resolve()),
            False,                      # open after export
            False,                      # optimise for print
        )
        return True
    except AttributeError:
        return False   # method doesn't exist on this Outlook version
    except Exception as e:
        err = str(e).lower()
        # "unknown name" or "does not support" means the method isn't there
        if "unknown name" in err or "not supported" in err or "0x80020006" in err:
            return False
        raise   # real error — surface it


def _mail_to_pdf_via_word(mail, out_path: Path, word) -> bool:
    """
    Method 2: Pull HTML body from the Outlook COM item, write a temp HTML file,
    open it in the provided Word instance, export as PDF.
    Works on Office 2007+ without any printer or admin access.
    `word` is a shared Word.Application COM object (may be None if Word unavailable).
    Raises on failure so the caller can surface the real error message.
    """
    if word is None:
        raise RuntimeError(
            "Microsoft Word is not installed or could not be launched.\n"
            "Word is required as a fallback PDF renderer on this version of Outlook."
        )

    tmp_dir = Path(tempfile.mkdtemp())
    doc = None
    try:
        # ── Pull content from the COM mail item ───────────────────────────
        html_body = None
        plain_body = None
        try:
            html_body = mail.HTMLBody
        except Exception:
            pass
        if not html_body:
            try:
                plain_body = mail.Body
            except Exception:
                pass

        def safe_get(attr):
            try:
                return str(getattr(mail, attr) or "").strip()
            except Exception:
                return ""

        subject  = safe_get("Subject") or "(No Subject)"
        from_    = safe_get("SenderName") or safe_get("SenderEmailAddress")
        to_      = safe_get("To")
        cc_      = safe_get("CC")
        try:
            received = str(mail.ReceivedTime)
        except Exception:
            received = ""

        # ── Build header block ────────────────────────────────────────────
        def hrow(label, value):
            if not value:
                return ""
            return (
                f'<tr>'
                f'<td style="font-weight:bold;color:#444;white-space:nowrap;padding:2px 8px 2px 0;vertical-align:top">{label}:</td>'
                f'<td style="color:#111">{htmllib.escape(value)}</td>'
                f'</tr>'
            )

        header_html = (
            '<div style="border-bottom:2px solid #0078d4;padding-bottom:10px;margin-bottom:16px;">'
            '<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:10pt">'
            + hrow("From", from_)
            + hrow("To", to_)
            + hrow("CC", cc_)
            + hrow("Date", received)
            + f'<tr><td style="font-weight:bold;color:#444;padding:2px 8px 2px 0;vertical-align:top">Subject:</td>'
              f'<td style="font-size:13pt;font-weight:bold;color:#0078d4">{htmllib.escape(subject)}</td></tr>'
            + '</table></div>'
        )

        if html_body:
            if re.search(r"<body[^>]*>", html_body, re.IGNORECASE):
                full_html = re.sub(
                    r"(<body[^>]*>)",
                    r"\1" + header_html,
                    html_body, count=1, flags=re.IGNORECASE,
                )
            else:
                full_html = (
                    "<html><head><meta charset='utf-8'></head><body>"
                    + header_html + html_body + "</body></html>"
                )
        else:
            body_escaped = htmllib.escape(plain_body or "(No message body)")
            full_html = (
                "<html><head><meta charset='utf-8'></head><body>"
                + header_html
                + f"<pre style='font-family:Arial,sans-serif;white-space:pre-wrap'>{body_escaped}</pre>"
                + "</body></html>"
            )

        # ── Write temp HTML and open in Word ──────────────────────────────
        full_html = _inject_print_css(full_html)
        tmp_html = tmp_dir / "email.html"
        tmp_html.write_text(full_html, encoding="utf-8")
        out_path.parent.mkdir(parents=True, exist_ok=True)

        doc = word.Documents.Open(
            str(tmp_html.resolve()),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )

        doc.ExportAsFixedFormat(
            OutputFileName=str(out_path.resolve()),
            ExportFormat=17,        # wdExportFormatPDF
            OpenAfterExport=False,
            OptimizeFor=0,          # wdExportOptimizeForPrint
            Range=0,                # wdExportAllDocument
            Item=0,                 # wdExportDocumentContent
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=0,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )
        return True

    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        shutil.rmtree(tmp_dir, ignore_errors=True)


def _print_mail_to_pdf(mail, out_path: Path, word):
    """
    Convert an open Outlook mail item to PDF.
    Tries Outlook-native export first, falls back to Word.
    `word` is a shared Word.Application COM object (or None).
    Raises RuntimeError with a descriptive message on failure.
    """
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Method 1 — Word HTML rendering (preferred).
    # Outlook renders HTML mail with the Word engine internally, so fidelity is
    # equivalent to its native export, but going through Word ourselves lets us
    # fit wide tables/images to the page so content doesn't run off the edge.
    word_err = None
    if word is not None:
        try:
            _mail_to_pdf_via_word(mail, out_path, word)
            return
        except Exception as e:
            word_err = e   # fall back to Outlook native below

    # Method 2 — Outlook native export (2010+); fallback when Word is
    # unavailable or the Word path failed.
    try:
        if _mail_to_pdf_via_outlook(mail, out_path):
            return
    except Exception as e:
        raise RuntimeError(f"Outlook PDF export failed: {e}")

    if word_err is not None:
        raise RuntimeError(f"Word PDF export failed: {word_err}")
    raise RuntimeError(
        "No available method to export this message to PDF "
        "(Outlook native export is unsupported on this version and Word is unavailable)."
    )


def _visio_to_pdf(src_path: Path, out_path: Path, visio):
    """
    Convert a .vsd or .vsdx file to PDF via Visio COM.
    Visio's Document.ExportAsFixedFormat handles both formats natively.
    """
    if visio is None:
        raise RuntimeError(
            "Microsoft Visio is required for Visio conversion.\n"
            "Please ensure Visio is installed."
        )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc = None
    try:
        doc = visio.Documents.Open(str(src_path.resolve()))

        # visFixedFormatPDF = 1
        doc.ExportAsFixedFormat(
            1,                          # visFixedFormatPDF
            str(out_path.resolve()),    # output path
            1,                          # Intent: print (visIntentPrint)
            0,                          # All pages
        )
    except Exception as e:
        raise RuntimeError(f"Visio to PDF conversion failed: {e}")
    finally:
        if doc is not None:
            try:
                doc.Close()
            except Exception:
                pass


def _fit_excel_sheets_to_page(wb, excel):
    """Make every worksheet print legibly before exporting to PDF.

    By default Excel prints at 100%, so any table wider than the paper is sliced
    across pages column-wise — which looks messy and is hard for a downstream
    reader (human or AI) to follow. For each sheet we instead:
      • scale all columns onto ONE page wide (FitToPagesWide=1) so a table is
        never cut left-to-right, while letting rows flow onto as many pages tall
        as needed (FitToPagesTall=False) so nothing is crushed vertically;
      • switch to landscape when the used range is wider than it is tall, which
        gives wide tables more room before they have to scale down;
      • trim margins to use more of the page.
    All settings are best-effort per sheet — a sheet that can't take a setting is
    skipped rather than failing the whole conversion.
    """
    try:
        sheets = wb.Worksheets
        count = sheets.Count
    except Exception:
        return
    for i in range(1, count + 1):
        try:
            ws = sheets(i)
            ps = ws.PageSetup
            # Fit columns to one page wide; paginate rows naturally.
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = False
        except Exception:
            continue
        # Landscape for wide content (more width before Excel scales it down).
        try:
            used = ws.UsedRange
            if used.Width > used.Height:
                ps.Orientation = 2          # xlLandscape
        except Exception:
            pass
        # Narrow margins so more of the table fits at a larger scale.
        try:
            m = excel.InchesToPoints(0.25)
            hm = excel.InchesToPoints(0.1)
            ps.LeftMargin = ps.RightMargin = m
            ps.TopMargin = ps.BottomMargin = m
            ps.HeaderMargin = ps.FooterMargin = hm
        except Exception:
            pass


def _excel_to_pdf(src_path: Path, out_path: Path, excel):
    """
    Convert a .xls, .xlsx, .xlsm, or .xlsb file to PDF via Excel COM.
    Prints all sheets to a single PDF.
    """
    if excel is None:
        raise RuntimeError(
            "Microsoft Excel is required for Excel conversion.\n"
            "Please ensure Excel is installed."
        )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb = None
    try:
        wb = excel.Workbooks.Open(
            str(src_path.resolve()),
            UpdateLinks=False,
            ReadOnly=True,
            AddToMru=False,
        )

        # Scale every sheet to fit one page wide so wide tables aren't sliced.
        _fit_excel_sheets_to_page(wb, excel)

        # xlTypePDF = 0, xlQualityStandard = 0
        wb.ExportAsFixedFormat(
            0,                          # xlTypePDF
            str(out_path.resolve()),    # output path
            0,                          # xlQualityStandard
            True,                       # IncludeDocProperties
            False,                      # IgnorePrintAreas
            OpenAfterPublish=False,
        )
    except Exception as e:
        raise RuntimeError(f"Excel to PDF conversion failed: {e}")
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass


def _word_to_pdf(src_path: Path, out_path: Path, word):
    """
    Convert a .doc or .docx file to PDF via Word COM.
    """
    if word is None:
        raise RuntimeError(
            "Microsoft Word is required for Word document conversion.\n"
            "Please ensure Word is installed."
        )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc = None
    try:
        doc = word.Documents.Open(
            str(src_path.resolve()),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        _fit_word_doc_to_page(doc, word)
        doc.ExportAsFixedFormat(
            OutputFileName=str(out_path.resolve()),
            ExportFormat=17,        # wdExportFormatPDF
            OpenAfterExport=False,
            OptimizeFor=0,          # wdExportOptimizeForPrint
            Range=0,                # wdExportAllDocument
            Item=0,                 # wdExportDocumentContent
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=0,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )
    except Exception as e:
        raise RuntimeError(f"Word document to PDF conversion failed: {e}")
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass


# Plain-text report dumps produced by legacy / ERP / mainframe systems. Word
# picks an importer by extension and has none registered for these, so they go
# through a temp .txt copy (see _text_report_to_pdf).
TEXT_REPORT_EXTS = (".rep", ".rpt", ".prn", ".lst")

# Consolas' advance width is ~0.55 em, so N characters at F points occupy about
# N * 0.55 * F points. Used to pick a font size that fits the widest line.
_MONO_ADVANCE     = 0.55
_REPORT_FONT_MIN  = 5.0
_REPORT_FONT_MAX  = 11.0
# Wider than this many columns and the report gets a landscape page before we
# start shrinking type — a readable landscape report beats a 6pt portrait one.
_REPORT_LANDSCAPE_COLS = 96


def _looks_binary(raw: bytes) -> bool:
    """True if a byte sample is not plausibly plain text.

    Report dumps are ASCII, so a NUL byte means we were handed a binary file
    wearing a report extension — a compiled Crystal Reports .rpt, say. Pushing
    one through Word would emit pages of mojibake instead of failing, so
    callers reject it up front.
    """
    if b"\x00" in raw:
        return True
    if not raw:
        return False
    # Tab/BS/LF/FF/CR/ESC plus printable ASCII, and any high byte (code pages).
    textish = bytes((7, 8, 9, 10, 12, 13, 27)) + bytes(range(0x20, 0x100))
    return sum(b not in textish for b in raw) / len(raw) > 0.30


def _read_report_text(src_path: Path) -> str:
    """Decode a report dump, tolerating the code pages legacy systems emit."""
    raw = src_path.read_bytes()
    if _looks_binary(raw[:8192]):
        raise RuntimeError(
            f"{src_path.name} looks like a binary file, not a plain-text report. "
            "Only text report dumps can be converted."
        )
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", errors="replace")


def _fit_report_to_page(doc, word, longest_line):
    """Lay an imported report out so no line wraps.

    Word imports plain text in the Normal style — a proportional font with
    paragraph spacing — which destroys the column alignment a fixed-width
    report depends on. This forces monospace, removes the inter-line spacing
    that would otherwise double-space the whole report, and sizes the page and
    type to the widest line.
    """
    wdOrientLandscape = 1
    wdLineSpaceSingle = 0

    try:
        doc.Content.Font.Name = "Consolas"
    except Exception:
        pass

    # Word's Normal style adds space between paragraphs; every report line is
    # its own paragraph, so without this the output is double-spaced.
    try:
        pf = doc.Content.ParagraphFormat
        pf.SpaceBefore = 0
        pf.SpaceAfter = 0
        pf.LineSpacingRule = wdLineSpaceSingle
    except Exception:
        pass

    try:
        for section in doc.Sections:
            ps = section.PageSetup
            if longest_line > _REPORT_LANDSCAPE_COLS:
                ps.Orientation = wdOrientLandscape
            ps.LeftMargin = ps.RightMargin = word.InchesToPoints(0.4)
            ps.TopMargin = ps.BottomMargin = word.InchesToPoints(0.4)
    except Exception:
        pass

    if longest_line <= 0:
        return
    try:
        ps = doc.Sections(1).PageSetup
        printable = ps.PageWidth - ps.LeftMargin - ps.RightMargin
        size = printable / (_MONO_ADVANCE * longest_line)
        doc.Content.Font.Size = round(
            max(_REPORT_FONT_MIN, min(_REPORT_FONT_MAX, size)), 1)
    except Exception:
        pass


def _text_report_to_pdf(src_path: Path, out_path: Path, word):
    """Convert a plain-text report dump to PDF via Word COM.

    Handles TEXT_REPORT_EXTS and extensionless text files. The text is copied
    to a temp .txt because that is the only reliable way to get Word to open
    an unknown extension as plain text rather than running format detection.
    Form feeds are left intact — Word maps them to page breaks, so the
    report's own pagination survives into the PDF.
    """
    if word is None:
        raise RuntimeError(
            "Microsoft Word is required to convert text reports.\n"
            "Please ensure Word is installed."
        )

    text = _read_report_text(src_path)
    longest = max((len(ln.rstrip("\r")) for ln in text.split("\n")), default=0)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_dir = Path(tempfile.mkdtemp(prefix="apfp_report_"))
    tmp_txt = tmp_dir / ((src_path.stem or "report") + ".txt")
    doc = None
    try:
        # Write the BOM so Word detects UTF-8 instead of guessing a code page.
        tmp_txt.write_text(text, encoding="utf-8-sig", newline="")
        doc = word.Documents.Open(
            str(tmp_txt.resolve()),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        _fit_report_to_page(doc, word, longest)
        doc.ExportAsFixedFormat(
            OutputFileName=str(out_path.resolve()),
            ExportFormat=17,        # wdExportFormatPDF
            OpenAfterExport=False,
            OptimizeFor=0,          # wdExportOptimizeForPrint
            Range=0,                # wdExportAllDocument
            Item=0,                 # wdExportDocumentContent
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=0,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )
    except Exception as e:
        raise RuntimeError(f"Text report to PDF conversion failed: {e}")
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        shutil.rmtree(tmp_dir, ignore_errors=True)


def _ppt_to_pdf(src_path: Path, out_path: Path, powerpoint):
    """
    Convert a .ppt, .pptx, .pps, or .ppsx file to PDF via PowerPoint COM.
    """
    if powerpoint is None:
        raise RuntimeError(
            "Microsoft PowerPoint is required for PowerPoint conversion.\n"
            "Please ensure PowerPoint is installed."
        )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs = None
    try:
        prs = powerpoint.Presentations.Open(
            str(src_path.resolve()),
            ReadOnly=True,
            Untitled=False,
            WithWindow=False,   # don't show a window
        )

        # ppFixedFormatTypePDF = 2, ppFixedFormatIntentPrint = 2
        prs.ExportAsFixedFormat(
            str(out_path.resolve()),
            2,      # ppFixedFormatTypePDF
            Intent=2,           # ppFixedFormatIntentPrint
            FrameSlides=False,
            HandoutOrder=1,
            OutputType=1,       # ppPrintOutputSlides
            PrintHiddenSlides=False,
            PrintRange=None,
            RangeType=1,        # ppPrintAll
            SlideShowName="",
            IncludeDocProperties=True,
            KeepIRMSettings=True,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )
    except Exception as e:
        raise RuntimeError(f"PowerPoint to PDF conversion failed: {e}")
    finally:
        if prs is not None:
            try:
                prs.Close()
            except Exception:
                pass


def _unique_pdf_path(directory: Path, stem: str) -> Path:
    """Return a PDF path in directory that does not collide with an existing file.

    Tries '<stem>.pdf' first, then '<stem>_2.pdf', '<stem>_3.pdf', … so an
    existing PDF is never overwritten.
    """
    candidate = directory / f"{stem}.pdf"
    counter = 2
    while candidate.exists():
        candidate = directory / f"{stem}_{counter}.pdf"
        counter += 1
    return candidate


def _unique_output_path(directory: Path, stem: str, ext: str) -> Path:
    """Return a '<stem><ext>' path in directory that doesn't collide with an
    existing file (suffixing _2, _3, … as needed). ext includes the dot."""
    candidate = directory / f"{stem}{ext}"
    counter = 2
    while candidate.exists():
        candidate = directory / f"{stem}_{counter}{ext}"
        counter += 1
    return candidate


def _resolved_set(paths) -> set:
    """Resolve an iterable of paths to a set of canonical Paths, for safe
    membership checks. Unresolvable paths are skipped."""
    out = set()
    for p in paths:
        try:
            out.add(Path(p).resolve())
        except Exception:
            pass
    return out


def _delete_original_file(src: Path, allowed: set) -> bool:
    """HARD SAFEGUARD for every 'delete originals after operation' feature.

    Delete `src` ONLY if its resolved path is in `allowed` — the explicit set of
    files the user actually supplied to this operation. This guarantees the app
    can never delete a file that wasn't part of the upload, even if an output
    folder happens to contain other, unrelated files. Returns True iff deleted.

    Callers must build `allowed` from the operation's own input list (e.g.
    `_resolved_set(self._ai_files)`), never from a directory scan.
    """
    try:
        if src.resolve() not in allowed:
            return False           # not part of the upload — refuse to delete
    except Exception:
        return False
    try:
        src.unlink()
        return True
    except Exception:
        return False


# ─────────────────────────────────────────────
# Legacy Office → modern Office (one-to-one, by family)
#
# Each legacy (97-2003 / earlier) extension is keyed to exactly one Office
# application, its modern OpenXML extension, and the SaveAs file-format code.
# Keying by family is what guarantees a .doc can only ever become a Word file,
# a .xls only an Excel file, etc. — the types can never be crossed.
#   Word        wdFormatXMLDocument=12, wdFormatXMLTemplate=14
#   Excel       xlOpenXMLWorkbook=51,   xlOpenXMLTemplate=54
#   PowerPoint  ppSaveAsOpenXMLPresentation=24, ppSaveAsOpenXMLShow=28,
#               ppSaveAsOpenXMLTemplate=26
LEGACY_OFFICE = {
    # Word family  → .docx / .dotx
    ".doc": ("word", ".docx", 12),
    ".dot": ("word", ".dotx", 14),
    # Excel family → .xlsx / .xltx
    ".xls": ("excel", ".xlsx", 51),
    ".xlt": ("excel", ".xltx", 54),
    # PowerPoint family → .pptx / .ppsx / .potx
    ".ppt": ("powerpoint", ".pptx", 24),
    ".pps": ("powerpoint", ".ppsx", 28),
    ".pot": ("powerpoint", ".potx", 26),
}


# Executable name + COM ProgID for each Office family.
OFFICE_INFO = {
    "word":       ("winword.exe", "Word.Application"),
    "excel":      ("excel.exe",   "Excel.Application"),
    "powerpoint": ("powerpnt.exe", "PowerPoint.Application"),
    "visio":      ("visio.exe",   "Visio.Application"),
}


def _office_pids(exe_name: str) -> set:
    """Return the set of PIDs of running processes whose exe base name matches
    (case-insensitive). Used to identify the dedicated Office instance we launch
    so we can later kill ONLY ours, never the user's open Office windows."""
    import ctypes
    from ctypes import wintypes

    TH32CS_SNAPPROCESS = 0x00000002

    class PROCESSENTRY32(ctypes.Structure):
        _fields_ = [
            ("dwSize", wintypes.DWORD), ("cntUsage", wintypes.DWORD),
            ("th32ProcessID", wintypes.DWORD),
            ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_size_t)),
            ("th32ModuleID", wintypes.DWORD), ("cntThreads", wintypes.DWORD),
            ("th32ParentProcessID", wintypes.DWORD),
            ("pcPriClassBase", ctypes.c_long), ("dwFlags", wintypes.DWORD),
            ("szExeFile", ctypes.c_char * 260),
        ]

    k32 = ctypes.windll.kernel32
    pids = set()
    target = exe_name.lower()
    snap = k32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if snap == -1:
        return pids
    try:
        entry = PROCESSENTRY32()
        entry.dwSize = ctypes.sizeof(PROCESSENTRY32)
        ok = k32.Process32First(snap, ctypes.byref(entry))
        while ok:
            if entry.szExeFile.decode("ascii", "ignore").lower() == target:
                pids.add(int(entry.th32ProcessID))
            ok = k32.Process32Next(snap, ctypes.byref(entry))
    finally:
        k32.CloseHandle(snap)
    return pids


def _terminate_pid(pid: int):
    """Force-terminate a single process by PID (best effort, ignores errors)."""
    try:
        import ctypes
        PROCESS_TERMINATE = 0x0001
        h = ctypes.windll.kernel32.OpenProcess(PROCESS_TERMINATE, False, int(pid))
        if h:
            ctypes.windll.kernel32.TerminateProcess(h, 1)
            ctypes.windll.kernel32.CloseHandle(h)
    except Exception:
        pass


# Per-file conversion watchdog: if a single file hasn't finished within this
# many seconds, its dedicated Office instance is assumed stuck (a hidden modal
# dialog) and force-killed so the batch can skip it and continue. Set generously
# — even very large/complex legacy files convert in well under a minute — so a
# genuinely slow conversion is never aborted by mistake. (The manual Cancel
# button still force-stops instantly.)
OFFICE_FILE_TIMEOUT = 300   # 5 minutes


class _FileWatchdog:
    """Bounds one file's conversion so a single stuck file can't hang a batch.

    Used as a context manager around a single `convert`/`modernize` call. If the
    body hasn't finished within OFFICE_FILE_TIMEOUT the tracked PIDs — which are
    only ever the instances *we* launched — are force-killed, which makes the
    blocked COM call raise so the loop can record a failure and move on.

    `fired` lets the caller report "timed out" instead of the incidental COM
    error the kill produced. With no PIDs to kill (e.g. a PowerPoint that
    attached to the user's own instance) the watchdog stays inert rather than
    risk touching a process we didn't start.
    """

    def __init__(self, pids):
        self._pids = set(pids or ())
        self._done = threading.Event()
        self._thread = None
        self.fired = False

    def __enter__(self):
        if self._pids:
            self._thread = threading.Thread(target=self._run, daemon=True)
            self._thread.start()
        return self

    def _run(self):
        if not self._done.wait(OFFICE_FILE_TIMEOUT):
            self.fired = True
            for pid in self._pids:
                _terminate_pid(pid)

    def __exit__(self, *exc_info):
        self._done.set()
        if self._thread is not None:
            self._thread.join(timeout=2)
        return False


def _launch_office_isolated(family: str):
    """Launch a DEDICATED, isolated Office instance for `family` via DispatchEx
    and return (app, pids). `pids` is the set of newly-created Office process
    IDs — i.e. only the instance we just started — so a stuck conversion can be
    force-killed later without ever touching the user's own Office windows.

    If no new process appears (e.g. PowerPoint attached to the user's existing
    instance), `pids` is empty and that instance is deliberately never killed.
    """
    exe, progid = OFFICE_INFO[family]
    before = _office_pids(exe)
    app = win32com.client.DispatchEx(progid)
    try:
        if family == "powerpoint":
            app.WindowState = 2          # ppWindowMinimized (PP can't be fully hidden)
        else:
            app.Visible = False
    except Exception:
        pass
    # Disable macros and suppress modal prompts (the usual cause of hangs).
    try:
        app.AutomationSecurity = 3       # msoAutomationSecurityForceDisable
    except Exception:
        pass
    try:
        app.DisplayAlerts = 1 if family == "powerpoint" else False
    except Exception:
        pass
    # Never pop an "install this feature?" dialog — it blocks the hidden app.
    try:
        app.FeatureInstall = 0           # msoFeatureInstallNone (raise instead of prompt)
    except Exception:
        pass
    # Word/Excel-specific dialogs that otherwise hang on larger/complex legacy
    # files: update-links, older-format encoding/convert, and overwrite prompts.
    if family == "word":
        for opt, val in (("UpdateLinksAtOpen", False),
                         ("ConfirmConversions", False),
                         ("DoNotPromptForConvert", True),
                         ("WarnBeforeSavingPrintingSendingMarkup", False),
                         ("SaveNormalPrompt", False),
                         # Background repagination reflows the document on every
                         # edit we make while fitting it to the page. Export
                         # repaginates anyway, so this is pure wasted work.
                         ("Pagination", False)):
            try:
                setattr(app.Options, opt, val)
            except Exception:
                pass
    elif family == "excel":
        for prop, val in (("AskToUpdateLinks", False),
                          ("AlertBeforeOverwriting", False),
                          # Stop workbook event handlers (Workbook_Open and
                          # friends) from running in our hidden instance.
                          ("EnableEvents", False)):
            try:
                setattr(app, prop, val)
            except Exception:
                pass
    # Nothing is on screen in a hidden instance, so painting is wasted time.
    try:
        app.ScreenUpdating = False
    except Exception:
        pass
    pids = _office_pids(exe) - before
    return app, pids


def _modernize_office_file(src_path: Path, out_path: Path, family: str, fmt: int, app) -> None:
    """Convert one legacy Office file to its modern OpenXML equivalent at
    `out_path`, using the already-launched `app` for `family`. Routing is fixed
    by the caller's family/format, so types can never be crossed. The original
    is opened read-only and never modified.
    """
    if app is None:
        need = {"word": "Word", "excel": "Excel", "powerpoint": "PowerPoint"}[family]
        raise RuntimeError(f"Microsoft {need} is required for this file type.")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    src = str(src_path.resolve())
    dst = str(out_path.resolve())

    if family == "word":
        doc = None
        try:
            doc = app.Documents.Open(
                src, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False)
            doc.SaveAs2(dst, FileFormat=fmt)
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass

    elif family == "excel":
        wb = None
        try:
            wb = app.Workbooks.Open(
                src, UpdateLinks=False, ReadOnly=True, AddToMru=False)
            wb.SaveAs(dst, FileFormat=fmt)
        finally:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass

    else:  # powerpoint
        prs = None
        try:
            prs = app.Presentations.Open(
                src, ReadOnly=True, Untitled=False, WithWindow=False)
            prs.SaveAs(dst, fmt)
        finally:
            if prs is not None:
                try:
                    prs.Close()
                except Exception:
                    pass


# ─────────────────────────────────────────────
# Password Removal — strip the open-password (encryption) from Office files.
#
# Office's "Encrypt with Password" wraps the document in a standard encrypted
# container (ECMA-376 for .docx/.xlsx/.pptx; the OLE crypto streams for legacy
# .doc/.xls/.ppt). We decrypt that container with msoffcrypto-tool — pure Python,
# so NO Office process is launched, nothing can pop a hidden modal prompt, and a
# wrong password fails instantly and cleanly (instead of Excel's COM Open, which
# blocks on an interactive password dialog). The decrypted bytes ARE the original
# document, so the unlocked copy keeps the exact same format/extension.

# Office file types we accept for password removal. (msoffcrypto auto-detects the
# real format from the file contents; this set just filters what can be added.)
PWD_EXTS = {
    ".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".rtf",
    ".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm",
    ".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".pot", ".potx",
}


class _BadPassword(Exception):
    """Raised when the supplied password is wrong for an encrypted file. Kept
    distinct so the UI can show a specific 'incorrect password' message and
    re-prompt for the same file."""


class _NotEncrypted(Exception):
    """Raised when a file has no open-password encryption to remove (e.g. it
    isn't protected, or only carries a write-reservation/edit password, which is
    not encryption and can't be stripped this way)."""


def _office_is_encrypted(src_path: Path) -> bool:
    """True if `src_path` is a password-encrypted Office file. Never raises —
    anything unreadable/unrecognized is reported as not-encrypted."""
    try:
        import msoffcrypto
        with open(src_path, "rb") as f:
            return bool(msoffcrypto.OfficeFile(f).is_encrypted())
    except Exception:
        return False


# The only formats that can carry open-password encryption: the Office
# containers msoffcrypto understands, plus ZIP archives. Checking anything else
# means opening and parsing a file that can never be protected, which is pure
# overhead on batches of images, text, or email.
ENCRYPTABLE_EXTS = PWD_EXTS | {".zip"}


def _is_password_protected(src_path: Path) -> bool:
    """True if `src_path` is locked with a password and therefore can't be
    processed by the other tools. Covers password-encrypted Office files
    (.docx/.xlsx/.pptx and legacy .doc/.xls/.ppt) and password-encrypted ZIPs.
    Best-effort and never raises — anything unrecognized is treated as not
    protected. Used by every tool EXCEPT Password Removal to skip such files and
    point the user at the Password Removal tool.

    Judged by extension first, so a protected file given a misleading extension
    (an encrypted .docx renamed to .txt) reads as unprotected — it would fail
    later in conversion anyway, and every tool already filters by extension."""
    if src_path.suffix.lower() not in ENCRYPTABLE_EXTS:
        return False
    try:
        if _office_is_encrypted(src_path):
            return True
        # An Office-encrypted file is an OLE container, not a ZIP, so this only
        # ever flags genuinely password-protected archives (a normal .docx/.zip
        # has its entry encryption bit clear).
        if zipfile.is_zipfile(src_path):
            with zipfile.ZipFile(src_path) as z:
                if any(zi.flag_bits & 0x1 for zi in z.infolist()):
                    return True
    except Exception:
        pass
    return False


def _password_skip_note(protected, advice=None) -> str:
    """Trailing advisory block for a completion dialog listing files skipped
    because they're password-protected, or '' if there were none. Shown below
    the normal results so the user knows what to do next. `advice` overrides the
    default guidance (the default points at the Password Removal tool, which only
    applies to Office documents)."""
    if not protected:
        return ""
    names = "\n".join(f"   • {n}" for n in protected)
    if advice is None:
        advice = ("These files are locked with a password. Remove it first using the "
                  "\"Password Removal\" tool, then run them through this tool again.")
    return ("\n\n⚠️  Skipped — password-protected (could not be opened):\n"
            f"{names}\n\n{advice}")


def _strip_office_password(src_path: Path, out_path: Path, password: str) -> None:
    """Decrypt the password-protected Office file at `src_path` to `out_path`
    using `password`. Pure Python (msoffcrypto) — no Office process, no prompts.
    The original file is never modified. Raises _BadPassword on a wrong password
    and _NotEncrypted if there's no encryption to remove."""
    try:
        import msoffcrypto
        from msoffcrypto.exceptions import InvalidKeyError
    except ImportError as e:
        raise RuntimeError(
            "The 'msoffcrypto-tool' package is required for Password Removal.\n"
            "Install it with:  pip install msoffcrypto-tool"
        ) from e

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(src_path, "rb") as f:
        of = msoffcrypto.OfficeFile(f)
        try:
            encrypted = of.is_encrypted()
        except Exception:
            encrypted = True   # let the decrypt attempt surface the real error
        if not encrypted:
            raise _NotEncrypted()
        try:
            of.load_key(password=password)
            with open(out_path, "wb") as g:
                of.decrypt(g)
        except InvalidKeyError as e:
            # Remove any partial output left by a failed decrypt.
            try:
                if out_path.exists():
                    out_path.unlink()
            except Exception:
                pass
            raise _BadPassword(str(e))
        except Exception:
            try:
                if out_path.exists():
                    out_path.unlink()
            except Exception:
                pass
            raise


# ─────────────────────────────────────────────
# AI Preparation — convert any file type to LLM-friendly Markdown
# via Microsoft's markitdown.
# ─────────────────────────────────────────────

_MARKITDOWN = None


def _get_markitdown():
    """Return a shared MarkItDown instance, importing it lazily so the rest of
    the app works even if markitdown isn't installed. Raises a clear error if
    the package is missing."""
    global _MARKITDOWN
    if _MARKITDOWN is None:
        try:
            from markitdown import MarkItDown
        except ImportError as e:
            raise RuntimeError(
                "The 'markitdown' package is required for AI Preparation.\n"
                "Install it with:  pip install \"markitdown[all]\""
            ) from e
        _MARKITDOWN = MarkItDown()
    return _MARKITDOWN


def _to_markdown(src_path: Path, out_dir: Path) -> Path:
    """Convert any supported file to a Markdown (.md) file in out_dir and return
    its path. Output is plain Markdown — easy for an LLM to read. Never
    overwrites an existing file."""
    md = _get_markitdown()
    result = md.convert(str(src_path.resolve()))
    text = (result.text_content or "").strip()
    # Strip NUL and other C0 control characters (keep tab/newline/carriage-return)
    # so the .md is always valid text — some sources can emit stray control bytes,
    # which otherwise make editors treat the file as binary.
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    out_path = _unique_output_path(out_dir, src_path.stem, ".md")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    # A short source note helps the downstream LLM know where the text came from.
    header = f"<!-- Source file: {src_path.name} -->\n\n"
    out_path.write_text(header + text + "\n", encoding="utf-8")
    return out_path


# Image attachment types are skipped during AI Preparation (not useful as text).
_AI_IMAGE_EXTS = {
    ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif",
    ".webp", ".svg", ".ico", ".heic", ".heif",
}


def _unique_dir(parent: Path, name: str) -> Path:
    """Return a folder path under parent that doesn't already exist (name, name_2, …)."""
    candidate = parent / name
    counter = 2
    while candidate.exists():
        candidate = parent / f"{name}_{counter}"
        counter += 1
    return candidate


def _extract_eml_attachments(src_path: Path, dest_dir: Path) -> list:
    """Save non-image attachments from a .eml into dest_dir (created lazily).
    Returns the list of saved file paths. No Outlook/COM needed."""
    import email as emaillib
    import email.policy
    from email.header import decode_header, make_header

    def decode_str(value):
        if not value:
            return ""
        try:
            return str(make_header(decode_header(value)))
        except Exception:
            return str(value)

    with open(src_path, "rb") as f:
        msg = emaillib.message_from_binary_file(f, policy=emaillib.policy.compat32)

    saved = []
    for part in msg.walk():
        if part.is_multipart():
            continue
        ct  = part.get_content_type()
        cd  = str(part.get("Content-Disposition") or "")
        cid = part.get("Content-ID", "").strip("<>")

        is_inline_image = ct.startswith("image/") and cid and "attachment" not in cd
        is_body = ct in ("text/html", "text/plain") and "attachment" not in cd
        if is_inline_image or is_body:
            continue

        payload = part.get_payload(decode=True)
        if not payload:
            continue

        fname = decode_str(part.get_filename() or "")
        if not fname:
            ext = ct.split("/")[-1].split(";")[0].strip() or "bin"
            fname = f"attachment_{len(saved) + 1}.{ext}"

        # Skip image attachments
        if ct.startswith("image/") or Path(fname).suffix.lower() in _AI_IMAGE_EXTS:
            continue

        safe = _safe_filename(fname)
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = _unique_output_path(dest_dir, Path(safe).stem, Path(safe).suffix)
        dest.write_bytes(payload)
        saved.append(dest)

    return saved


def _extract_msg_attachments(src_path: Path, dest_dir: Path) -> list:
    """Save non-image attachments from a .msg into dest_dir (created lazily),
    using the pure-Python extract_msg library (no Outlook/COM).
    Returns the list of saved file paths."""
    try:
        import extract_msg
    except ImportError:
        return []   # attachment extraction unavailable; body still converts

    saved = []
    msg = extract_msg.Message(str(src_path.resolve()))
    try:
        for att in msg.attachments:
            try:
                name = att.getFilename() or "attachment"
                data = att.data
                if not isinstance(data, (bytes, bytearray)):
                    continue   # embedded message / unsupported attachment type
                if Path(name).suffix.lower() in _AI_IMAGE_EXTS:
                    continue
                safe = _safe_filename(name)
                dest_dir.mkdir(parents=True, exist_ok=True)
                dest = _unique_output_path(dest_dir, Path(safe).stem, Path(safe).suffix)
                dest.write_bytes(bytes(data))
                saved.append(dest)
            except Exception:
                continue
    finally:
        try:
            msg.close()
        except Exception:
            pass
    return saved


def _email_attachments_to_markdown(src_path: Path, md_out_path: Path) -> int:
    """For an email (.eml/.msg), save its non-image attachments into a folder
    next to the converted Markdown file, and convert each attachment to Markdown
    too. Returns the number of attachments saved."""
    ext = src_path.suffix.lower()
    dest_dir = _unique_dir(md_out_path.parent, md_out_path.stem + "_attachments")

    if ext == ".eml":
        saved = _extract_eml_attachments(src_path, dest_dir)
    elif ext == ".msg":
        saved = _extract_msg_attachments(src_path, dest_dir)
    else:
        saved = []

    # Convert each saved attachment to Markdown alongside the original.
    for att_path in saved:
        try:
            _to_markdown(att_path, dest_dir)
        except Exception:
            pass   # keep the original attachment even if it can't be converted

    return len(saved)


# ─────────────────────────────────────────────
# Document Structuring — AI filename suggestions and organize-by-filetype.
# The AI naming reads the opening text of a document and asks the AI service
# for a professional, standardized filename; everything is best-effort and
# degrades gracefully when text or the API is unavailable.
# ─────────────────────────────────────────────

# AI naming backend. Credentials are NEVER hardcoded or baked into this build —
# the endpoint, key, and deployment are read at runtime from environment
# variables so no secret ever exists as plain text in the source or in the
# shipped .exe. This keeps the build safe to distribute and run on a local
# laptop: a malicious reader of the source/binary finds no key to exploit. The
# secret lives only in the OS/user environment (or a secrets manager / launcher
# script) on machines that are meant to use the feature.
#
# Required environment variables (set them outside this build):
#   APFP_AI_ENDPOINT     e.g. https://my-resource.openai.azure.com
#   APFP_AI_API_KEY      the Azure OpenAI key
#   APFP_AI_DEPLOYMENT   the deployment/model name
#   APFP_AI_API_VERSION  optional; defaults to 2024-06-01
#
# When any of the three required variables is unset, _ai_backend_ready() is
# False, both AI checkboxes render disabled with "(coming soon)", and the
# feature cannot run.
_AI_ENV_PREFIX = "APFP_AI_"


def _load_ai_backend() -> dict:
    """Read the AI backend config from the environment at call time. Returns a
    dict with the keys the request helpers expect; required values are '' when
    unset. No secret is ever stored in source or this module — it only reads
    os.environ on demand."""
    return {
        "endpoint":    os.environ.get(_AI_ENV_PREFIX + "ENDPOINT", "").strip().rstrip("/"),
        "api_key":     os.environ.get(_AI_ENV_PREFIX + "API_KEY", "").strip(),
        "deployment":  os.environ.get(_AI_ENV_PREFIX + "DEPLOYMENT", "").strip(),
        "api_version": os.environ.get(_AI_ENV_PREFIX + "API_VERSION", "").strip() or "2024-06-01",
    }


def _ai_backend_ready() -> bool:
    """True only when the endpoint, key, and deployment are all present in the
    environment. The credentials live entirely outside this build."""
    cfg = _load_ai_backend()
    return all(cfg[k] for k in ("endpoint", "api_key", "deployment"))


def _extract_document_text(src_path: Path, limit: int = 1500) -> str:
    """Return up to `limit` characters of text from a document, for AI naming.
    Reuses markitdown (already bundled) so it works across PDF/Word/Excel/etc.
    Returns '' if no usable text can be extracted (the caller then falls back)."""
    try:
        md = _get_markitdown()
        result = md.convert(str(src_path.resolve()))
        text = (result.text_content or "")
        text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text).strip()
        return text[:limit]
    except Exception:
        return ""


def _sanitize_suggested_name(name: str) -> str:
    """Clean an AI-suggested filename: strip illegal characters, any extension it
    tacked on, surrounding quotes/whitespace, and cap the length. '' if nothing
    usable remains."""
    name = (name or "").strip().strip('"').strip("'").strip()
    # Drop a trailing extension the model may have appended.
    name = re.sub(r"\.[A-Za-z0-9]{1,5}$", "", name).strip()
    # Remove characters illegal in Windows filenames.
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", name).strip().strip(".")
    return name[:120].strip()


def _ai_suggest_filename(text: str, cfg: dict, timeout: int = 30):
    """Ask the AI backend for a professional, standardized base filename (no
    extension) given a document's opening text. Returns the suggestion, or None
    on any failure (network, auth, bad response) so the caller can fall back.
    Pure stdlib (urllib) — no extra dependency to bundle. The request shape
    targets an Azure-OpenAI-style chat-completions endpoint and may be adjusted
    when the real service details are provided."""
    import urllib.request
    url = (f"{cfg['endpoint'].rstrip('/')}/openai/deployments/{cfg['deployment']}"
           f"/chat/completions?api-version={cfg['api_version']}")
    system = (
        "You are a document-naming assistant for a professional services firm. "
        "Given the opening text of a document, propose ONE concise, professional, "
        "standardized file name WITHOUT a file extension. Base it on the document "
        "type, date, entity/organization name, and the period covered when those "
        "are present. Prefer the form 'Entity - Document Type - Period' using Title "
        "Case, at most ~12 words. Reply with ONLY the file name — no quotes, no "
        "extension, no explanation."
    )
    body = json.dumps({
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": "Document opening text:\n\n" + text},
        ],
        "temperature": 0.2,
        "max_tokens": 40,
    }).encode("utf-8")
    req = urllib.request.Request(url, data=body, method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("api-key", cfg["api_key"])
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            resp = json.load(r)
        suggestion = _sanitize_suggested_name(
            resp["choices"][0]["message"]["content"])
        return suggestion or None
    except Exception:
        return None


def _sanitize_date(raw: str):
    """Pull a YYYY-MM-DD date out of the model's reply. Returns the date string,
    or None if the reply has no valid date (e.g. it said 'NONE')."""
    m = re.search(r"\d{4}-\d{2}-\d{2}", raw or "")
    return m.group(0) if m else None


def _ai_infer_date(text: str, cfg: dict, timeout: int = 30):
    """Ask the AI backend for the date a document was created/issued, inferred
    from its opening text, as YYYY-MM-DD (used as a filename prefix). Returns the
    date string, or None on any failure or when no date can be determined.
    Pure stdlib (urllib); same Azure-OpenAI-style request shape as
    _ai_suggest_filename, adjustable when the real service details arrive."""
    import urllib.request
    url = (f"{cfg['endpoint'].rstrip('/')}/openai/deployments/{cfg['deployment']}"
           f"/chat/completions?api-version={cfg['api_version']}")
    system = (
        "You determine the single most relevant date of a document — the date it "
        "was created, issued, signed, or dated — from its opening text. Reply with "
        "ONLY that date in strict YYYY-MM-DD format. If only a month and year are "
        "clear, use 01 for the day; if only a year is clear, use 01-01. If no date "
        "can be determined, reply with exactly NONE. No other words."
    )
    body = json.dumps({
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": "Document opening text:\n\n" + text},
        ],
        "temperature": 0.0,
        "max_tokens": 16,
    }).encode("utf-8")
    req = urllib.request.Request(url, data=body, method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("api-key", cfg["api_key"])
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            resp = json.load(r)
        return _sanitize_date(resp["choices"][0]["message"]["content"])
    except Exception:
        return None


# Extension → friendly folder name for "Organize by filetype". Unknown
# extensions fall back to the upper-cased extension (e.g. ".odt" -> "ODT").
FILETYPE_FOLDERS = {
    ".pdf": "PDF",
    ".doc": "Word Documents", ".docx": "Word Documents", ".docm": "Word Documents",
    ".dot": "Word Documents", ".dotx": "Word Documents", ".rtf": "Word Documents",
    ".txt": "Text Files", ".md": "Text Files",
    ".xls": "Excel Spreadsheets", ".xlsx": "Excel Spreadsheets", ".xlsm": "Excel Spreadsheets",
    ".xlsb": "Excel Spreadsheets", ".csv": "Excel Spreadsheets",
    ".ppt": "PowerPoint", ".pptx": "PowerPoint", ".pps": "PowerPoint", ".ppsx": "PowerPoint",
    ".jpg": "Images", ".jpeg": "Images", ".png": "Images", ".gif": "Images",
    ".bmp": "Images", ".tiff": "Images", ".tif": "Images", ".webp": "Images",
    ".msg": "Emails", ".eml": "Emails",
    ".html": "Web Pages", ".htm": "Web Pages", ".mht": "Web Pages", ".mhtml": "Web Pages",
    ".zip": "Archives",
    ".vsd": "Visio", ".vsdx": "Visio",
}


def _filetype_folder(ext: str) -> str:
    ext = (ext or "").lower()
    return FILETYPE_FOLDERS.get(ext, ext.lstrip(".").upper() or "Other")


def convert_file(src_path: Path, out_dir: Path, outlook, word, visio, excel, powerpoint, mode: str) -> Path:
    """
    Convert a single file to PDF.
    mode: 'email', 'html', 'visio', 'excel', 'word_doc', 'powerpoint', 'image',
          'text_report'
    Returns the path of the created PDF. Never overwrites an existing PDF.
    """
    ext = src_path.suffix.lower()
    out_path = _unique_pdf_path(out_dir, src_path.stem)

    if mode == "email":
        if ext == ".msg":
            _msg_to_pdf(src_path, out_path, outlook, word)
        elif ext == ".eml":
            _eml_to_pdf(src_path, out_path, word)
        else:
            raise ValueError(f"Unsupported file type for Email mode: {ext}")
    elif mode == "html":
        _html_to_pdf(src_path, out_path, word)
    elif mode == "visio":
        _visio_to_pdf(src_path, out_path, visio)
    elif mode == "excel":
        _excel_to_pdf(src_path, out_path, excel)
    elif mode == "word_doc":
        _word_to_pdf(src_path, out_path, word)
    elif mode == "text_report":
        _text_report_to_pdf(src_path, out_path, word)
    elif mode == "powerpoint":
        _ppt_to_pdf(src_path, out_path, powerpoint)
    elif mode == "image":
        _image_to_pdf(src_path, out_path, word)
    else:
        raise ValueError(f"Unknown mode: {mode}")

    return out_path


def _html_to_pdf(src_path: Path, out_path: Path, word):
    """
    Convert a .html / .htm file to PDF via Word COM.
    Word opens the HTML file and exports it as PDF — no Outlook needed.
    """
    if word is None:
        raise RuntimeError(
            "Microsoft Word is required for HTML conversion.\n"
            "Please ensure Word is installed."
        )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc = None
    try:
        doc = word.Documents.Open(
            str(src_path.resolve()),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        _fit_word_doc_to_page(doc, word)
        doc.ExportAsFixedFormat(
            OutputFileName=str(out_path.resolve()),
            ExportFormat=17,        # wdExportFormatPDF
            OpenAfterExport=False,
            OptimizeFor=0,
            Range=0,
            Item=0,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=0,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )
    except Exception as e:
        raise RuntimeError(f"HTML to PDF conversion failed: {e}")
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass


# ─────────────────────────────────────────────
# Unzip helpers
# ─────────────────────────────────────────────

def _find_extraction_root(path: Path) -> Path:
    """Drill down through single-folder wrappers to find the real content root.

    If the extracted directory contains exactly one subdirectory and no files,
    descend into that subdirectory and repeat. This strips the common pattern
    where a zip wraps everything inside a single named folder.
    """
    current = path
    while True:
        children = list(current.iterdir())
        subdirs = [c for c in children if c.is_dir()]
        files   = [c for c in children if c.is_file()]
        if len(subdirs) == 1 and len(files) == 0:
            current = subdirs[0]
        else:
            return current


def _recursive_unzip_in_place(directory: Path):
    """Find and extract every nested zip anywhere inside directory, repeatedly.

    Each zip is extracted into a folder named after the zip file, then the zip
    is deleted. The loop repeats until no more zips exist in the tree —
    handling arbitrary nesting depth (zip inside zip inside zip …).
    Zips that fail to extract (corrupted, encrypted) are tracked so the loop
    cannot spin forever on an unextractable file.
    """
    failed_zips: set = set()
    while True:
        zips = [p for p in directory.rglob("*.zip") if p not in failed_zips]
        if not zips:
            break
        for zip_path in zips:
            if not zip_path.exists() or zip_path in failed_zips:
                continue
            tmp = Path(tempfile.mkdtemp())
            try:
                with zipfile.ZipFile(zip_path, "r") as zf:
                    zf.extractall(tmp)
                content_root = _find_extraction_root(tmp)
                target_dir = zip_path.parent / zip_path.stem
                target_dir.mkdir(parents=True, exist_ok=True)
                for item in content_root.iterdir():
                    target = target_dir / item.name
                    if item.is_dir():
                        if target.exists():
                            shutil.copytree(str(item), str(target), dirs_exist_ok=True)
                        else:
                            shutil.copytree(str(item), str(target))
                    else:
                        shutil.copy2(str(item), str(target))
                zip_path.unlink()
            except Exception:
                failed_zips.add(zip_path)
            finally:
                shutil.rmtree(tmp, ignore_errors=True)


def _collect_and_flatten(src_dir: Path, dest_dir: Path):
    """Copy every file found anywhere inside src_dir into dest_dir as a flat list.

    Folder structure is not preserved. Name collisions get a numeric suffix:
    report.docx, report_2.docx, report_3.docx, etc. Pre-existing files in
    dest_dir are never overwritten — they get the suffixed name instead.

    Returns the list of file paths actually created in dest_dir, so callers can
    act on exactly the extracted files without touching unrelated siblings.
    """
    created = []
    for item in src_dir.rglob("*"):
        if item.is_file():
            target = dest_dir / item.name
            if target.exists():
                stem, suffix = item.stem, item.suffix
                counter = 2
                while target.exists():
                    target = dest_dir / f"{stem}_{counter}{suffix}"
                    counter += 1
            shutil.copy2(str(item), str(target))
            created.append(target)
    return created


def _image_to_pdf(src: Path, out_path: Path, word):
    """Embed an image in a blank Word document and export as PDF."""
    doc = word.Documents.Add()
    try:
        doc.InlineShapes.AddPicture(
            FileName=str(src.resolve()),
            LinkToFile=False,
            SaveWithDocument=True,
        )
        doc.SaveAs2(str(out_path.resolve()), FileFormat=17)  # wdFormatPDF
    finally:
        try:
            doc.Close(SaveChanges=False)
        except Exception:
            pass


# File-extension → conversion mode for the auto-PDF feature.
# Word handles txt/rtf/csv/xml natively; images are embedded via _image_to_pdf.
AUTO_PDF_EXTS = {
    ".eml": "email",      ".msg": "email",
    ".html": "html",      ".htm": "html",    ".mht": "html",   ".mhtml": "html",
    ".vsd": "visio",      ".vsdx": "visio",
    ".xls": "excel",      ".xlsx": "excel",   ".xlsm": "excel",  ".xlsb": "excel",
    ".csv": "excel",
    ".doc": "word_doc",   ".docx": "word_doc",
    ".txt": "word_doc",   ".rtf": "word_doc",  ".xml": "word_doc",
    ".ppt": "powerpoint", ".pptx": "powerpoint",
    ".pps": "powerpoint", ".ppsx": "powerpoint",
    ".rep": "text_report", ".rpt": "text_report",
    ".prn": "text_report", ".lst": "text_report",
    ".jpg": "image",      ".jpeg": "image",   ".png": "image",
    ".bmp": "image",      ".gif": "image",    ".tiff": "image",  ".tif": "image",
}


def _auto_pdf_scan_and_convert(candidate_files, status_cb, cancel_event=None):
    """Convert each file in candidate_files to PDF.

    candidate_files is the explicit list of files that were extracted from the
    zip(s). Only these are ever touched — unrelated files that happen to share
    the output folder are never scanned, converted, or deleted.

    Files with no extension are treated as plain text and converted via Word.
    Existing .pdf files are always skipped and never deleted.
    The original is deleted ONLY if the output PDF exists and has non-zero size.
    Returns a list of filenames that could not be converted.

    Office applications are launched lazily as DEDICATED, isolated instances and
    only those are quit at the end — the user's own open Office windows are
    never hidden, driven, or closed.
    """
    if not HAS_COM:
        status_cb("Auto-PDF skipped: pywin32 is not installed.")
        return []

    pdf_failed = []
    pythoncom.CoInitialize()
    apps = {}          # family -> COM app (None when that Office app is missing)
    app_pids = set()   # PIDs of ONLY the instances we launched
    outlook = None

    def _app(family):
        """Lazily launch a dedicated instance for `family`, tracking its PID.

        Must go through _launch_office_isolated: a plain Dispatch() attaches to
        the user's already-open Office window, which we would then hide and —
        with alerts suppressed — Quit at the end, discarding their unsaved work.
        """
        if family not in apps:
            try:
                app_obj, pids = _launch_office_isolated(family)
            except Exception:
                app_obj, pids = None, set()
            apps[family] = app_obj
            app_pids.update(pids)
        return apps[family]

    try:
        # Hard safeguard: auto-PDF may only ever delete files that were actually
        # extracted from the upload (the explicit candidate list) — never any
        # other file that happens to share the output folder.
        originals = _resolved_set(candidate_files)
        files = [
            f for f in candidate_files
            if f.is_file()
            and f.suffix.lower() != ".pdf"          # never touch existing PDFs
            and (f.suffix == "" or f.suffix.lower() in AUTO_PDF_EXTS)
        ]
        if not files:
            status_cb("Auto-PDF: no convertible files found.")
        else:
            total = len(files)
            converted = 0

            for i, src in enumerate(files):
                if cancel_event is not None and cancel_event.is_set():
                    break
                if not src.exists():
                    continue
                ext = src.suffix.lower()

                extensionless = (ext == "")
                if extensionless:
                    # No extension — treat as a plain-text report. The converter
                    # makes its own temp .txt copy so Word opens it without a
                    # format-detection dialog, and rejects binaries cleanly.
                    mode = "text_report"
                    status_cb(f"Auto-PDF ({i + 1}/{total}): {src.name}  (no extension, treating as text)")
                else:
                    mode = AUTO_PDF_EXTS[ext]
                    status_cb(f"Auto-PDF ({i + 1}/{total}): {src.name}")
                out_path = None

                # Bound each file so one stuck conversion can't hang the batch.
                wd = _FileWatchdog(app_pids)
                try:
                    with wd:
                        if extensionless:
                            # Converted directly rather than through convert_file
                            # so the PDF keeps the full name — these files have
                            # no stem to name it after.
                            out_path = _unique_pdf_path(src.parent, src.name)
                            _text_report_to_pdf(src, out_path, _app("word"))
                        elif mode == "image":
                            out_path = _unique_pdf_path(src.parent, src.stem)
                            _image_to_pdf(src, out_path, _app("word"))
                        else:
                            if mode == "email" and outlook is None:
                                outlook = _ensure_outlook()
                            out_path = convert_file(
                                src, src.parent, outlook,
                                _app("word") if mode in ("word_doc", "html",
                                                         "text_report", "email") else None,
                                _app("visio") if mode == "visio" else None,
                                _app("excel") if mode == "excel" else None,
                                _app("powerpoint") if mode == "powerpoint" else None,
                                mode,
                            )

                    # Only delete the original if the PDF was actually written
                    if out_path is not None and out_path.exists() and out_path.stat().st_size > 0:
                        _delete_original_file(src, originals)
                        converted += 1
                    else:
                        pdf_failed.append(src.name)
                except Exception:
                    pdf_failed.append(src.name)

                # The watchdog killed our instances — drop them so the next file
                # lazily launches fresh, healthy ones.
                if wd.fired:
                    apps.clear()
                    app_pids.clear()

            status_cb(f"Auto-PDF done: {converted}/{total} file(s) converted.")

    finally:
        # Quit only the instances we launched, then make sure none linger (e.g.
        # after a watchdog kill). Outlook is the user's mail client: never quit.
        for app in apps.values():
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
        for pid in app_pids:
            _terminate_pid(pid)
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    return pdf_failed


def _unzip_worker(zip_paths, output_dir, separate_folders, auto_pdf,
                  status_cb, progress_cb, finish_cb, cancel_event=None,
                  item_cb=None):
    """Thread worker: extract zips, handle flat vs structured output, then
    optionally auto-convert all non-PDF files to PDF.

    cancel_event: optional threading.Event; when set, stop before the next item.
    """
    def cancelled():
        return cancel_event is not None and cancel_event.is_set()

    success, failed = 0, []
    protected = []         # password-protected archives we skipped
    extracted_files = []   # explicit list of files extracted from the zip(s)
    total = len(zip_paths)

    for i, zip_path in enumerate(zip_paths):
        if cancelled():
            break
        # Skip password-protected archives — they can't be extracted here.
        if _is_password_protected(zip_path):
            protected.append(zip_path.name)
            progress_cb()
            continue
        status_cb(f"Extracting: {zip_path.name}  ({i + 1}/{total})")
        if item_cb is not None:
            item_cb(zip_path.name)
        tmp = Path(tempfile.mkdtemp())
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                zf.extractall(tmp)

            if separate_folders:
                # Preserve structure; expand nested zips in the destination folder.
                content_root = _find_extraction_root(tmp)
                dest = output_dir / zip_path.stem
                dest.mkdir(parents=True, exist_ok=True)
                for item in content_root.iterdir():
                    target = dest / item.name
                    if item.is_dir():
                        if target.exists():
                            shutil.copytree(str(item), str(target), dirs_exist_ok=True)
                        else:
                            shutil.copytree(str(item), str(target))
                    else:
                        shutil.copy2(str(item), str(target))
                status_cb(f"Expanding nested zips in {zip_path.stem}…")
                _recursive_unzip_in_place(dest)
                # Only files under this zip's own destination folder are eligible
                # for auto-PDF — never siblings already in the output folder.
                extracted_files.extend(p for p in dest.rglob("*") if p.is_file())
            else:
                # Flat mode: expand ALL nested zips inside the temp dir first,
                # then copy every file at any depth into a single master folder.
                status_cb(f"Expanding all nested zips in {zip_path.stem}…")
                _recursive_unzip_in_place(tmp)
                dest = output_dir
                dest.mkdir(parents=True, exist_ok=True)
                status_cb(f"Flattening files from {zip_path.stem} into master folder…")
                # _collect_and_flatten returns exactly the files it created, so
                # pre-existing siblings in the output folder are never touched.
                extracted_files.extend(_collect_and_flatten(tmp, dest))

            success += 1
        except zipfile.BadZipFile:
            failed.append((zip_path.name, "Not a valid ZIP file or the file is corrupted."))
        except RuntimeError as e:
            failed.append((zip_path.name, f"Cannot extract (possibly encrypted): {e}"))
        except Exception as e:
            failed.append((zip_path.name, str(e)))
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
            progress_cb()

    pdf_failed = []
    if auto_pdf and success > 0 and not cancelled():
        status_cb("Auto-PDF: converting extracted files…")
        pdf_failed = _auto_pdf_scan_and_convert(extracted_files, status_cb, cancel_event)

    finish_cb(success, failed, pdf_failed, protected)


# ─────────────────────────────────────────────
# Self-update from GitHub Releases
# ─────────────────────────────────────────────
#
# The app updates itself in place: it downloads the newest BoxOfScraps.exe from
# the repo's latest release and swaps it over the running one. Two details make
# that safe and painless, and both are load-bearing:
#
#  * NO "Unblock" STEP. The Properties > Unblock checkbox exists because the
#    *downloading application* tags a file with the Mark-of-the-Web alternate
#    data stream (:Zone.Identifier, ZoneId=3). Browsers and mail clients do this
#    deliberately; urllib does not. Because we fetch the exe ourselves the file
#    arrives untagged, so it runs immediately and SmartScreen's shell prompt —
#    which fires *on* the MOTW tag — never appears. `_strip_mark_of_the_web`
#    removes the tag explicitly anyway so the guarantee doesn't rest on that.
#
#  * NO HELPER BATCH SCRIPT. Windows locks a running image against deletion and
#    overwriting, but permits it to be RENAMED. So the swap is two renames done
#    in-process (see `_swap_in_new_exe`) rather than a .bat that has to outlive
#    the process, poll for it to exit, and then move files. That avoids a
#    flashing console window and a second artifact for antivirus to distrust.

GITHUB_OWNER      = "lbangrazi43"
GITHUB_REPO       = "Allpowerfulfileprep"
UPDATE_ASSET_NAME = "BoxOfScraps.exe"
GITHUB_API_LATEST = (f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
                     f"/releases/latest")
RELEASES_PAGE_URL = f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"

UPDATE_USER_AGENT     = f"BoxOfScraps/{APP_VERSION}"   # GitHub rejects UA-less calls
UPDATE_NET_TIMEOUT    = 30                  # seconds, version check
UPDATE_DL_TIMEOUT     = 120                 # seconds, per read on the download
UPDATE_MIN_EXE_BYTES  = 5 * 1024 * 1024     # sanity floor for a real build
UPDATE_BACKUP_SUFFIX  = ".old"
UPDATE_CHUNK          = 256 * 1024
# Minimum gap between version checks. GitHub allows 60 unauthenticated API calls
# an hour per *IP address* — which everyone behind one office connection shares —
# so a user leaning on the button could burn the quota for their whole building.
UPDATE_CHECK_COOLDOWN = 30                  # seconds


class _UpdateCancelled(Exception):
    """Raised inside the download loop when the user presses Cancel."""


def _parse_version(tag: str) -> tuple:
    """Turn a release tag or version string into a comparable tuple.

    'v1.7.0' -> (1, 7, 0, 1);  '1.7.0-beta.2' -> (1, 7, 0, 0).
    The trailing element ranks a final release above any of its own prereleases,
    so someone running 1.7.0 is never offered 1.7.0-beta.2 as an 'update'."""
    s = (tag or "").strip().lstrip("vV")
    core, _, pre = s.partition("-")
    parts = []
    for chunk in core.split(".")[:3]:
        m = re.match(r"\d+", chunk)
        parts.append(int(m.group(0)) if m else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts) + (0 if pre else 1,)


def _rate_limit_reset_text(err) -> str:
    """Local clock time at which GitHub's hourly quota refills, read from the
    X-RateLimit-Reset header (a Unix timestamp). Returns "" when the header is
    missing or unparseable, so callers can drop the clause entirely rather than
    print a wrong time. Formatted by hand because strftime's 12-hour,
    no-leading-zero flag differs between platforms."""
    try:
        epoch = int((err.headers.get("X-RateLimit-Reset") or "").strip())
        t = time.localtime(epoch)
    except (AttributeError, TypeError, ValueError, OSError):
        return ""
    hour = t.tm_hour % 12 or 12
    return f"{hour}:{t.tm_min:02d} {'AM' if t.tm_hour < 12 else 'PM'}"


def _fetch_latest_release(timeout: int = UPDATE_NET_TIMEOUT) -> dict:
    """Fetch the newest published, non-prerelease GitHub release.

    /releases/latest skips drafts and prereleases by design, which lines up with
    the release runbook's invariant that only one stable release is public at a
    time — so beta tags are never offered to users and the drafted previous
    release stops being advertised the moment a new one ships.

    Raises RuntimeError carrying a message fit to show a user (offline, blocked
    by a proxy, rate-limited) rather than leaking a urllib traceback."""
    import urllib.request
    import urllib.error
    req = urllib.request.Request(GITHUB_API_LATEST)
    req.add_header("User-Agent", UPDATE_USER_AGENT)
    req.add_header("Accept", "application/vnd.github+json")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        # A 403 is only a quota problem when the remaining-calls header says 0;
        # 429 is always one. Anything else 403 is a genuine refusal.
        if e.code in (403, 429):
            remaining = (e.headers.get("X-RateLimit-Remaining") or "").strip()
            if e.code == 429 or remaining == "0":
                when = _rate_limit_reset_text(e)
                raise RuntimeError(
                    "GitHub's update-check limit has been reached. It allows 60 "
                    "checks an hour per internet connection, shared by everyone "
                    "on your network — so this can happen even if you've only "
                    "checked once."
                    + (f"\n\nTry again after {when}." if when
                       else "\n\nThe limit resets within the hour."))
            raise RuntimeError(f"GitHub refused the request (HTTP {e.code}).")
        if e.code == 404:
            raise RuntimeError(
                "No published release was found for this app on GitHub.")
        raise RuntimeError(f"GitHub returned HTTP {e.code} ({e.reason}).")
    except Exception as e:
        raise RuntimeError(
            "Couldn't reach GitHub. Check your internet connection — a company "
            f"firewall or proxy may also be blocking it.\n\nDetail: {e}")


def _find_release_asset(release: dict, name: str):
    """Find a named file attached to a release, or None."""
    for asset in release.get("assets") or []:
        if (asset.get("name") or "").lower() == name.lower():
            return asset
    return None


def _current_exe_path():
    """Path to the running one-file .exe, or None when running from source.

    The single most important safeguard in this module: from a source checkout
    `sys.executable` is python.exe, and every path below would happily rename
    and overwrite *that*. Everything that touches disk refuses when this is
    None."""
    if not getattr(sys, "frozen", False):
        return None
    try:
        return Path(sys.executable).resolve()
    except OSError:
        return Path(sys.executable)


def _strip_mark_of_the_web(path: Path) -> None:
    """Remove the :Zone.Identifier alternate data stream ('this came from the
    internet') if one is present, so the new exe never needs the manual
    Properties > Unblock step before Windows will run it.

    Best-effort by design: alternate data streams only exist on NTFS, so this is
    a silent no-op on FAT32/exFAT/network shares — which cannot carry a MOTW tag
    in the first place."""
    try:
        os.remove(f"{path}:Zone.Identifier")
    except OSError:
        pass


def _looks_like_windows_exe(path: Path) -> bool:
    """True if `path` is a plausible Windows executable: an 'MZ' DOS header whose
    e_lfanew offset lands on a 'PE\\0\\0' signature.

    Catches a truncated download, and — more usefully — a captive-portal or
    proxy error page served with a 200 status, which would otherwise be swapped
    in as the application."""
    try:
        with open(path, "rb") as f:
            if f.read(2) != b"MZ":
                return False
            f.seek(0x3C)
            raw = f.read(4)
            if len(raw) != 4:
                return False
            f.seek(int.from_bytes(raw, "little"))
            return f.read(4) == b"PE\0\0"
    except OSError:
        return False


def _sha256_file(path: Path) -> str:
    """Streaming SHA-256 of a file (the exe is ~120 MB — never read it whole)."""
    import hashlib
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


# Folder names that mean "this directory is continuously synced by a cloud
# client". Matched against whole path components, so "OneDrive - Contoso"
# hits via its prefix while an ordinary folder called "Boxes" does not.
_CLOUD_SYNC_FOLDERS = {
    "onedrive":       "OneDrive",
    "dropbox":        "Dropbox",
    "google drive":   "Google Drive",
    "googledrive":    "Google Drive",
    "box":            "Box",
    "icloud drive":   "iCloud Drive",
    "icloudrive":     "iCloud Drive",
}


def _cloud_sync_provider(path: Path):
    """Name the cloud-sync service whose folder `path` sits in, or None.

    Consults OneDrive's own environment variables first (authoritative, and it
    catches a sync root the user has renamed), then falls back to matching
    well-known folder names. Drives the pre-update warning and justifies the
    retry loop in `_replace_with_retry`: a sync client can hold a short-lived
    lock on a file it is uploading, which would otherwise fail the swap."""
    try:
        resolved = str(path.resolve())
    except OSError:
        resolved = str(path)
    low = resolved.lower()
    for var in ("OneDrive", "OneDriveCommercial", "OneDriveConsumer"):
        root = os.environ.get(var, "").strip().rstrip("\\/")
        if root and (low == root.lower() or low.startswith(root.lower() + os.sep)):
            return "OneDrive"
    for part in Path(resolved).parts:
        # "OneDrive - Contoso" and "Google Drive" both reduce to their brand.
        base = part.split(" - ")[0].strip().lower()
        if base in _CLOUD_SYNC_FOLDERS:
            return _CLOUD_SYNC_FOLDERS[base]
    return None


def _dir_is_writable(directory: Path) -> bool:
    """True if we can create and delete a file in `directory`.

    Checked *before* a ~120 MB download so an install sitting under Program
    Files fails immediately with useful advice, instead of after a long download
    and only at the moment of the swap."""
    try:
        fd, probe = tempfile.mkstemp(prefix=".bos-wtest-", dir=str(directory))
        os.close(fd)
        os.remove(probe)
        return True
    except OSError:
        return False


def _has_room_for_update(directory: Path, need_bytes: int) -> bool:
    """True if `directory`'s volume can hold the download plus the backup copy
    the swap keeps until the next launch."""
    try:
        return shutil.disk_usage(str(directory)).free >= need_bytes
    except OSError:
        return True     # can't measure — let the write itself be the judge


def _replace_with_retry(src, dst, attempts: int = 8, delay: float = 0.25) -> None:
    """os.replace with a short backoff on PermissionError.

    A cloud-sync client or an antivirus scanner routinely holds a sub-second
    lock on a file it has just watched change; without a retry that surfaces as
    a spurious 'access denied' mid-swap. Roughly 9 seconds of total patience,
    then the original error is re-raised so the caller can roll back."""
    last = None
    for i in range(attempts):
        try:
            os.replace(src, dst)
            return
        except PermissionError as e:
            last = e
            time.sleep(delay * (i + 1))
    raise last


def _stage_beside(src: Path, target: Path) -> Path:
    """Put the downloaded exe on the same volume as `target`, ready to be
    renamed into place, and return its path.

    os.replace cannot rename across volumes, and the download lives in the
    system temp folder (which may be on a different drive) — so this is a cheap
    rename when they match and a copy when they don't."""
    staged = target.with_name("." + target.stem + ".new.tmp")
    try:
        os.remove(staged)
    except OSError:
        pass
    try:
        os.replace(src, staged)
    except OSError:
        shutil.copy2(src, staged)      # different volume: copy instead
    return staged


def _swap_in_new_exe(staged: Path, target: Path) -> Path:
    """Replace the running .exe with `staged`; return the backup path.

    Windows will not let a running image be deleted or overwritten, but it does
    allow it to be RENAMED — the open handle follows the file, only its
    directory entry moves. So the swap is two renames and needs no helper
    process:

        target -> target.old     (move the running exe out of the way)
        staged -> target         (new build takes its place)

    If the second rename fails the first is undone, so a failed update always
    leaves a working application behind. `staged` must already be on `target`'s
    volume (see `_stage_beside`)."""
    backup = target.with_name(target.name + UPDATE_BACKUP_SUFFIX)
    try:
        os.remove(backup)          # clear a leftover we couldn't delete earlier
    except OSError:
        pass
    _replace_with_retry(target, backup)
    try:
        _replace_with_retry(staged, target)
    except Exception:
        try:                       # roll back: put the working exe back
            _replace_with_retry(backup, target)
        except Exception:
            pass
        raise
    return backup


def _cleanup_previous_update(exe: Path) -> None:
    """Delete what a previous update left behind: the backup of the old exe and
    any abandoned staging file.

    Runs once at startup because the backup cannot be removed during the update
    that created it — at that moment it is still the running image. Entirely
    best-effort: a leftover that is briefly locked (a sync client is uploading
    it) is harmless and gets cleaned up on a later launch."""
    for leftover in (exe.with_name(exe.name + UPDATE_BACKUP_SUFFIX),
                     exe.with_name("." + exe.stem + ".new.tmp")):
        try:
            if leftover.exists():
                os.remove(leftover)
        except OSError:
            pass


def _download_release_asset(url: str, dest: Path, expected_size: int = 0,
                            progress_cb=None, cancel_event=None,
                            timeout: int = UPDATE_DL_TIMEOUT) -> int:
    """Stream a release asset to `dest`, reporting progress and honouring Cancel.

    Fetching this ourselves rather than handing the URL to a browser is what
    keeps the downloaded exe free of the Mark-of-the-Web tag (see the module
    header). Returns the byte count actually written."""
    import urllib.request
    req = urllib.request.Request(url)
    req.add_header("User-Agent", UPDATE_USER_AGENT)
    req.add_header("Accept", "application/octet-stream")
    got = 0
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            total = expected_size or int(r.headers.get("Content-Length") or 0)
            with open(dest, "wb") as f:
                while True:
                    if cancel_event is not None and cancel_event.is_set():
                        raise _UpdateCancelled()
                    chunk = r.read(UPDATE_CHUNK)
                    if not chunk:
                        break
                    f.write(chunk)
                    got += len(chunk)
                    if progress_cb is not None:
                        progress_cb(got, total)
    except _UpdateCancelled:
        raise
    except Exception as e:
        raise RuntimeError(f"Download failed: {e}")
    if expected_size and got != expected_size:
        raise RuntimeError(
            f"The download is incomplete — {got:,} of {expected_size:,} bytes "
            "arrived. The connection dropped part-way; please try again.")
    return got


def _relaunch(exe: Path) -> None:
    """Start the freshly-installed exe, detached so it outlives this process."""
    import subprocess
    DETACHED_PROCESS = 0x00000008
    CREATE_NEW_PROCESS_GROUP = 0x00000200
    subprocess.Popen(
        [str(exe)], cwd=str(exe.parent), close_fds=True,
        creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP)


def _relaunch_as_admin(exe: Path) -> bool:
    """Restart the app behind a UAC prompt, for installs in a location only an
    administrator can write (Program Files), where the rename swap would fail
    with access denied. Returns True if the elevation prompt was accepted."""
    try:
        import ctypes
        rc = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", str(exe), None, str(exe.parent), 1)
        return int(rc) > 32        # ShellExecuteW's success threshold
    except Exception:
        return False


def _open_releases_page() -> None:
    """Fall back to the browser when an in-app update can't proceed."""
    try:
        import webbrowser
        webbrowser.open(RELEASES_PAGE_URL)
    except Exception:
        pass


# ─────────────────────────────────────────────
# GUI
# ─────────────────────────────────────────────

APP_BG      = "#f0f4f8"
ACCENT      = "#0078d4"
BTN_FG      = "#ffffff"
DROP_BG     = "#e8f0fe"
DROP_BD     = "#aac4e8"
CANCEL_BG          = "#c0392b"   # red: active Cancel button
CANCEL_BG_ACTIVE   = "#a93226"   # red: pressed
CANCEL_OFF_BG      = "#cccccc"   # grey: disabled Cancel button
CANCEL_OFF_FG      = "#888888"
SIDEBAR_BG  = "#1e2d3d"
SIDEBAR_FG  = "#c9d8e8"
SIDEBAR_SEL = "#0078d4"
SIDEBAR_W   = 190
# The informational footer entry (About & Updates) is deliberately quieter than
# the tool list above it: dimmer text and a muted slate selection instead of the
# accent blue, so it reads as "about the app" rather than a seventh feature.
SIDEBAR_INFO_FG   = "#7e93ab"
SIDEBAR_INFO_SEL  = "#33475e"
SIDEBAR_RULE      = "#3a5068"

# Toggle-switch palette
TOGGLE_ON       = "#0078d4"    # filled track when on (accent)
TOGGLE_OFF      = "#bcc5cf"    # grey track when off
TOGGLE_DISABLED = "#d8dde2"    # washed-out track when disabled


def _hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _lerp_rgb(c1, c2, t):
    """Blend two '#rrggbb' colors; t=0 -> c1, t=1 -> c2. Returns an (r,g,b) tuple."""
    a, b = _hex_to_rgb(c1), _hex_to_rgb(c2)
    return tuple(int(round(a[k] + (b[k] - a[k]) * t)) for k in range(3))


def _shade(hex_color, t):
    """Darken `hex_color` toward black by fraction t (0..1); returns '#rrggbb'.
    Used to derive a hover/pressed shade from a button's OWN color so a click
    darkens it in place, keeping its format, instead of Tk's grey 'active' look."""
    r, g, b = _lerp_rgb(hex_color, "#000000", t)
    return f"#{r:02x}{g:02x}{b:02x}"


def _style_press_button(btn):
    """Make a flat tk.Button visibly press in on click: darken its own background
    a little on hover and more while held, then restore on release. Keeps the
    button's color/format (no flip to Tk's default grey/white active state).
    Guards on the disabled state so inactive buttons don't react."""
    try:
        base = btn.cget("bg")
        fg = btn.cget("fg")
    except Exception:
        return
    hover = _shade(base, 0.06)
    pressed = _shade(base, 0.18)
    btn.config(activebackground=hover, activeforeground=fg)

    def _press(_e=None):
        if str(btn["state"]) != "disabled":
            btn.config(activebackground=pressed)

    def _release(_e=None):
        if str(btn["state"]) != "disabled":
            btn.config(activebackground=hover)

    btn.bind("<ButtonPress-1>", _press, add="+")
    btn.bind("<ButtonRelease-1>", _release, add="+")


class _ToggleSwitch(tk.Frame):
    """A modern pill-style on/off switch bound to a BooleanVar, with a text label
    to its right — a rounded, drop-in replacement for tk.Checkbutton that matches
    the app theme. The pill is rendered anti-aliased (supersampled with Pillow) so
    it looks smooth rather than blocky, and the knob slides with a short animation
    when toggled. Supports `command`, a "disabled" state, and live updates when the
    variable is set elsewhere. Pack/grid it like any widget."""

    W, H, PAD = 32, 18, 2     # display size; knob diameter = H - 2*PAD
    SS = 4                    # supersampling factor for anti-aliasing
    FRAMES, FRAME_MS = 8, 12  # knob-slide animation length / cadence

    def __init__(self, parent, text, variable, command=None, state="normal",
                 bg=APP_BG, fg="#333", font=("Segoe UI", 9)):
        super().__init__(parent, bg=bg)
        self._var = variable
        self._command = command
        self._state = state
        self._enabled_fg = fg
        self._bg = bg
        self._pos = 1.0 if variable.get() else 0.0   # current knob position 0..1
        self._anim_id = None
        self._photo = None                            # keep a ref so it isn't GC'd
        self._canvas = tk.Canvas(self, width=self.W, height=self.H, bg=bg,
                                 highlightthickness=0, bd=0)
        self._img_id = self._canvas.create_image(0, 0, anchor="nw")
        self._canvas.pack(side="left", pady=1)
        self._label = tk.Label(self, text=text, bg=bg, fg=fg, font=font,
                               anchor="w", justify="left")
        self._label.pack(side="left", padx=(8, 0))
        for w in (self._canvas, self._label):
            w.bind("<Button-1>", self._on_click)
        self._apply_cursor()
        self._update_label_fg()
        self._var.trace_add("write", self._on_var_change)
        self._render(self._pos)

    def _apply_cursor(self):
        cur = "hand2" if self._state == "normal" else "arrow"
        self._canvas.config(cursor=cur)
        self._label.config(cursor=cur)

    def _update_label_fg(self):
        self._label.config(
            fg="#9aa3ad" if self._state == "disabled" else self._enabled_fg)

    def _render(self, pos):
        """Draw the pill + knob at knob position `pos` (0..1), anti-aliased."""
        from PIL import Image, ImageDraw, ImageTk
        ss = self.SS
        W, H, pad = self.W * ss, self.H * ss, self.PAD * ss
        bg = _hex_to_rgb(self._bg)
        if self._state == "disabled":
            track = _hex_to_rgb(TOGGLE_DISABLED)
        else:
            track = _lerp_rgb(TOGGLE_OFF, TOGGLE_ON, pos)
        img = Image.new("RGB", (W, H), bg)
        d = ImageDraw.Draw(img)
        d.rounded_rectangle([0, 0, W - 1, H - 1], radius=H // 2, fill=track)
        knob = H - 2 * pad
        travel = W - 2 * pad - knob
        x0 = pad + travel * pos
        d.ellipse([x0, pad, x0 + knob, pad + knob], fill=(255, 255, 255))
        img = img.resize((self.W, self.H), Image.LANCZOS)
        self._photo = ImageTk.PhotoImage(img)
        self._canvas.itemconfig(self._img_id, image=self._photo)

    def _animate_to(self, target):
        """Slide the knob from its current position to `target` (0 or 1)."""
        if self._anim_id is not None:
            try:
                self.after_cancel(self._anim_id)
            except Exception:
                pass
            self._anim_id = None
        start = self._pos
        delta = target - start
        if abs(delta) < 1e-6:
            return

        def _step(i):
            if i >= self.FRAMES:
                self._pos = target
                self._render(target)
                self._anim_id = None
                return
            f = (i + 1) / self.FRAMES
            f = 1 - (1 - f) * (1 - f)            # ease-out
            self._pos = start + delta * f
            self._render(self._pos)
            self._anim_id = self.after(self.FRAME_MS, lambda: _step(i + 1))

        _step(0)

    def _on_var_change(self, *_a):
        target = 1.0 if self._var.get() else 0.0
        self._update_label_fg()
        if self._state == "disabled":
            self._pos = target                  # snap; no animation when disabled
            self._render(target)
        else:
            self._animate_to(target)

    def _on_click(self, event=None):
        if self._state == "disabled":
            return
        self._var.set(not self._var.get())      # trace drives the animation
        if self._command is not None:
            self._command()

    def config(self, **kwargs):
        if "state" in kwargs:
            self._state = kwargs.pop("state")
            self._apply_cursor()
            self._update_label_fg()
            self._render(self._pos)
        if "text" in kwargs:
            self._label.config(text=kwargs.pop("text"))
        if kwargs:
            super().config(**kwargs)

    configure = config


class _ModernRadio(tk.Frame):
    """A Windows 11-style radio option: an anti-aliased ring with a filled accent
    dot when selected, plus a label. Several instances share one variable; clicking
    selects this instance's value. Bigger and smoother than tk.Radiobutton. The
    accent dot grows in (and the deselected one shrinks out) with a short ease-out
    animation when the selection changes, like a modern web radio."""

    D, SS = 18, 4             # circle diameter (display px) / supersample factor
    FRAMES, FRAME_MS = 8, 12  # dot fill-in animation length / cadence

    def __init__(self, parent, text, variable, value, command=None,
                 bg=APP_BG, fg="#333", font=("Segoe UI", 9)):
        super().__init__(parent, bg=bg)
        self._var = variable
        self._value = value
        self._command = command
        self._bg = bg
        self._photo = None
        self._anim_id = None
        # Current fill amount of the accent dot: 0 = empty ring, 1 = full dot.
        self._fill = 1.0 if variable.get() == value else 0.0
        self._canvas = tk.Canvas(self, width=self.D, height=self.D, bg=bg,
                                 highlightthickness=0, bd=0)
        self._img_id = self._canvas.create_image(0, 0, anchor="nw")
        self._canvas.pack(side="left")
        self._label = tk.Label(self, text=text, bg=bg, fg=fg, font=font)
        self._label.pack(side="left", padx=(7, 0))
        for w in (self._canvas, self._label):
            w.bind("<Button-1>", self._on_click)
            w.config(cursor="hand2")
        self._var.trace_add("write", lambda *a: self._on_var_change())
        self._render(self._fill)

    def _render(self, fill):
        """Draw the ring + accent dot at fill amount `fill` (0..1), anti-aliased.
        The ring color eases from grey to accent and the dot grows from the centre
        as `fill` rises, so selecting/deselecting animates smoothly."""
        from PIL import Image, ImageDraw, ImageTk
        ss = self.SS
        D = self.D * ss
        bg = _hex_to_rgb(self._bg)
        img = Image.new("RGB", (D, D), bg)
        d = ImageDraw.Draw(img)
        ring = _lerp_rgb("#8b97a3", TOGGLE_ON, fill)
        bw = max(2, D // 11)
        d.ellipse([bw // 2, bw // 2, D - 1 - bw // 2, D - 1 - bw // 2],
                  outline=ring, width=bw, fill=(255, 255, 255))
        if fill > 0:
            # Full dot radius equals the old inset of 0.30*D from the edge; scale
            # it by `fill` so the dot grows out from the centre.
            r = (D / 2 - int(D * 0.30)) * fill
            c = D / 2
            d.ellipse([c - r, c - r, c + r, c + r], fill=_hex_to_rgb(TOGGLE_ON))
        img = img.resize((self.D, self.D), Image.LANCZOS)
        self._photo = ImageTk.PhotoImage(img)
        self._canvas.itemconfig(self._img_id, image=self._photo)

    def _animate_to(self, target):
        """Grow/shrink the accent dot from its current fill to `target` (0 or 1)."""
        if self._anim_id is not None:
            try:
                self.after_cancel(self._anim_id)
            except Exception:
                pass
            self._anim_id = None
        start = self._fill
        delta = target - start
        if abs(delta) < 1e-6:
            return

        def _step(i):
            if i >= self.FRAMES:
                self._fill = target
                self._render(target)
                self._anim_id = None
                return
            f = (i + 1) / self.FRAMES
            f = 1 - (1 - f) * (1 - f)            # ease-out
            self._fill = start + delta * f
            self._render(self._fill)
            self._anim_id = self.after(self.FRAME_MS, lambda: _step(i + 1))

        _step(0)

    def _on_var_change(self):
        # Fires on every instance sharing the variable: the newly selected radio
        # animates its dot in, the previously selected one animates its dot out.
        self._animate_to(1.0 if self._var.get() == self._value else 0.0)

    def _on_click(self, event=None):
        if self._var.get() != self._value:
            self._var.set(self._value)       # trace animates the whole group
            if self._command is not None:
                self._command()


def _apply_round_corners(win):
    """Best-effort Windows 11 slightly-rounded corners for a borderless Toplevel
    (DWM). No-op off Windows or if the call fails."""
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        win.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(win.winfo_id())
        # DWMWA_WINDOW_CORNER_PREFERENCE = 33, DWMWCP_ROUNDSMALL = 3
        pref = ctypes.c_int(3)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, 33, ctypes.byref(pref), ctypes.sizeof(pref))
    except Exception:
        pass


class _ModernDropdown(tk.Frame):
    """A readonly combobox-style selector with a Windows 11 feel: a white field +
    chevron that opens a popup which animates open (expanding downward) and has
    slightly rounded corners. Bound to a StringVar; calls `command` when the
    selection changes (mirroring a Combobox's <<ComboboxSelected>>)."""

    ROW_H = 28
    FRAMES, FRAME_MS = 7, 12

    def __init__(self, parent, textvariable, values, command=None,
                 width=26, font=("Segoe UI", 9), bg=APP_BG):
        super().__init__(parent, bg=bg)
        self._var = textvariable
        self._values = list(values)
        self._command = command
        self._font = font
        self._popup = None
        self._anim_id = None

        self._box = tk.Frame(self, bg="#ffffff", highlightthickness=1,
                             highlightbackground="#aab4bf", bd=0)
        self._box.pack()
        self._value_lbl = tk.Label(self._box, textvariable=self._var, bg="#ffffff",
                                   fg="#1a1a1a", font=font, anchor="w",
                                   width=width, padx=8, pady=4)
        self._value_lbl.pack(side="left")
        self._chevron = tk.Label(self._box, text="▾", bg="#ffffff", fg="#5a6b7b",
                                 font=("Segoe UI", 10), padx=6)
        self._chevron.pack(side="left")
        for w in (self._box, self._value_lbl, self._chevron):
            w.bind("<Button-1>", self._toggle_popup)
            w.config(cursor="hand2")

    def _toggle_popup(self, event=None):
        if self._popup is not None:
            self._close_popup()
        else:
            self._open_popup()

    def _open_popup(self):
        import tkinter.font as tkfont
        self.update_idletasks()
        x = self._box.winfo_rootx()
        y = self._box.winfo_rooty() + self._box.winfo_height()
        box_w = self._box.winfo_width()
        fm = tkfont.Font(font=self._font)
        text_w = max((fm.measure(v) for v in self._values), default=0)
        w = max(box_w, text_w + 28)
        full_h = self.ROW_H * len(self._values) + 4

        pop = tk.Toplevel(self)
        pop.overrideredirect(True)
        try:
            pop.attributes("-topmost", True)
        except Exception:
            pass
        pop.configure(bg="#ffffff")
        self._popup = pop

        panel = tk.Frame(pop, bg="#ffffff", highlightthickness=1,
                         highlightbackground="#aab4bf")
        panel.place(x=0, y=0, width=w, height=full_h)
        for i, val in enumerate(self._values):
            sel = (val == self._var.get())
            row = tk.Label(panel, text=val, bg=("#eaf2fb" if sel else "#ffffff"),
                           fg="#1a1a1a", font=self._font, anchor="w", padx=10)
            row.place(x=2, y=2 + i * self.ROW_H, width=w - 4, height=self.ROW_H)
            row.bind("<Enter>", lambda e, r=row: r.config(bg="#d6e7fa"))
            row.bind("<Leave>", lambda e, r=row, v=val: r.config(
                bg="#eaf2fb" if v == self._var.get() else "#ffffff"))
            row.bind("<Button-1>", lambda e, v=val: self._choose(v))
            row.config(cursor="hand2")

        _apply_round_corners(pop)

        pop.geometry(f"{w}x1+{x}+{y}")
        pop.bind("<Escape>", lambda e: self._close_popup())
        pop.bind("<Button-1>", self._maybe_close_outside)
        try:
            pop.grab_set()
        except Exception:
            pass
        self._animate_open(x, y, w, full_h, 0)

    def _animate_open(self, x, y, w, full_h, i):
        self._anim_id = None
        if self._popup is None:
            return
        if i > self.FRAMES:
            self._popup.geometry(f"{w}x{full_h}+{x}+{y}")
            return
        f = i / self.FRAMES
        f = 1 - (1 - f) * (1 - f)             # ease-out
        h = max(1, int(full_h * f))
        self._popup.geometry(f"{w}x{h}+{x}+{y}")
        # Schedule on self (a persistent widget), not the popup, so closing the
        # popup mid-animation can cleanly cancel the pending tick.
        self._anim_id = self.after(
            self.FRAME_MS, lambda: self._animate_open(x, y, w, full_h, i + 1))

    def _maybe_close_outside(self, event):
        pop = self._popup
        if pop is None:
            return
        px, py = pop.winfo_rootx(), pop.winfo_rooty()
        pw, ph = pop.winfo_width(), pop.winfo_height()
        if not (px <= event.x_root < px + pw and py <= event.y_root < py + ph):
            self._close_popup()

    def _choose(self, value):
        changed = (value != self._var.get())
        self._var.set(value)
        self._close_popup()
        if changed and self._command is not None:
            self._command()

    def _close_popup(self):
        if self._anim_id is not None:
            try:
                self.after_cancel(self._anim_id)
            except Exception:
                pass
            self._anim_id = None
        pop = self._popup
        self._popup = None
        if pop is not None:
            try:
                pop.grab_release()
            except Exception:
                pass
            try:
                pop.destroy()
            except Exception:
                pass


def _resource_base():
    """Directory where bundled assets live (next to the script, or sys._MEIPASS
    inside a one-file PyInstaller build)."""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    return Path(__file__).parent


def _asset_path(name):
    return _resource_base() / name


class _SplashScreen:
    """Small borderless splash shown while the app initializes.

    The 'b' logo runs to the left while the owl chases it from behind; the
    progress bar fills the opposite way (left→right) and 'Loading… X%' counts up.
    """
    W, H = 470, 230
    NAVY, GREEN, BG = "#073b5e", "#9aca3c", "#ffffff"

    def __init__(self, root):
        from PIL import ImageTk
        self.root = root
        self.win = tk.Toplevel(root)
        self.win.overrideredirect(True)          # no title bar / borders
        self.win.configure(bg=self.BG)
        sw, sh = self.win.winfo_screenwidth(), self.win.winfo_screenheight()
        self.win.geometry("%dx%d+%d+%d" % (
            self.W, self.H, (sw - self.W) // 2, (sh - self.H) // 2))
        try:
            self.win.attributes("-topmost", True)
        except Exception:
            pass

        self.cv = tk.Canvas(self.win, width=self.W, height=self.H, bg=self.BG,
                            highlightthickness=2, highlightbackground=self.NAVY)
        self.cv.pack(fill="both", expand=True)

        # sprites: navy owl silhouette + the colour logo
        self._imgs = {}
        self.ground_y = 92
        owl = ImageTk.PhotoImage(self._owl_sprite(70))
        logo = ImageTk.PhotoImage(self._logo_sprite(60))
        self._imgs["owl"], self._imgs["logo"] = owl, logo
        self.owl_w, self.logo_w = owl.width(), logo.width()
        self.logo_id = self.cv.create_image(-200, self.ground_y, image=logo)
        self.owl_id = self.cv.create_image(-200, self.ground_y, image=owl)

        # loading text below the sprites
        self.pct_id = self.cv.create_text(self.W // 2, 170,
                                          text="Loading… 0%", fill=self.NAVY,
                                          font=("Segoe UI", 11))
        self._render(0, 0)          # place sprites on-screen before first paint
        self.win.update()

    def _owl_sprite(self, h):
        from PIL import Image, ImageOps
        g = Image.open(_asset_path("owl_source.png")).convert("L")
        g = g.crop(ImageOps.invert(g).getbbox())   # trim white space
        alpha = ImageOps.invert(g)                  # dark owl -> opaque
        owl = Image.new("RGBA", g.size, self.NAVY)
        owl.putalpha(alpha)
        w = max(1, int(owl.width * h / owl.height))
        return owl.resize((w, h), Image.LANCZOS)

    def _logo_sprite(self, h):
        from PIL import Image
        logo = Image.open(_asset_path("logo_b.png")).convert("RGBA")
        w = max(1, int(logo.width * h / logo.height))
        return logo.resize((w, h), Image.LANCZOS)

    def _render(self, pct, frame):
        self.cv.itemconfigure(self.pct_id, text="Loading… %d%%" % int(pct))
        # One clean pass right -> left tied to the load %, so it runs through
        # exactly once (no looping) and finishes just as the app appears. At 0%
        # the logo is just barely peeking in at the right edge; the owl chases
        # ~140px behind (entering from off-screen right). Both exit off the left
        # edge by 100%.
        p = pct / 100.0
        gap = 100
        logo_start = self.W + self.logo_w / 2 - 12   # ~12px sliver at right edge
        owl_end = -self.owl_w / 2 - 10               # owl just off the left at 100%
        logo_end = owl_end - gap                      # so the trailing owl clears too
        logo_x = logo_start + (logo_end - logo_start) * p
        owl_x = logo_x + gap
        # slight hop as they run; owl bobs slightly out of phase with the logo
        A = 8
        logo_y = self.ground_y - A * abs(math.sin(frame * 0.30))
        owl_y = self.ground_y - A * abs(math.sin(frame * 0.30 + 0.9))
        self.cv.coords(self.logo_id, logo_x, logo_y)
        self.cv.coords(self.owl_id, owl_x, owl_y)

    def run(self, min_ms=1500):
        import time
        start = time.time()
        frame = 0
        while True:
            elapsed = (time.time() - start) * 1000.0
            self._render(min(100.0, elapsed / min_ms * 100.0), frame)
            self.root.update()
            if elapsed >= min_ms:
                break
            time.sleep(0.016)
            frame += 1
        self._render(100.0, frame)
        self.root.update()

    def close(self):
        try:
            self.win.destroy()
        except Exception:
            pass


class _OpTimer:
    """Live elapsed-time display for an operation: a 'Total Time Elapsed' timer
    plus a per-file timer, shown as small stacked labels under a progress bar.

    start()/stop() bracket the whole operation; set_file(name) is called as each
    file begins. The display refreshes ~4×/second on the Tk main thread.
    """
    def __init__(self, root, total_var, file_var, total_lbl, file_lbl):
        self.root = root
        self.total_var = total_var
        self.file_var = file_var
        self.total_lbl = total_lbl
        self.file_lbl = file_lbl
        self._start = None
        self._file_name = None
        self._file_start = None
        self._running = False
        self._shown = False
        self._final_secs = 0
        self._after_id = None   # pending root.after callback, so we never stack ticks

    def _cancel_pending(self):
        if self._after_id is not None:
            try:
                self.root.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    @staticmethod
    def _fmt(secs):
        secs = max(0, int(secs))
        h, rem = divmod(secs, 3600)
        m, s = divmod(rem, 60)
        return f"{h}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"

    def start(self):
        # Reveal the labels only once an operation actually begins.
        if not self._shown:
            self.total_lbl.pack(fill="x", padx=20)
            self.file_lbl.pack(fill="x", padx=20, pady=(0, 2))
            self._shown = True
        # Drop any tick still queued from a previous run so loops can't overlap.
        self._cancel_pending()
        self._start = time.monotonic()
        self._file_name = None
        self._file_start = None
        self._running = True
        self.total_var.set("Total Time Elapsed:  00:00")
        self.file_var.set("Current file:  —")
        self._tick()

    def set_file(self, name):
        # Called from the worker thread — only assigns plain attributes.
        self._file_name = name
        self._file_start = time.monotonic()

    def _tick(self):
        if not self._running:
            self._after_id = None
            return
        now = time.monotonic()
        if self._start is not None:
            self.total_var.set("Total Time Elapsed:  " + self._fmt(now - self._start))
        if self._file_name is not None and self._file_start is not None:
            self.file_var.set(
                f"Current file ({self._file_name}):  " + self._fmt(now - self._file_start))
        self._after_id = self.root.after(250, self._tick)

    def stop(self):
        if not self._running and self._start is None:
            return
        self._running = False
        self._cancel_pending()
        if self._start is not None:
            self._final_secs = time.monotonic() - self._start
            self.total_var.set("Total Time Elapsed:  " + self._fmt(self._final_secs))
        self.file_var.set("")

    def elapsed_str(self):
        """Formatted total elapsed time of the last run (call after stop())."""
        return self._fmt(self._final_secs)

    def reset(self):
        """Clear the timer and hide it from the window (ready for the next run)."""
        self._running = False
        self._cancel_pending()
        self._start = None
        self._file_name = None
        self._file_start = None
        self.total_var.set("")
        self.file_var.set("")
        if self._shown:
            self.total_lbl.pack_forget()
            self.file_lbl.pack_forget()
            self._shown = False


class _ProgressTracker:
    """Drives a ttk.Progressbar to reflect BOTH completed items and a smooth
    'working on the current item' trickle — reusable by every tool.

    File-count progress is exact: for N items with K done, the bar sits at K/N.
    True per-item progress isn't available (COM/markitdown don't report how far
    through a single file they are), so while an item is being worked on the bar
    EASES forward toward the next item's mark without quite reaching it, then
    snaps exactly when the item completes. That gives an honest 'moving little by
    little' feel. The bar does NOT move until an item is actually being
    processed: the worker calls `begin_item()` when a file starts (which kicks
    off the trickle) and `complete_item()` when it finishes (which snaps the bar
    and stops the trickle until the next `begin_item`). So between files — and
    before the first file starts — the bar holds still.

    Thread model: `start`/`finish`/`reset` are called on the Tk thread; the
    worker-thread signals `begin_item()`/`complete_item()` marshal onto the Tk
    thread via `root.after`. The bar runs 0..SCALE so fractional motion stays
    smooth.
    """
    SCALE = 1000        # bar maximum, so fractional movement looks continuous
    TICK_MS = 50        # trickle animation cadence
    EASE = 0.08         # fraction of the remaining gap covered per tick
    CAP = 0.9           # within one item, ease at most this far to the next mark

    def __init__(self, root, bar):
        self.root = root
        self.bar = bar
        self._total = 1
        self._done = 0
        self._trickle = 0.0          # 0..1 progress *within* the current item
        self._anim_id = None
        self._running = False

    def start(self, total):
        """Begin an operation over `total` items (Tk thread). The bar starts empty
        and stays still until the first `begin_item()`."""
        self._total = max(1, int(total))
        self._done = 0
        self._trickle = 0.0
        self._running = True
        self._cancel_anim()
        self.bar.config(maximum=self.SCALE)
        self.bar["value"] = 0

    def begin_item(self):
        """Signal that processing of the next item has actually begun (safe from a
        worker thread). Starts the within-item trickle toward the next mark."""
        self.root.after(0, self._begin_item_main)

    def complete_item(self):
        """Mark one item done (safe from a worker thread)."""
        self.root.after(0, self._complete_item_main)

    def finish(self):
        """End the operation, stopping the trickle (Tk thread)."""
        self._running = False
        self._cancel_anim()
        self._render()

    def reset(self):
        """Clear the bar back to empty (Tk thread)."""
        self._running = False
        self._cancel_anim()
        self._done = 0
        self._trickle = 0.0
        self.bar["value"] = 0

    def _begin_item_main(self):
        if not self._running:
            return
        self._trickle = 0.0
        self._cancel_anim()
        self._animate()              # trickle toward this item's completion

    def _complete_item_main(self):
        if not self._running:
            return
        self._done = min(self._done + 1, self._total)
        self._trickle = 0.0
        self._cancel_anim()          # stop trickling; wait for the next begin_item
        self._render()               # snap to the exact completed fraction
        if self._done >= self._total:
            self._running = False    # everything done; bar sits full

    def _animate(self):
        if not self._running:
            self._anim_id = None
            return
        self._trickle += (self.CAP - self._trickle) * self.EASE
        self._render()
        self._anim_id = self.root.after(self.TICK_MS, self._animate)

    def _render(self):
        frac = (self._done + self._trickle) / self._total
        frac = max(0.0, min(1.0, frac))
        self.bar["value"] = frac * self.SCALE

    def _cancel_anim(self):
        if self._anim_id is not None:
            try:
                self.root.after_cancel(self._anim_id)
            except Exception:
                pass
            self._anim_id = None


class ConverterApp:
    def __init__(self):
        # Give the process its own AppUserModelID *before* the window exists so
        # Windows ties the taskbar button to our icon, not python/Tk's default.
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "BoxOfScraps.App")
        except Exception:
            pass

        if HAS_DND:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()

        # Surface (rather than silently swallow) any exception raised inside a Tk
        # callback. In the windowed build there is no console, so the default
        # handler's traceback goes nowhere and a wedged callback looks like a
        # freeze; this shows it once and keeps the app running.
        self._error_dialog_open = False
        self.root.report_callback_exception = self._on_tk_error

        # Hide the window *immediately* — before any geometry/icon work forces a
        # paint — so the blank Tk window never flashes on screen at launch.
        self.root.withdraw()

        self.root.title("Box of Scraps")
        self.root.resizable(True, True)
        self.root.minsize(860, 600)
        self.root.configure(bg=APP_BG)
        self._set_icon()

        # Show an animated splash while we build the UI.
        splash = None
        try:
            splash = _SplashScreen(self.root)
        except Exception:
            splash = None

        # Close PyInstaller's built-in bootloader splash (shown during the
        # one-file unpack) now that our own splash is on screen. This MUST be a
        # static `import pyi_splash` — PyInstaller only bundles the pyi_splash
        # runtime module when it sees a static import, so a dynamic import would
        # leave it out of the build and the bootloader splash would never close
        # (it then resurfaces on top of the app). The module is absent in
        # dev/editor, hence the guard; the type-ignore silences that editor hint.
        try:
            import pyi_splash  # type: ignore
            pyi_splash.close()
        except Exception:
            pass

        # PDF page state
        self.files    = []
        self._out_dir = None
        self._mode    = tk.StringVar(value="Email (.eml, .msg)")
        self._cancel_convert = threading.Event()
        self._pdf_app_pids   = []   # PIDs of dedicated Office instances (not Outlook)

        # Unzip page state
        self._zip_files     = []
        self._unzip_out_dir = None
        self._cancel_unzip  = threading.Event()

        # AI Preparation page state
        self._ai_files     = []
        self._ai_out_dir   = None
        self._cancel_ai    = threading.Event()

        # Office Modernizer page state
        self._office_files    = []
        self._office_out_dir  = None
        self._cancel_office   = threading.Event()
        self._office_app_pids = {}   # family -> set of PIDs of our dedicated instances

        # Password Removal page state
        self._pwd_files   = []
        self._pwd_out_dir = None
        self._cancel_pwd  = threading.Event()
        self._pwd_dialog  = None     # the open password prompt, if any

        # Document Structuring page state
        self._ds_files    = []
        self._ds_out_dir  = None
        self._cancel_ds   = threading.Event()
        self._ds_rename   = tk.BooleanVar(value=False)
        self._ds_date     = tk.BooleanVar(value=False)
        self._ds_organize = tk.BooleanVar(value=True)

        # Updates page state
        self._upd_release   = None    # the newer release dict, once one is found
        self._upd_asset     = None    # its BoxOfScraps.exe asset
        self._upd_busy      = False   # guards against two update runs at once
        self._upd_last_pct  = 0.0     # throttles download progress UI updates
        self._cancel_update = threading.Event()
        self._upd_cooldown_until = 0.0    # monotonic deadline; see _upd_begin_cooldown
        self._upd_cooldown_after = None   # pending countdown tick, so none stack

        # Clear what a previous self-update left behind. The old exe can't be
        # deleted by the update that replaced it — at that moment it is still
        # the running image — so it is cleaned up on the next launch instead.
        _exe = _current_exe_path()
        if _exe is not None:
            _cleanup_previous_update(_exe)

        # Navigation state
        self._page_frames = {}
        self._nav_buttons = {}
        self._nav_rest_fg = {}   # per-entry unselected fg (tools vs. info footer)

        self._build_shell()
        self._show_page("pdf")

        # Run the splash animation, then reveal the main window.
        if splash is not None:
            try:
                splash.run(min_ms=1500)
            except Exception:
                pass

        # Reveal the main window fully painted *before* removing the splash, so
        # no blank/white frame shows in between: position it, map it invisible
        # (alpha 0), force a full paint, make it opaque, then drop the splash.
        self._center_window(1000, 680)
        try:
            self.root.attributes("-alpha", 0.0)
        except Exception:
            pass
        self.root.deiconify()
        self.root.update()
        try:
            self.root.attributes("-alpha", 1.0)
        except Exception:
            pass
        self.root.lift()
        if splash is not None:
            splash.close()

    def _on_tk_error(self, exc, val, tb):
        """Handle an uncaught exception from any Tk callback: log the traceback
        and show a single, non-blocking-storm error dialog. Guarded so a callback
        that errors repeatedly (e.g. a periodic `after`) can't stack dialogs."""
        import traceback
        try:
            sys.stderr.write("".join(traceback.format_exception(exc, val, tb)))
        except Exception:
            pass
        if self._error_dialog_open:
            return
        self._error_dialog_open = True
        try:
            messagebox.showerror(
                "Unexpected Error",
                "Something went wrong and the action couldn't be completed:\n\n"
                f"{type(val).__name__}: {val}\n\n"
                "The app is still running — you can try again.",
            )
        except Exception:
            pass
        finally:
            self._error_dialog_open = False

    def _set_icon(self):
        """Load icon.ico and set both the title bar and taskbar icon at a legible size."""
        if getattr(sys, "frozen", False):
            # One-file PyInstaller build extracts bundled data to sys._MEIPASS,
            # not next to the .exe — fall back to the exe dir for one-dir builds.
            base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
        else:
            base = Path(__file__).parent
        ico = base / "icon.ico"
        if not ico.exists():
            return
        try:
            # iconbitmap — sets the title bar icon (Windows native)
            self.root.iconbitmap(str(ico))
        except Exception:
            pass
        try:
            # iconphoto — sets the Alt+Tab icon and the default for new toplevels.
            from PIL import Image, ImageTk
            src = Image.open(str(ico)).convert("RGBA")
            self._icon_photos = [
                ImageTk.PhotoImage(src.resize((n, n), Image.LANCZOS))
                for n in (16, 32, 48, 64, 128, 256)
            ]  # keep references alive
            self.root.iconphoto(True, *self._icon_photos)
        except Exception:
            pass
        # The Windows taskbar button uses the window's ICON_BIG, and Tk's
        # iconphoto doesn't reliably set it — so push both icons to the native
        # top-level window via WM_SETICON. This is what makes the owl (not Tk's
        # default feather) show on the taskbar while the app is running.
        try:
            import ctypes
            self.root.update_idletasks()  # ensure the native HWND exists
            user32 = ctypes.windll.user32
            user32.LoadImageW.restype = ctypes.c_void_p
            user32.SendMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint,
                                            ctypes.c_void_p, ctypes.c_void_p]
            hwnd = user32.GetParent(self.root.winfo_id()) or self.root.winfo_id()
            IMAGE_ICON, LR_LOADFROMFILE = 1, 0x00000010
            WM_SETICON, ICON_SMALL, ICON_BIG = 0x0080, 0, 1
            self._hicons = []
            for size, which in ((32, ICON_BIG), (16, ICON_SMALL)):
                hicon = user32.LoadImageW(None, str(ico), IMAGE_ICON,
                                          size, size, LR_LOADFROMFILE)
                if hicon:
                    self._hicons.append(hicon)  # keep handles alive
                    user32.SendMessageW(hwnd, WM_SETICON, which, hicon)
        except Exception:
            pass

    # ── shell / navigation ───────────────────
    def _build_shell(self):
        """Create the two-column sidebar + content frame layout."""
        self._sidebar = tk.Frame(self.root, bg=SIDEBAR_BG, width=SIDEBAR_W)
        self._sidebar.pack(side="left", fill="y")
        self._sidebar.pack_propagate(False)

        self._content = tk.Frame(self.root, bg=APP_BG)
        self._content.pack(side="left", fill="both", expand=True)

        tk.Label(
            self._sidebar,
            text="Box of\nScraps",
            bg=SIDEBAR_BG, fg="#ffffff",
            font=("Segoe UI", 11, "bold"),
            justify="center", pady=20,
        ).pack(fill="x")

        tk.Frame(self._sidebar, bg=SIDEBAR_RULE, height=1).pack(fill="x", padx=12, pady=(0, 8))

        # The tools — the things the app actually does.
        for page_key, label in [("pdf", "  PDF Conversion"),
                                ("unzip", "  Folder Unzipping"),
                                ("aiprep", "  Markdown Conversion"),
                                ("office", "  Office Modernizer"),
                                ("pwd", "  Password Removal"),
                                ("structure", "  Document Structuring")]:
            btn = tk.Button(
                self._sidebar,
                text=label,
                bg=SIDEBAR_BG, fg=SIDEBAR_FG,
                font=("Segoe UI", 10),
                relief="flat", anchor="w",
                padx=12, pady=10,
                cursor="hand2",
                activebackground=SIDEBAR_SEL,
                activeforeground="#ffffff",
                bd=0,
                command=lambda k=page_key: self._show_page(k),
            )
            btn.pack(fill="x")
            self._nav_buttons[page_key] = btn
            self._nav_rest_fg[page_key] = SIDEBAR_FG

        # Informational footer, pinned to the bottom of the sidebar and set apart
        # from the tool list above: this page is about the app itself, not another
        # feature. Packed bottom-up, so the order here is version line first (it
        # ends up lowest), then the nav entry, then the rule that separates the
        # whole footer from the tools.
        tk.Label(
            self._sidebar, text=f"Version {APP_VERSION}",
            bg=SIDEBAR_BG, fg=SIDEBAR_INFO_FG, font=("Segoe UI", 8),
            anchor="w", padx=22,   # aligns under the nav entry's text above it
        ).pack(side="bottom", fill="x", pady=(4, 12))

        info_btn = tk.Button(
            self._sidebar,
            text="  About & Updates",
            bg=SIDEBAR_BG, fg=SIDEBAR_INFO_FG,
            font=("Segoe UI", 9),
            relief="flat", anchor="w",
            padx=12, pady=8,
            cursor="hand2",
            activebackground=SIDEBAR_INFO_SEL,
            activeforeground="#ffffff",
            bd=0,
            command=lambda: self._show_page("update"),
        )
        info_btn.pack(side="bottom", fill="x")
        self._nav_buttons["update"] = info_btn
        self._nav_rest_fg["update"] = SIDEBAR_INFO_FG

        tk.Frame(self._sidebar, bg=SIDEBAR_RULE, height=1).pack(
            side="bottom", fill="x", padx=12, pady=(8, 8))

    def _show_page(self, key: str):
        """Switch the visible content page, building it lazily on first visit."""
        # Each entry returns to its OWN resting colour — the tools sit brighter
        # than the informational footer entry, so they can't share one value.
        for nav_key, btn in self._nav_buttons.items():
            btn.config(bg=SIDEBAR_BG, fg=self._nav_rest_fg[nav_key])
        self._nav_buttons[key].config(
            bg=(SIDEBAR_INFO_SEL if key == "update" else SIDEBAR_SEL),
            fg="#ffffff")

        for frame in self._page_frames.values():
            frame.pack_forget()

        if key not in self._page_frames:
            frame = tk.Frame(self._content, bg=APP_BG)
            self._page_frames[key] = frame
            if key == "pdf":
                self._build_pdf_page(frame)
            elif key == "unzip":
                self._build_unzip_page(frame)
            elif key == "aiprep":
                self._build_aiprep_page(frame)
            elif key == "office":
                self._build_office_page(frame)
            elif key == "pwd":
                self._build_pwd_page(frame)
            elif key == "structure":
                self._build_structure_page(frame)
            elif key == "update":
                self._build_update_page(frame)
            # Give this page's flat grey utility buttons a themed press-in effect.
            self._enhance_page_buttons(frame)

        self._page_frames[key].pack(fill="both", expand=True)

    def _enhance_page_buttons(self, widget):
        """Recursively give a freshly-built page's flat grey utility buttons
        (the #e0e8f0 Add / Remove / Clear / Choose… buttons) a press-in effect.
        Buttons that already define their own active colors — the accent action
        buttons, the red Cancel, the sidebar nav — are left untouched, so only
        the ones that would otherwise flash Tk's grey 'active' look are styled.
        Runs per page on first build, so new pages pick this up automatically."""
        for child in widget.winfo_children():
            if isinstance(child, tk.Button):
                try:
                    base = str(child.cget("bg")).lower()
                except Exception:
                    base = ""
                if base == "#e0e8f0":
                    _style_press_button(child)
            self._enhance_page_buttons(child)

    def _build_timer(self, parent):
        """Reserve a spot for the stacked 'Total Time Elapsed' + current-file
        labels (small, left-aligned). The labels stay hidden until the operation
        starts; the holder frame keeps their position without taking space when
        empty. Returns the _OpTimer that reveals and drives them."""
        holder = tk.Frame(parent, bg=APP_BG)
        holder.pack(fill="x")
        total_var = tk.StringVar(value="")
        file_var = tk.StringVar(value="")
        total_lbl = tk.Label(holder, textvariable=total_var, bg=APP_BG, fg="#666",
                             font=("Segoe UI", 8, "bold"), anchor="w")
        file_lbl = tk.Label(holder, textvariable=file_var, bg=APP_BG, fg="#888",
                            font=("Segoe UI", 8), anchor="w")
        return _OpTimer(self.root, total_var, file_var, total_lbl, file_lbl)

    # ── layout ──────────────────────────────
    def _build_pdf_page(self, parent):
        # Section header
        tk.Label(
            parent,
            text="PDF Conversion",
            bg=APP_BG, fg="#1a1a1a",
            font=("Segoe UI", 14, "bold"),
            anchor="w", padx=18,
        ).pack(fill="x", pady=(14, 4))

        tk.Label(
            parent,
            text=("Convert files to PDF. Pick a conversion mode, add the matching files, "
                  "and each one is saved as a PDF in your chosen folder."),
            bg=APP_BG, fg="#555", font=("Segoe UI", 9),
            anchor="w", padx=18, justify="left", wraplength=620,
        ).pack(fill="x", pady=(0, 4))

        # Mode selector row
        mode_frame = tk.Frame(parent, bg=APP_BG)
        mode_frame.pack(fill="x", padx=18, pady=(10, 2))
        tk.Label(
            mode_frame, text="Conversion mode:",
            bg=APP_BG, fg="#555", font=("Segoe UI", 9),
        ).pack(side="left")
        mode_menu = _ModernDropdown(
            mode_frame,
            textvariable=self._mode,
            values=["Email (.eml, .msg)",
                    "Text Report (.rep, .rpt, .prn, .lst)",
                    "HTML (.html, .htm, .mht)",
                    "Image (.jpg, .png, .bmp, .gif, .tiff)",
                    "Word (.doc, .docx, .txt, .rtf, .xml)",
                    "Excel (.xls, .xlsx, .csv)",
                    "PowerPoint (.ppt, .pptx)",
                    "Visio (.vsd, .vsdx)"],
            command=self._on_mode_change,
            width=26,
            font=("Segoe UI", 9),
        )
        mode_menu.pack(side="left", padx=8)

        # Drop zone
        self.drop_frame = tk.Frame(
            parent, bg=DROP_BG,
            highlightbackground=DROP_BD,
            highlightthickness=2,
            relief="flat",
        )
        self.drop_frame.pack(fill="both", expand=True, padx=18, pady=(16, 6))

        self.drop_label = tk.Label(
            self.drop_frame,
            text="Drop .eml / .msg files here\nor click 'Add Files'",
            bg=DROP_BG, fg="#4a6fa5",
            font=("Segoe UI", 11),
            justify="center",
        )
        self.drop_label.pack(expand=True, pady=20)
        # File list
        list_frame = tk.Frame(self.drop_frame, bg=DROP_BG)
        list_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        self.file_list = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            selectmode="extended",
            bg="#ffffff", fg="#1a1a1a",
            font=("Segoe UI", 9),
            relief="flat", bd=1,
            activestyle="none",
            highlightthickness=0,
        )
        scrollbar.config(command=self.file_list.yview)
        scrollbar.pack(side="right", fill="y")
        self.file_list.pack(fill="both", expand=True)

        if HAS_DND:
            for widget in (self.drop_frame, self.drop_label, list_frame, self.file_list):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._on_drop)

        # Buttons row
        btn_frame = tk.Frame(parent, bg=APP_BG)
        btn_frame.pack(fill="x", padx=18, pady=(4, 4))

        for label, cmd in [
            ("Add Files",        self._add_files),
            ("Remove Selected",  self._remove_selected),
            ("Clear All",        self._clear_files),
        ]:
            tk.Button(
                btn_frame, text=label, command=cmd,
                bg="#e0e8f0", fg="#333",
                font=("Segoe UI", 9), relief="flat",
                padx=12, pady=4, cursor="hand2",
            ).pack(side="left", padx=(0, 6))

        # Output folder row
        out_frame = tk.Frame(parent, bg=APP_BG)
        out_frame.pack(fill="x", padx=18, pady=(2, 2))

        tk.Label(
            out_frame, text="Output folder:",
            bg=APP_BG, fg="#555", font=("Segoe UI", 9),
        ).pack(side="left")

        self.out_var = tk.StringVar(value="Same as source file")
        tk.Label(
            out_frame, textvariable=self.out_var,
            bg=APP_BG, fg=ACCENT, font=("Segoe UI", 9, "italic"),
        ).pack(side="left", padx=4)

        tk.Button(
            out_frame, text="Choose…", command=self._choose_output,
            bg="#e0e8f0", fg="#333",
            font=("Segoe UI", 9), relief="flat",
            padx=8, pady=2, cursor="hand2",
        ).pack(side="left", padx=4)

        # Progress bar
        self.progress = ttk.Progressbar(parent, mode="determinate")
        self.progress.pack(fill="x", padx=18, pady=(4, 2))
        self._pdf_tracker = _ProgressTracker(self.root, self.progress)

        # Status label
        self.status_var = tk.StringVar(value="Ready — add files to get started.")
        tk.Label(
            parent, textvariable=self.status_var,
            bg=APP_BG, fg="#555", font=("Segoe UI", 8),
            anchor="w",
        ).pack(fill="x", padx=20, pady=(0, 4))
        self._pdf_timer = self._build_timer(parent)

        # Convert button (+ smaller Cancel button beside it)
        btn_row = tk.Frame(parent, bg=APP_BG)
        btn_row.pack(pady=(4, 14))
        self.convert_btn = tk.Button(
            btn_row, text="Convert",
            command=self._start_convert,
            bg=ACCENT, fg=BTN_FG,
            font=("Segoe UI", 12, "bold"),
            relief="flat", padx=30, pady=8,
            cursor="hand2",
            activebackground="#005a9e",
            activeforeground=BTN_FG,
        )
        self.convert_btn.pack(side="left")
        self.cancel_btn = tk.Button(
            btn_row, text="Cancel",
            command=self._cancel_conversion,
            font=("Segoe UI", 10, "bold"),
            relief="flat", padx=16, pady=6,
            activebackground=CANCEL_BG_ACTIVE,
            activeforeground=BTN_FG,
        )
        self.cancel_btn.pack(side="left", padx=(10, 0))
        self._set_cancel_state(self.cancel_btn, False)

    # ── unzip page ───────────────────────────
    def _build_unzip_page(self, parent):
        tk.Label(
            parent,
            text="Folder Unzipping",
            bg=APP_BG, fg="#1a1a1a",
            font=("Segoe UI", 14, "bold"),
            anchor="w", padx=18,
        ).pack(fill="x", pady=(14, 4))

        tk.Label(
            parent,
            text=("Extract .zip files to your chosen folder. Nested zips (zips inside "
                  "zips) are unpacked automatically, all the way down."),
            bg=APP_BG, fg="#555", font=("Segoe UI", 9),
            anchor="w", padx=18, justify="left", wraplength=620,
        ).pack(fill="x", pady=(0, 4))

        # Drop zone
        uz_drop_frame = tk.Frame(
            parent, bg=DROP_BG,
            highlightbackground=DROP_BD,
            highlightthickness=2,
            relief="flat",
        )
        uz_drop_frame.pack(fill="both", expand=True, padx=18, pady=(10, 6))

        self._uz_drop_label = tk.Label(
            uz_drop_frame,
            text="Drop .zip files here\nor click 'Add Zip Files'",
            bg=DROP_BG, fg="#4a6fa5",
            font=("Segoe UI", 11),
            justify="center",
        )
        self._uz_drop_label.pack(expand=True, pady=20)

        uz_list_frame = tk.Frame(uz_drop_frame, bg=DROP_BG)
        uz_list_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        uz_scrollbar = tk.Scrollbar(uz_list_frame, orient="vertical")
        self._uz_file_list = tk.Listbox(
            uz_list_frame,
            yscrollcommand=uz_scrollbar.set,
            selectmode="extended",
            bg="#ffffff", fg="#1a1a1a",
            font=("Segoe UI", 9),
            relief="flat", bd=1,
            activestyle="none",
            highlightthickness=0,
        )
        uz_scrollbar.config(command=self._uz_file_list.yview)
        uz_scrollbar.pack(side="right", fill="y")
        self._uz_file_list.pack(fill="both", expand=True)

        if HAS_DND:
            for widget in (uz_drop_frame, self._uz_drop_label, uz_list_frame, self._uz_file_list):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._uz_on_drop)

        # Buttons row
        uz_btn_frame = tk.Frame(parent, bg=APP_BG)
        uz_btn_frame.pack(fill="x", padx=18, pady=(4, 4))

        for label, cmd in [
            ("Add Zip Files",    self._uz_add_files),
            ("Remove Selected",  self._uz_remove_selected),
            ("Clear All",        self._uz_clear_files),
        ]:
            tk.Button(
                uz_btn_frame, text=label, command=cmd,
                bg="#e0e8f0", fg="#333",
                font=("Segoe UI", 9), relief="flat",
                padx=12, pady=4, cursor="hand2",
            ).pack(side="left", padx=(0, 6))

        # Output folder row
        uz_out_frame = tk.Frame(parent, bg=APP_BG)
        uz_out_frame.pack(fill="x", padx=18, pady=(2, 2))

        tk.Label(
            uz_out_frame, text="Output folder:",
            bg=APP_BG, fg="#555", font=("Segoe UI", 9),
        ).pack(side="left")

        self._uz_out_var = tk.StringVar(value="Choose an output folder")
        self._uz_out_label = tk.Label(
            uz_out_frame, textvariable=self._uz_out_var,
            bg=APP_BG, fg="#c0392b", font=("Segoe UI", 9, "italic"),
        )
        self._uz_out_label.pack(side="left", padx=4)

        tk.Button(
            uz_out_frame, text="Choose…", command=self._uz_choose_output,
            bg="#e0e8f0", fg="#333",
            font=("Segoe UI", 9), relief="flat",
            padx=8, pady=2, cursor="hand2",
        ).pack(side="left", padx=4)

        # Organization options
        uz_org_frame = tk.Frame(parent, bg=APP_BG)
        uz_org_frame.pack(fill="x", padx=18, pady=(6, 2))

        tk.Label(
            uz_org_frame, text="Output organization:",
            bg=APP_BG, fg="#555", font=("Segoe UI", 9),
        ).pack(side="left", padx=(0, 10))

        self._uz_separate = tk.BooleanVar(value=True)
        for text, val in [
            ("Each zip into its own folder", True),
            ("All into one folder",          False),
        ]:
            _ModernRadio(
                uz_org_frame, text=text, variable=self._uz_separate, value=val,
                font=("Segoe UI", 9),
            ).pack(side="left", padx=(0, 16))

        # Auto-PDF checkbox
        uz_autopdf_frame = tk.Frame(parent, bg=APP_BG)
        uz_autopdf_frame.pack(fill="x", padx=18, pady=(6, 2))

        self._uz_auto_pdf = tk.BooleanVar(value=False)
        _ToggleSwitch(
            uz_autopdf_frame,
            text="Auto-convert files to PDF after unzipping  (originals deleted on success)",
            variable=self._uz_auto_pdf,
        ).pack(side="left")

        # Progress bar
        self._uz_progress = ttk.Progressbar(parent, mode="determinate")
        self._uz_progress.pack(fill="x", padx=18, pady=(6, 2))
        self._uz_tracker = _ProgressTracker(self.root, self._uz_progress)

        # Status label
        self._uz_status_var = tk.StringVar(value="Ready — add .zip files to get started.")
        tk.Label(
            parent, textvariable=self._uz_status_var,
            bg=APP_BG, fg="#555", font=("Segoe UI", 8),
            anchor="w",
        ).pack(fill="x", padx=20, pady=(0, 4))
        self._uz_timer = self._build_timer(parent)

        # Unzip button (disabled until output folder is chosen) + Cancel beside it
        uz_btn_row = tk.Frame(parent, bg=APP_BG)
        uz_btn_row.pack(pady=(4, 14))
        self._uz_btn = tk.Button(
            uz_btn_row, text="Unzip Files",
            command=self._start_unzip,
            bg=ACCENT, fg=BTN_FG,
            font=("Segoe UI", 12, "bold"),
            relief="flat", padx=30, pady=8,
            cursor="hand2",
            activebackground="#005a9e",
            activeforeground=BTN_FG,
            state="disabled",
        )
        self._uz_btn.pack(side="left")
        self._uz_cancel_btn = tk.Button(
            uz_btn_row, text="Cancel",
            command=self._cancel_unzip_op,
            font=("Segoe UI", 10, "bold"),
            relief="flat", padx=16, pady=6,
            activebackground=CANCEL_BG_ACTIVE,
            activeforeground=BTN_FG,
        )
        self._uz_cancel_btn.pack(side="left", padx=(10, 0))
        self._set_cancel_state(self._uz_cancel_btn, False)

    def _uz_add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select ZIP files",
            filetypes=[("ZIP archives", "*.zip"), ("All files", "*.*")],
        )
        for p in paths:
            self._uz_add_path(p)

    def _uz_on_drop(self, event):
        for p in self.root.tk.splitlist(event.data):
            self._uz_add_path(p)

    def _uz_add_path(self, p):
        path = Path(p)
        if path.suffix.lower() != ".zip":
            return
        if path not in self._zip_files:
            self._zip_files.append(path)
            self._uz_file_list.insert("end", path.name)
        self._uz_update_drop_label()

    def _uz_remove_selected(self):
        for i in sorted(self._uz_file_list.curselection(), reverse=True):
            self._uz_file_list.delete(i)
            del self._zip_files[i]
        self._uz_update_drop_label()

    def _uz_clear_files(self):
        self._zip_files.clear()
        self._uz_file_list.delete(0, "end")
        self._uz_update_drop_label()
        # Return the page to its ready state (clear any "Done!"/error status).
        self._uz_status_var.set("Ready — add .zip files to get started.")
        self._uz_tracker.reset()

    def _uz_update_drop_label(self):
        if self._zip_files:
            self._uz_drop_label.config(text=f"{len(self._zip_files)} zip file(s) queued")
        else:
            self._uz_drop_label.config(text="Drop .zip files here\nor click 'Add Zip Files'")

    def _uz_choose_output(self):
        d = filedialog.askdirectory(title="Select an empty output folder")
        if d:
            path = Path(d)
            try:
                non_empty = any(path.iterdir())
            except Exception:
                non_empty = False
            if non_empty:
                messagebox.showerror(
                    "Folder Not Empty",
                    "The selected destination must be empty.\n\n"
                    f"\"{path}\" already contains files or folders.\n\n"
                    "Please choose an empty folder, or use the dialog's "
                    "\"New folder\" option to create one for the unzipped output.",
                )
                return   # keep the previous (valid) selection, if any
            self._unzip_out_dir = path
            self._uz_out_var.set(str(self._unzip_out_dir))
            self._uz_out_label.config(fg=ACCENT)
            self._uz_btn.config(state="normal")
        elif self._unzip_out_dir is None:
            self._uz_out_var.set("Choose an output folder")
            self._uz_out_label.config(fg="#c0392b")

    def _start_unzip(self):
        if not self._zip_files:
            messagebox.showwarning("No Files", "Please add ZIP files first.")
            return
        if self._unzip_out_dir is None:
            messagebox.showwarning("No Output Folder", "Please choose an output folder first.")
            return
        try:
            not_empty = any(self._unzip_out_dir.iterdir())
        except Exception:
            not_empty = False
        if not_empty:
            messagebox.showerror(
                "Folder Not Empty",
                "The output folder is no longer empty.\n\n"
                f"\"{self._unzip_out_dir}\" now contains files or folders.\n\n"
                "Please choose an empty folder for the unzipped output.",
            )
            self._unzip_out_dir = None
            self._uz_out_var.set("Choose an output folder")
            self._uz_out_label.config(fg="#c0392b")
            self._uz_btn.config(state="disabled")
            return
        self._cancel_unzip.clear()
        self._uz_btn.config(state="disabled")
        self._set_cancel_state(self._uz_cancel_btn, True)
        self._uz_tracker.start(len(self._zip_files))
        separate  = self._uz_separate.get()
        auto_pdf  = self._uz_auto_pdf.get()
        self._uz_timer.start()
        threading.Thread(
            target=_unzip_worker,
            args=(
                list(self._zip_files),
                self._unzip_out_dir,
                separate,
                auto_pdf,
                lambda msg: self.root.after(0, lambda m=msg: self._uz_status_var.set(m)),
                self._uz_tracker.complete_item,
                lambda s, f, pf, pr: self.root.after(0, lambda: self._uz_finish(s, f, pf, pr)),
                self._cancel_unzip,
                lambda name: (self._uz_timer.set_file(name), self._uz_tracker.begin_item()),
            ),
            daemon=True,
        ).start()

    def _cancel_unzip_op(self):
        """Request that the running unzip/auto-convert stop as soon as possible."""
        self._cancel_unzip.set()
        self._set_cancel_state(self._uz_cancel_btn, False)
        self._uz_status_var.set("Cancelling… finishing the current item, then stopping.")

    def _uz_finish(self, success, failed, pdf_failed=None, protected=None):
        self._uz_timer.stop()
        self._uz_tracker.finish()
        elapsed = self._uz_timer.elapsed_str()
        self._uz_timer.reset()
        self._uz_btn.config(state="normal")
        self._set_cancel_state(self._uz_cancel_btn, False)
        protected = protected or []
        # ZIP passwords aren't Office encryption, so the Password Removal tool
        # can't help here — give archive-appropriate guidance instead.
        pw_note = _password_skip_note(
            protected,
            advice=("These archives are password-protected and can't be extracted here. "
                    "Open them with your unzip program using the password. (The Password "
                    "Removal tool only unlocks Office documents, not ZIP files.)"))
        skip_status = f"  {len(protected)} skipped (password-protected)." if protected else ""
        if self._cancel_unzip.is_set():
            self._uz_status_var.set(f"⏹  Cancelled. {success} archive(s) extracted before stopping.")
            messagebox.showinfo(
                "Unzip Cancelled",
                f"Operation was cancelled.\n{success} archive(s) extracted before stopping."
                f"\n\nTotal time elapsed: {elapsed}{pw_note}",
            )
            return
        if not failed:
            self._uz_status_var.set(f"Done! {success} archive(s) extracted successfully.{skip_status}")
            messagebox.showinfo(
                "Unzip Complete",
                f"{success} archive(s) extracted successfully."
                f"\n\nTotal time elapsed: {elapsed}{pw_note}",
            )
        else:
            err_lines = "\n".join(f"• {n}: {e}" for n, e in failed)
            self._uz_status_var.set(f"{success} extracted, {len(failed)} failed.{skip_status}")
            messagebox.showerror(
                "Unzip Errors",
                f"{success} succeeded, {len(failed)} failed:\n\n{err_lines}"
                f"\n\nTotal time elapsed: {elapsed}{pw_note}",
            )

        if pdf_failed:
            file_list = "\n".join(f"• {name}" for name in pdf_failed)
            messagebox.showwarning(
                "PDF Conversion Errors",
                f"The following files could not be converted to PDF:\n\n{file_list}",
            )

    # ── AI Preparation page ──────────────────
    def _build_aiprep_page(self, parent):
        tk.Label(
            parent, text="Markdown Conversion",
            bg=APP_BG, fg="#1a1a1a", font=("Segoe UI", 14, "bold"),
            anchor="w", padx=18,
        ).pack(fill="x", pady=(14, 4))

        tk.Label(
            parent,
            text="Convert any files into clean Markdown (.md) text — ideal for feeding into another AI.",
            bg=APP_BG, fg="#555", font=("Segoe UI", 9),
            anchor="w", padx=18, justify="left",
        ).pack(fill="x", pady=(0, 6))

        # Drop zone
        self._ai_drop_frame = tk.Frame(
            parent, bg=DROP_BG, highlightbackground=DROP_BD,
            highlightthickness=2, relief="flat",
        )
        self._ai_drop_frame.pack(fill="both", expand=True, padx=18, pady=(10, 6))
        self._ai_drop_label = tk.Label(
            self._ai_drop_frame,
            text="Drop any files here\nor click 'Add Files'",
            bg=DROP_BG, fg="#4a6fa5", font=("Segoe UI", 11), justify="center",
        )
        self._ai_drop_label.pack(expand=True, pady=20)
        list_frame = tk.Frame(self._ai_drop_frame, bg=DROP_BG)
        list_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        self._ai_file_list = tk.Listbox(
            list_frame, yscrollcommand=scrollbar.set, selectmode="extended",
            bg="#ffffff", fg="#1a1a1a", font=("Segoe UI", 9),
            relief="flat", bd=1, activestyle="none", highlightthickness=0,
        )
        scrollbar.config(command=self._ai_file_list.yview)
        scrollbar.pack(side="right", fill="y")
        self._ai_file_list.pack(fill="both", expand=True)
        if HAS_DND:
            for widget in (self._ai_drop_frame, self._ai_drop_label, list_frame, self._ai_file_list):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._ai_on_drop)

        # Add / Remove / Clear
        btn_frame = tk.Frame(parent, bg=APP_BG)
        btn_frame.pack(fill="x", padx=18, pady=(4, 4))
        for label, cmd in [
            ("Add Files",       self._ai_add_files),
            ("Remove Selected", self._ai_remove_selected),
            ("Clear All",       self._ai_clear_files),
        ]:
            tk.Button(
                btn_frame, text=label, command=cmd,
                bg="#e0e8f0", fg="#333", font=("Segoe UI", 9),
                relief="flat", padx=12, pady=4, cursor="hand2",
            ).pack(side="left", padx=(0, 6))

        # Output folder row
        out_frame = tk.Frame(parent, bg=APP_BG)
        out_frame.pack(fill="x", padx=18, pady=(2, 2))

        tk.Label(
            out_frame, text="Output folder:",
            bg=APP_BG, fg="#555", font=("Segoe UI", 9),
        ).pack(side="left")

        self._ai_out_var = tk.StringVar(value="Same as source file")
        tk.Label(
            out_frame, textvariable=self._ai_out_var,
            bg=APP_BG, fg=ACCENT, font=("Segoe UI", 9, "italic"),
        ).pack(side="left", padx=4)

        tk.Button(
            out_frame, text="Choose…", command=self._ai_choose_output,
            bg="#e0e8f0", fg="#333",
            font=("Segoe UI", 9), relief="flat",
            padx=8, pady=2, cursor="hand2",
        ).pack(side="left", padx=4)

        # Delete-originals option
        opt_frame = tk.Frame(parent, bg=APP_BG)
        opt_frame.pack(fill="x", padx=18, pady=(2, 2))
        self._ai_delete_orig = tk.BooleanVar(value=False)
        _ToggleSwitch(
            opt_frame, text="Delete original files after successful conversion",
            variable=self._ai_delete_orig,
        ).pack(side="left")

        # Progress + status
        self._ai_progress = ttk.Progressbar(parent, mode="determinate")
        self._ai_progress.pack(fill="x", padx=18, pady=(6, 2))
        self._ai_tracker = _ProgressTracker(self.root, self._ai_progress)
        self._ai_status_var = tk.StringVar(value="Ready — add files to convert to Markdown.")
        tk.Label(parent, textvariable=self._ai_status_var, bg=APP_BG, fg="#555",
                 font=("Segoe UI", 8), anchor="w").pack(fill="x", padx=20, pady=(0, 4))
        self._ai_timer = self._build_timer(parent)

        # Convert + Cancel
        ai_btn_row = tk.Frame(parent, bg=APP_BG)
        ai_btn_row.pack(pady=(4, 14))
        self._ai_btn = tk.Button(
            ai_btn_row, text="Convert to Markdown", command=self._start_aiprep,
            bg=ACCENT, fg=BTN_FG, font=("Segoe UI", 12, "bold"),
            relief="flat", padx=30, pady=8, cursor="hand2",
            activebackground="#005a9e", activeforeground=BTN_FG,
        )
        self._ai_btn.pack(side="left")
        self._ai_cancel_btn = tk.Button(
            ai_btn_row, text="Cancel", command=self._cancel_aiprep_op,
            font=("Segoe UI", 10, "bold"), relief="flat", padx=16, pady=6,
            activebackground=CANCEL_BG_ACTIVE, activeforeground=BTN_FG,
        )
        self._ai_cancel_btn.pack(side="left", padx=(10, 0))
        self._set_cancel_state(self._ai_cancel_btn, False)

    def _ai_add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select files to convert to Markdown",
            filetypes=[("All files", "*.*")],
        )
        for p in paths:
            self._ai_add_path(p)

    def _ai_on_drop(self, event):
        for p in self.root.tk.splitlist(event.data):
            self._ai_add_path(p)

    def _ai_add_path(self, p):
        path = Path(p)
        if not path.is_file():
            return   # AI Prep accepts any file type, but only files
        if path not in self._ai_files:
            self._ai_files.append(path)
            self._ai_file_list.insert("end", path.name)
        self._ai_update_drop_label()

    def _ai_remove_selected(self):
        for i in sorted(self._ai_file_list.curselection(), reverse=True):
            self._ai_file_list.delete(i)
            del self._ai_files[i]
        self._ai_update_drop_label()

    def _ai_clear_files(self):
        self._ai_files.clear()
        self._ai_file_list.delete(0, "end")
        self._ai_update_drop_label()
        # Return the page to its ready state (clear any "Done!"/error status).
        self._ai_status_var.set("Ready — add files to convert to Markdown.")
        self._ai_tracker.reset()

    def _ai_update_drop_label(self):
        if self._ai_files:
            self._ai_drop_label.config(text=f"{len(self._ai_files)} file(s) queued")
        else:
            self._ai_drop_label.config(text="Drop any files here\nor click 'Add Files'")

    def _ai_choose_output(self):
        d = filedialog.askdirectory(title="Select output folder")
        if d:
            self._ai_out_dir = Path(d)
            self._ai_out_var.set(str(self._ai_out_dir))
        else:
            self._ai_out_dir = None
            self._ai_out_var.set("Same as source file")

    def _cancel_aiprep_op(self):
        self._cancel_ai.set()
        self._set_cancel_state(self._ai_cancel_btn, False)
        self._ai_status_var.set("Cancelling… finishing the current file, then stopping.")

    def _start_aiprep(self):
        if not self._ai_files:
            messagebox.showwarning("No Files", "Please add files to convert first.")
            return
        self._cancel_ai.clear()
        self._ai_btn.config(state="disabled")
        self._set_cancel_state(self._ai_cancel_btn, True)
        self._ai_tracker.start(len(self._ai_files))
        self._ai_timer.start()
        threading.Thread(target=self._aiprep_worker, daemon=True).start()

    def _aiprep_worker(self):
        files = list(self._ai_files)
        delete_orig = self._ai_delete_orig.get()
        # Only ever delete files the user actually uploaded — never anything else
        # that may live in the output folder. See _delete_original_file.
        originals = _resolved_set(files)
        total = len(files)
        converted = 0
        failed = []
        protected = []

        # Ensure markitdown is available before looping; surface a clear error if not.
        try:
            _get_markitdown()
        except Exception as exc:
            msg = str(exc)
            self.root.after(0, lambda: (
                self._ai_btn.config(state="normal"),
                self._set_cancel_state(self._ai_cancel_btn, False),
                self._ai_timer.reset(),
                self._ai_status_var.set("markitdown is not installed."),
                messagebox.showerror("markitdown Not Available", msg),
            ))
            return

        for i, src in enumerate(files):
            if self._cancel_ai.is_set():
                break
            # Skip password-protected files — point the user at Password Removal.
            if _is_password_protected(src):
                protected.append(src.name)
                self._ai_tracker.complete_item()
                continue
            self.root.after(0, lambda s=src, n=i: self._ai_status_var.set(
                f"Converting: {s.name}  ({n + 1}/{total})"))
            self._ai_timer.set_file(src.name)
            self._ai_tracker.begin_item()
            try:
                out_dir = self._ai_out_dir if self._ai_out_dir else src.parent
                out = _to_markdown(src, out_dir)
                if out.exists() and out.stat().st_size > 0:
                    converted += 1
                    # For emails, save non-image attachments next to the .md and
                    # convert each of those to Markdown too.
                    if src.suffix.lower() in (".eml", ".msg"):
                        try:
                            n_att = _email_attachments_to_markdown(src, out)
                            if n_att:
                                self.root.after(0, lambda s=src, n=n_att: self._ai_status_var.set(
                                    f"{s.name}: saved {n} attachment(s) to a folder beside the .md"))
                        except Exception:
                            pass   # attachment handling is best-effort
                    if delete_orig:
                        _delete_original_file(src, originals)
                else:
                    failed.append(src.name)
            except Exception as exc:
                failed.append(f"{src.name}: {exc}")
            self._ai_tracker.complete_item()

        self.root.after(0, lambda: self._ai_finish(converted, failed, protected))

    def _ai_finish(self, converted, failed, protected=None):
        self._ai_timer.stop()
        self._ai_tracker.finish()
        elapsed = self._ai_timer.elapsed_str()
        self._ai_timer.reset()
        self._ai_btn.config(state="normal")
        self._set_cancel_state(self._ai_cancel_btn, False)
        protected = protected or []
        pw_note = _password_skip_note(protected)
        skip_status = f"  {len(protected)} skipped (password-protected)." if protected else ""
        if self._cancel_ai.is_set():
            self._ai_status_var.set(f"⏹  Cancelled. {converted} file(s) converted before stopping.")
            messagebox.showinfo(
                "AI Preparation Cancelled",
                f"Conversion cancelled.\n{converted} file(s) converted before stopping."
                f"\n\nTotal time elapsed: {elapsed}{pw_note}",
            )
            return
        if not failed:
            self._ai_status_var.set(f"✅  Done! {converted} file(s) converted to Markdown.{skip_status}")
            messagebox.showinfo(
                "AI Preparation Complete",
                f"{converted} file(s) converted to Markdown (.md)."
                f"\n\nTotal time elapsed: {elapsed}{pw_note}",
            )
        else:
            err_lines = "\n".join(f"• {n}" for n in failed)
            self._ai_status_var.set(f"⚠️  {converted} converted, {len(failed)} failed.{skip_status}")
            messagebox.showerror(
                "AI Preparation — Some Files Failed",
                f"{converted} converted, {len(failed)} failed:\n\n{err_lines}"
                f"\n\nTotal time elapsed: {elapsed}{pw_note}",
            )

    # ── Office Modernizer ────────────────────
    def _build_office_page(self, parent):
        tk.Label(
            parent, text="Office Modernizer",
            bg=APP_BG, fg="#1a1a1a", font=("Segoe UI", 14, "bold"),
            anchor="w", padx=18,
        ).pack(fill="x", pady=(14, 4))

        tk.Label(
            parent,
            text=("Batch-convert old Office files to their newest formats. Drop a mix of "
                  "types at once — each is converted one-to-one:\n"
                  ".doc/.dot → .docx/.dotx    .xls/.xlt → .xlsx/.xltx    .ppt/.pps/.pot → .pptx/.ppsx/.potx"),
            bg=APP_BG, fg="#555", font=("Segoe UI", 9),
            anchor="w", padx=18, justify="left",
        ).pack(fill="x", pady=(0, 6))

        # Drop zone
        self._office_drop_frame = tk.Frame(
            parent, bg=DROP_BG, highlightbackground=DROP_BD,
            highlightthickness=2, relief="flat",
        )
        self._office_drop_frame.pack(fill="both", expand=True, padx=18, pady=(10, 6))
        self._office_drop_label = tk.Label(
            self._office_drop_frame,
            text="Drop legacy Office files here\nor click 'Add Files'",
            bg=DROP_BG, fg="#4a6fa5", font=("Segoe UI", 11), justify="center",
        )
        self._office_drop_label.pack(expand=True, pady=20)
        list_frame = tk.Frame(self._office_drop_frame, bg=DROP_BG)
        list_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        self._office_file_list = tk.Listbox(
            list_frame, yscrollcommand=scrollbar.set, selectmode="extended",
            bg="#ffffff", fg="#1a1a1a", font=("Segoe UI", 9),
            relief="flat", bd=1, activestyle="none", highlightthickness=0,
        )
        scrollbar.config(command=self._office_file_list.yview)
        scrollbar.pack(side="right", fill="y")
        self._office_file_list.pack(fill="both", expand=True)
        if HAS_DND:
            for widget in (self._office_drop_frame, self._office_drop_label, list_frame, self._office_file_list):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._office_on_drop)

        # Add / Remove / Clear
        btn_frame = tk.Frame(parent, bg=APP_BG)
        btn_frame.pack(fill="x", padx=18, pady=(4, 4))
        for label, cmd in [
            ("Add Files",       self._office_add_files),
            ("Remove Selected", self._office_remove_selected),
            ("Clear All",       self._office_clear_files),
        ]:
            tk.Button(
                btn_frame, text=label, command=cmd,
                bg="#e0e8f0", fg="#333", font=("Segoe UI", 9),
                relief="flat", padx=12, pady=4, cursor="hand2",
            ).pack(side="left", padx=(0, 6))

        # Output folder row
        out_frame = tk.Frame(parent, bg=APP_BG)
        out_frame.pack(fill="x", padx=18, pady=(2, 2))
        tk.Label(
            out_frame, text="Output folder:",
            bg=APP_BG, fg="#555", font=("Segoe UI", 9),
        ).pack(side="left")
        self._office_out_var = tk.StringVar(value="Same as source file")
        tk.Label(
            out_frame, textvariable=self._office_out_var,
            bg=APP_BG, fg=ACCENT, font=("Segoe UI", 9, "italic"),
        ).pack(side="left", padx=4)
        tk.Button(
            out_frame, text="Choose…", command=self._office_choose_output,
            bg="#e0e8f0", fg="#333", font=("Segoe UI", 9), relief="flat",
            padx=8, pady=2, cursor="hand2",
        ).pack(side="left", padx=4)

        # Delete-originals option
        opt_frame = tk.Frame(parent, bg=APP_BG)
        opt_frame.pack(fill="x", padx=18, pady=(2, 2))
        self._office_delete_orig = tk.BooleanVar(value=False)
        _ToggleSwitch(
            opt_frame, text="Delete original files after successful conversion",
            variable=self._office_delete_orig,
        ).pack(side="left")

        # Progress + status
        self._office_progress = ttk.Progressbar(parent, mode="determinate")
        self._office_progress.pack(fill="x", padx=18, pady=(6, 2))
        self._office_tracker = _ProgressTracker(self.root, self._office_progress)
        self._office_status_var = tk.StringVar(
            value="Ready — add legacy Office files to modernize.")
        tk.Label(parent, textvariable=self._office_status_var, bg=APP_BG, fg="#555",
                 font=("Segoe UI", 8), anchor="w").pack(fill="x", padx=20, pady=(0, 4))
        self._office_timer = self._build_timer(parent)

        # Convert + Cancel
        office_btn_row = tk.Frame(parent, bg=APP_BG)
        office_btn_row.pack(pady=(4, 14))
        self._office_btn = tk.Button(
            office_btn_row, text="Modernize Files", command=self._start_office,
            bg=ACCENT, fg=BTN_FG, font=("Segoe UI", 12, "bold"),
            relief="flat", padx=30, pady=8, cursor="hand2",
            activebackground="#005a9e", activeforeground=BTN_FG,
        )
        self._office_btn.pack(side="left")
        self._office_cancel_btn = tk.Button(
            office_btn_row, text="Cancel", command=self._cancel_office_op,
            font=("Segoe UI", 10, "bold"), relief="flat", padx=16, pady=6,
            activebackground=CANCEL_BG_ACTIVE, activeforeground=BTN_FG,
        )
        self._office_cancel_btn.pack(side="left", padx=(10, 0))
        self._set_cancel_state(self._office_cancel_btn, False)

    def _office_add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select legacy Office files",
            filetypes=[("Legacy Office files", "*.doc *.dot *.xls *.xlt *.ppt *.pps *.pot"),
                       ("All files", "*.*")],
        )
        for p in paths:
            self._office_add_path(p)

    def _office_on_drop(self, event):
        for p in self.root.tk.splitlist(event.data):
            self._office_add_path(p)

    def _office_add_path(self, p):
        path = Path(p)
        if not path.is_file():
            return
        if path.suffix.lower() not in LEGACY_OFFICE:
            return   # only accept legacy Office files; ignore anything else
        if path not in self._office_files:
            self._office_files.append(path)
            self._office_file_list.insert("end", path.name)
        self._office_update_drop_label()

    def _office_remove_selected(self):
        for i in sorted(self._office_file_list.curselection(), reverse=True):
            self._office_file_list.delete(i)
            del self._office_files[i]
        self._office_update_drop_label()

    def _office_clear_files(self):
        self._office_files.clear()
        self._office_file_list.delete(0, "end")
        self._office_update_drop_label()
        # Return the page to its ready state (clear any "Done!"/error status).
        self._office_status_var.set("Ready — add legacy Office files to modernize.")
        self._office_tracker.reset()

    def _office_update_drop_label(self):
        if self._office_files:
            self._office_drop_label.config(text=f"{len(self._office_files)} file(s) queued")
        else:
            self._office_drop_label.config(
                text="Drop legacy Office files here\nor click 'Add Files'")

    def _office_choose_output(self):
        d = filedialog.askdirectory(title="Select output folder")
        if d:
            self._office_out_dir = Path(d)
            self._office_out_var.set(str(self._office_out_dir))
        else:
            self._office_out_dir = None
            self._office_out_var.set("Same as source file")

    def _cancel_office_op(self):
        # True kill: force-close the dedicated Office instance(s) WE launched so
        # a stuck/hung conversion is aborted immediately. Only our own tracked
        # PIDs are terminated — the user's open Office windows are never touched.
        self._cancel_office.set()
        self._set_cancel_state(self._office_cancel_btn, False)
        self._office_status_var.set("Stopping… force-closing the current conversion.")
        for pids in list(getattr(self, "_office_app_pids", {}).values()):
            for pid in list(pids):
                _terminate_pid(pid)

    def _start_office(self):
        if not self._office_files:
            messagebox.showwarning("No Files", "Please add legacy Office files first.")
            return
        self._cancel_office.clear()
        self._office_btn.config(state="disabled")
        self._set_cancel_state(self._office_cancel_btn, True)
        self._office_tracker.start(len(self._office_files))
        self._office_timer.start()
        threading.Thread(target=self._office_worker, daemon=True).start()

    def _office_worker(self):
        files = list(self._office_files)
        delete_orig = self._office_delete_orig.get()
        # Only ever delete files the user actually uploaded — never anything else
        # that may live in the output folder. See _delete_original_file.
        originals = _resolved_set(files)
        total = len(files)
        converted = 0
        failed = []
        protected = []

        pythoncom.CoInitialize()
        # Each family gets its own DEDICATED Office instance (launched lazily,
        # only if a matching file is queued). We track the PIDs of just those
        # instances so Cancel can force-kill a stuck conversion without ever
        # harming the user's own open Office windows.
        apps = {"word": None, "excel": None, "powerpoint": None}
        ensured = {"word": False, "excel": False, "powerpoint": False}
        self._office_app_pids = {}
        try:
            for i, src in enumerate(files):
                if self._cancel_office.is_set():
                    break
                # Skip password-protected files — point the user at Password Removal.
                if _is_password_protected(src):
                    protected.append(src.name)
                    self._office_tracker.complete_item()
                    continue
                self._set_office_status(f"Converting: {src.name}  ({i + 1}/{total})")
                self._office_timer.set_file(src.name)
                self._office_tracker.begin_item()
                info = LEGACY_OFFICE.get(src.suffix.lower())
                if info is None:
                    failed.append((src.name, "unsupported file type"))
                    self._office_tracker.complete_item()
                    continue
                family, out_ext, fmt = info
                if not ensured[family]:
                    try:
                        app_obj, pids = _launch_office_isolated(family)
                    except Exception:
                        app_obj, pids = None, set()
                    apps[family] = app_obj
                    self._office_app_pids[family] = set(pids)
                    ensured[family] = True

                out_dir = self._office_out_dir if self._office_out_dir else src.parent
                out_path = _unique_output_path(out_dir, src.stem, out_ext)

                # Bound this file: if it isn't done within OFFICE_FILE_TIMEOUT
                # its dedicated instance is assumed stuck on a hidden dialog and
                # force-killed, so the batch can skip it and continue.
                wd = _FileWatchdog(self._office_app_pids.get(family))
                try:
                    with wd:
                        _modernize_office_file(src, out_path, family, fmt, apps[family])
                    converted += 1
                    if delete_orig:
                        _delete_original_file(src, originals)
                except Exception as exc:
                    # Remove any partial/corrupt output left behind — e.g. when
                    # the conversion was force-killed mid-save. The path is
                    # uniquely named, so this only ever deletes our own output.
                    try:
                        if out_path.exists():
                            out_path.unlink()
                    except Exception:
                        pass
                    if self._cancel_office.is_set():
                        pass                       # user cancelled — not a failure
                    elif wd.fired:
                        failed.append((src.name, "timed out — appeared stuck, skipped"))
                    else:
                        failed.append((src.name, str(exc)))

                # If the watchdog killed the instance, drop it so the next file
                # of this family gets a fresh, healthy one.
                if wd.fired:
                    apps[family] = None
                    ensured[family] = False
                    self._office_app_pids.pop(family, None)

                self._office_tracker.complete_item()
        finally:
            # Close our dedicated instances (already-killed ones error here —
            # ignored), then make sure none of our launched processes linger.
            for app in apps.values():
                if app is not None:
                    try:
                        app.Quit()
                    except Exception:
                        pass
            for pids in list(self._office_app_pids.values()):
                for pid in pids:
                    _terminate_pid(pid)
            self._office_app_pids = {}
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

        self.root.after(0, lambda: self._office_finish(converted, failed, protected))

    def _set_office_status(self, msg):
        self.root.after(0, lambda: self._office_status_var.set(msg))

    def _office_finish(self, converted, failed, protected=None):
        self._office_timer.stop()
        self._office_tracker.finish()
        elapsed = self._office_timer.elapsed_str()
        self._office_timer.reset()
        self._office_btn.config(state="normal")
        self._set_cancel_state(self._office_cancel_btn, False)
        protected = protected or []
        pw_note = _password_skip_note(protected)
        skip_status = f"  {len(protected)} skipped (password-protected)." if protected else ""
        if self._cancel_office.is_set():
            self._office_status_var.set(
                f"⏹  Cancelled. {converted} file(s) modernized before stopping.")
            messagebox.showinfo(
                "Office Modernizer Cancelled",
                f"Conversion cancelled.\n{converted} file(s) modernized before stopping."
                f"\n\nTotal time elapsed: {elapsed}{pw_note}",
            )
            return
        if not failed:
            self._office_status_var.set(f"✅  Done! {converted} file(s) modernized.{skip_status}")
            messagebox.showinfo(
                "Office Modernizer Complete",
                f"{converted} file(s) converted to their modern Office formats."
                f"\n\nTotal time elapsed: {elapsed}{pw_note}",
            )
        else:
            err_lines = "\n".join(f"• {n}: {e}" for n, e in failed)
            self._office_status_var.set(f"⚠️  {converted} modernized, {len(failed)} failed.{skip_status}")
            messagebox.showerror(
                "Office Modernizer — Some Files Failed",
                f"{converted} modernized, {len(failed)} failed:\n\n{err_lines}"
                f"\n\nTotal time elapsed: {elapsed}{pw_note}",
            )

    # ── Password Removal page ────────────────
    def _build_pwd_page(self, parent):
        tk.Label(
            parent, text="Password Removal",
            bg=APP_BG, fg="#1a1a1a", font=("Segoe UI", 14, "bold"),
            anchor="w", padx=18,
        ).pack(fill="x", pady=(14, 4))

        tk.Label(
            parent,
            text=("Remove the password from protected Office files. You'll be asked for "
                  "each file's password in turn; an unlocked copy is saved with a "
                  "'_nopass' suffix.\n"
                  "Works with Word, Excel and PowerPoint (.docx, .xlsx, .pptx and older/"
                  "template variants)."),
            bg=APP_BG, fg="#555", font=("Segoe UI", 9),
            anchor="w", padx=18, justify="left", wraplength=620,
        ).pack(fill="x", pady=(0, 6))

        # Drop zone
        self._pwd_drop_frame = tk.Frame(
            parent, bg=DROP_BG, highlightbackground=DROP_BD,
            highlightthickness=2, relief="flat",
        )
        self._pwd_drop_frame.pack(fill="both", expand=True, padx=18, pady=(10, 6))
        self._pwd_drop_label = tk.Label(
            self._pwd_drop_frame,
            text="Drop password-protected Office files here\nor click 'Add Files'",
            bg=DROP_BG, fg="#4a6fa5", font=("Segoe UI", 11), justify="center",
        )
        self._pwd_drop_label.pack(expand=True, pady=20)
        list_frame = tk.Frame(self._pwd_drop_frame, bg=DROP_BG)
        list_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        self._pwd_file_list = tk.Listbox(
            list_frame, yscrollcommand=scrollbar.set, selectmode="extended",
            bg="#ffffff", fg="#1a1a1a", font=("Segoe UI", 9),
            relief="flat", bd=1, activestyle="none", highlightthickness=0,
        )
        scrollbar.config(command=self._pwd_file_list.yview)
        scrollbar.pack(side="right", fill="y")
        self._pwd_file_list.pack(fill="both", expand=True)
        if HAS_DND:
            for widget in (self._pwd_drop_frame, self._pwd_drop_label, list_frame, self._pwd_file_list):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._pwd_on_drop)

        # Add / Remove / Clear
        btn_frame = tk.Frame(parent, bg=APP_BG)
        btn_frame.pack(fill="x", padx=18, pady=(4, 4))
        for label, cmd in [
            ("Add Files",       self._pwd_add_files),
            ("Remove Selected", self._pwd_remove_selected),
            ("Clear All",       self._pwd_clear_files),
        ]:
            tk.Button(
                btn_frame, text=label, command=cmd,
                bg="#e0e8f0", fg="#333", font=("Segoe UI", 9),
                relief="flat", padx=12, pady=4, cursor="hand2",
            ).pack(side="left", padx=(0, 6))

        # Output folder row
        out_frame = tk.Frame(parent, bg=APP_BG)
        out_frame.pack(fill="x", padx=18, pady=(2, 2))
        tk.Label(
            out_frame, text="Output folder:",
            bg=APP_BG, fg="#555", font=("Segoe UI", 9),
        ).pack(side="left")
        self._pwd_out_var = tk.StringVar(value="Same as source file")
        tk.Label(
            out_frame, textvariable=self._pwd_out_var,
            bg=APP_BG, fg=ACCENT, font=("Segoe UI", 9, "italic"),
        ).pack(side="left", padx=4)
        tk.Button(
            out_frame, text="Choose…", command=self._pwd_choose_output,
            bg="#e0e8f0", fg="#333", font=("Segoe UI", 9), relief="flat",
            padx=8, pady=2, cursor="hand2",
        ).pack(side="left", padx=4)

        # Progress + status
        self._pwd_progress = ttk.Progressbar(parent, mode="determinate")
        self._pwd_progress.pack(fill="x", padx=18, pady=(6, 2))
        self._pwd_tracker = _ProgressTracker(self.root, self._pwd_progress)
        self._pwd_status_var = tk.StringVar(
            value="Ready — add password-protected Office files to unlock.")
        tk.Label(parent, textvariable=self._pwd_status_var, bg=APP_BG, fg="#555",
                 font=("Segoe UI", 8), anchor="w").pack(fill="x", padx=20, pady=(0, 4))
        self._pwd_timer = self._build_timer(parent)

        # Remove + Cancel
        pwd_btn_row = tk.Frame(parent, bg=APP_BG)
        pwd_btn_row.pack(pady=(4, 14))
        self._pwd_btn = tk.Button(
            pwd_btn_row, text="Remove Passwords", command=self._start_pwd,
            bg=ACCENT, fg=BTN_FG, font=("Segoe UI", 12, "bold"),
            relief="flat", padx=30, pady=8, cursor="hand2",
            activebackground="#005a9e", activeforeground=BTN_FG,
        )
        self._pwd_btn.pack(side="left")
        self._pwd_cancel_btn = tk.Button(
            pwd_btn_row, text="Cancel", command=self._cancel_pwd_op,
            font=("Segoe UI", 10, "bold"), relief="flat", padx=16, pady=6,
            activebackground=CANCEL_BG_ACTIVE, activeforeground=BTN_FG,
        )
        self._pwd_cancel_btn.pack(side="left", padx=(10, 0))
        self._set_cancel_state(self._pwd_cancel_btn, False)

    def _pwd_add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select password-protected Office files",
            filetypes=[("Office files",
                        "*.doc *.docx *.docm *.dot *.dotx *.dotm *.rtf "
                        "*.xls *.xlsx *.xlsm *.xlsb *.xlt *.xltx *.xltm "
                        "*.ppt *.pptx *.pptm *.pps *.ppsx *.pot *.potx"),
                       ("All files", "*.*")],
        )
        for p in paths:
            self._pwd_add_path(p)

    def _pwd_on_drop(self, event):
        for p in self.root.tk.splitlist(event.data):
            self._pwd_add_path(p)

    def _pwd_add_path(self, p):
        path = Path(p)
        if not path.is_file():
            return
        if path.suffix.lower() not in PWD_EXTS:
            return   # only accept supported Office files; ignore anything else
        if path not in self._pwd_files:
            self._pwd_files.append(path)
            self._pwd_file_list.insert("end", path.name)
        self._pwd_update_drop_label()

    def _pwd_remove_selected(self):
        for i in sorted(self._pwd_file_list.curselection(), reverse=True):
            self._pwd_file_list.delete(i)
            del self._pwd_files[i]
        self._pwd_update_drop_label()

    def _pwd_clear_files(self):
        self._pwd_files.clear()
        self._pwd_file_list.delete(0, "end")
        self._pwd_update_drop_label()
        # Return the page to its ready state (clear any "Done!"/error status).
        self._pwd_status_var.set("Ready — add password-protected Office files to unlock.")
        self._pwd_tracker.reset()

    def _pwd_update_drop_label(self):
        if self._pwd_files:
            self._pwd_drop_label.config(text=f"{len(self._pwd_files)} file(s) queued")
        else:
            self._pwd_drop_label.config(
                text="Drop password-protected Office files here\nor click 'Add Files'")

    def _pwd_choose_output(self):
        d = filedialog.askdirectory(title="Select output folder")
        if d:
            self._pwd_out_dir = Path(d)
            self._pwd_out_var.set(str(self._pwd_out_dir))
        else:
            self._pwd_out_dir = None
            self._pwd_out_var.set("Same as source file")

    def _cancel_pwd_op(self):
        # Stop the batch. Decryption is fast pure-Python work, so the loop simply
        # exits between files; there are no Office processes to kill.
        self._cancel_pwd.set()
        self._set_cancel_state(self._pwd_cancel_btn, False)
        self._pwd_status_var.set("Stopping…")
        # If a password prompt is open, close it so the worker unblocks and the
        # batch can stop (an un-answered prompt would otherwise hold it).
        dlg = getattr(self, "_pwd_dialog", None)
        if dlg is not None:
            try:
                dlg.destroy()
            except Exception:
                pass

    def _start_pwd(self):
        if not self._pwd_files:
            messagebox.showwarning("No Files", "Please add protected Office files first.")
            return
        self._cancel_pwd.clear()
        self._pwd_btn.config(state="disabled")
        self._set_cancel_state(self._pwd_cancel_btn, True)
        self._pwd_tracker.start(len(self._pwd_files))
        self._pwd_timer.start()
        threading.Thread(target=self._pwd_worker, daemon=True).start()

    def _run_on_main(self, fn):
        """Run `fn` on the Tk main thread and block the calling (worker) thread
        until it has finished, returning whatever `fn` returned. Used so the
        worker can put up modal dialogs (which must live on the main thread)."""
        box = {"val": None}
        done = threading.Event()

        def _runner():
            try:
                box["val"] = fn()
            finally:
                done.set()

        self.root.after(0, _runner)
        done.wait()
        return box["val"]

    def _prompt_password(self, filename, note=""):
        """MAIN-THREAD ONLY. Modal dialog asking for `filename`'s password.
        Returns the entered password, or None if the user chose to skip the file.
        `note` shows an optional message (e.g. an 'incorrect password' reminder)."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Password Required")
        dlg.configure(bg=APP_BG)
        dlg.transient(self.root)
        dlg.resizable(False, False)
        # Build the window hidden so we can place it BEFORE it's first mapped —
        # otherwise the window manager drops it at the primary-screen corner
        # (very noticeable on a multi-monitor setup) and ignores a late move.
        dlg.withdraw()
        result = {"pwd": None}     # None == skip

        tk.Label(dlg, text="Enter the password for this document:",
                 bg=APP_BG, fg="#333", font=("Segoe UI", 9)).pack(
            anchor="w", padx=16, pady=(14, 0))
        tk.Label(dlg, text=filename, bg=APP_BG, fg=ACCENT,
                 font=("Segoe UI", 10, "bold"), wraplength=380, justify="left").pack(
            anchor="w", padx=16, pady=(2, 6))
        if note:
            tk.Label(dlg, text=note, bg=APP_BG, fg=CANCEL_BG_ACTIVE,
                     font=("Segoe UI", 8)).pack(anchor="w", padx=16)

        var = tk.StringVar()
        show_var = tk.BooleanVar(value=False)
        entry = tk.Entry(dlg, textvariable=var, show="•",
                         font=("Segoe UI", 10), width=38)
        entry.pack(padx=16, pady=(6, 2))

        def _toggle():
            entry.config(show="" if show_var.get() else "•")
        _ToggleSwitch(dlg, text="Show password", variable=show_var,
                      command=_toggle, fg="#555", font=("Segoe UI", 8)).pack(
            anchor="w", padx=14, pady=(2, 0))

        def _submit():
            result["pwd"] = var.get()
            dlg.destroy()

        def _skip():
            result["pwd"] = None
            dlg.destroy()

        row = tk.Frame(dlg, bg=APP_BG)
        row.pack(pady=(8, 14))
        tk.Button(row, text="Unlock", command=_submit, bg=ACCENT, fg=BTN_FG,
                  font=("Segoe UI", 10, "bold"), relief="flat", padx=18, pady=5,
                  cursor="hand2", activebackground="#005a9e",
                  activeforeground=BTN_FG).pack(side="left", padx=(0, 8))
        tk.Button(row, text="Skip File", command=_skip, bg="#e0e8f0", fg="#333",
                  font=("Segoe UI", 10), relief="flat", padx=14, pady=5,
                  cursor="hand2").pack(side="left")

        entry.bind("<Return>", lambda e: _submit())
        entry.bind("<Escape>", lambda e: _skip())
        dlg.protocol("WM_DELETE_WINDOW", _skip)

        # Center over the main window. Use the *requested* size (reliable while
        # the window is still hidden) and set a full WxH+X+Y geometry, then show
        # it — so it maps directly at the centered position.
        dlg.update_idletasks()
        w, h = dlg.winfo_reqwidth(), dlg.winfo_reqheight()
        rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
        rw, rh = self.root.winfo_width(), self.root.winfo_height()
        x = rx + (rw - w) // 2
        y = ry + (rh - h) // 2
        dlg.geometry(f"{w}x{h}+{x}+{y}")
        dlg.deiconify()

        # Keep the prompt in front and focused, but DON'T grab_set — a hard modal
        # grab would block the main-window Cancel button. Instead we expose the
        # dialog so Cancel can dismiss it (treated as a skip) to stop the batch.
        try:
            dlg.attributes("-topmost", True)
        except Exception:
            pass
        self._pwd_dialog = dlg
        entry.focus_force()
        self.root.wait_window(dlg)
        self._pwd_dialog = None
        return result["pwd"]

    def _pwd_worker(self):
        files = list(self._pwd_files)
        total = len(files)
        unlocked = 0
        failed = []     # (name, reason)
        skipped = []    # names

        # Decryption is pure Python (msoffcrypto) — no Office, no COM, no hidden
        # prompts — so each file is handled in a simple, interruptible loop.
        for i, src in enumerate(files):
            if self._cancel_pwd.is_set():
                break
            if src.suffix.lower() not in PWD_EXTS:
                failed.append((src.name, "unsupported file type"))
                self._pwd_tracker.complete_item()
                continue

            # Skip files that aren't actually encrypted — no point prompting.
            if not _office_is_encrypted(src):
                failed.append((src.name, "not password-protected (no open-password encryption)"))
                self._pwd_tracker.complete_item()
                continue

            # Ask for this file's password, re-prompting on a wrong one until it
            # succeeds or the user skips it.
            note = ""
            done_with_file = False
            while not self._cancel_pwd.is_set():
                self._set_pwd_status(
                    f"Waiting for password: {src.name}  ({i + 1}/{total})")
                pwd = self._run_on_main(
                    lambda n=note: self._prompt_password(src.name, n))
                if pwd is None:
                    skipped.append(src.name)
                    break   # user skipped this file

                self._set_pwd_status(f"Unlocking: {src.name}  ({i + 1}/{total})")
                self._pwd_timer.set_file(src.name)
                self._pwd_tracker.begin_item()
                out_dir = self._pwd_out_dir if self._pwd_out_dir else src.parent
                out_path = _unique_output_path(out_dir, src.stem + "_nopass", src.suffix)
                try:
                    _strip_office_password(src, out_path, pwd)
                    unlocked += 1
                    done_with_file = True
                except _BadPassword:
                    # Explicit "incorrect password" dialog, then loop to re-prompt.
                    self._run_on_main(lambda n=src.name: messagebox.showerror(
                        "Incorrect Password",
                        f"The password entered for \"{n}\" was incorrect.\n\n"
                        "Please try again, or choose Skip File to move on."))
                    note = "Incorrect password — please try again."
                    continue
                except _NotEncrypted:
                    failed.append((src.name, "not password-protected"))
                    done_with_file = True
                except Exception as exc:
                    failed.append((src.name, str(exc)))
                    done_with_file = True

                if done_with_file:
                    break

            self._pwd_tracker.complete_item()

        self.root.after(0, lambda: self._pwd_finish(unlocked, failed, skipped))

    def _set_pwd_status(self, msg):
        self.root.after(0, lambda: self._pwd_status_var.set(msg))

    def _pwd_finish(self, unlocked, failed, skipped):
        self._pwd_timer.stop()
        self._pwd_tracker.finish()
        elapsed = self._pwd_timer.elapsed_str()
        self._pwd_timer.reset()
        self._pwd_btn.config(state="normal")
        self._set_cancel_state(self._pwd_cancel_btn, False)

        extra = ""
        if skipped:
            extra += f"\nSkipped: {len(skipped)} file(s)."

        if self._cancel_pwd.is_set():
            self._pwd_status_var.set(
                f"⏹  Cancelled. {unlocked} file(s) unlocked before stopping.")
            messagebox.showinfo(
                "Password Removal Cancelled",
                f"Operation cancelled.\n{unlocked} file(s) unlocked before stopping."
                f"{extra}\n\nTotal time elapsed: {elapsed}",
            )
            return
        if not failed:
            self._pwd_status_var.set(f"✅  Done! {unlocked} file(s) unlocked.")
            messagebox.showinfo(
                "Password Removal Complete",
                f"{unlocked} file(s) saved without password protection (\"_nopass\")."
                f"{extra}\n\nTotal time elapsed: {elapsed}",
            )
        else:
            err_lines = "\n".join(f"• {n}: {e}" for n, e in failed)
            self._pwd_status_var.set(f"⚠️  {unlocked} unlocked, {len(failed)} failed.")
            messagebox.showerror(
                "Password Removal — Some Files Failed",
                f"{unlocked} unlocked, {len(failed)} failed:{extra}\n\n{err_lines}"
                f"\n\nTotal time elapsed: {elapsed}",
            )

    # ── Document Structuring page ──────────────────────
    def _build_structure_page(self, parent):
        tk.Label(
            parent, text="Document Structuring",
            bg=APP_BG, fg="#1a1a1a", font=("Segoe UI", 14, "bold"),
            anchor="w", padx=18,
        ).pack(fill="x", pady=(14, 4))

        tk.Label(
            parent,
            text=("Add ZIP archives. Each one is extracted first (including any nested "
                  "zips), then the extracted documents are standardized and organized: "
                  "optionally renamed to AI-suggested professional names and/or sorted "
                  "into folders by file type. Originals are never modified, and any "
                  "files already in the output folder are left untouched — only the "
                  "extracted documents are saved (and organized) there."),
            bg=APP_BG, fg="#555", font=("Segoe UI", 9),
            anchor="w", padx=18, justify="left", wraplength=620,
        ).pack(fill="x", pady=(0, 6))

        # Drop zone
        self._ds_drop_frame = tk.Frame(
            parent, bg=DROP_BG, highlightbackground=DROP_BD,
            highlightthickness=2, relief="flat",
        )
        self._ds_drop_frame.pack(fill="both", expand=True, padx=18, pady=(10, 6))
        self._ds_drop_label = tk.Label(
            self._ds_drop_frame,
            text="Drop .zip files here\nor click 'Add Zip Files'",
            bg=DROP_BG, fg="#4a6fa5", font=("Segoe UI", 11), justify="center",
        )
        self._ds_drop_label.pack(expand=True, pady=20)
        list_frame = tk.Frame(self._ds_drop_frame, bg=DROP_BG)
        list_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        self._ds_file_list = tk.Listbox(
            list_frame, yscrollcommand=scrollbar.set, selectmode="extended",
            bg="#ffffff", fg="#1a1a1a", font=("Segoe UI", 9),
            relief="flat", bd=1, activestyle="none", highlightthickness=0,
        )
        scrollbar.config(command=self._ds_file_list.yview)
        scrollbar.pack(side="right", fill="y")
        self._ds_file_list.pack(fill="both", expand=True)
        if HAS_DND:
            for widget in (self._ds_drop_frame, self._ds_drop_label, list_frame, self._ds_file_list):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._ds_on_drop)

        # Add / Remove / Clear
        btn_frame = tk.Frame(parent, bg=APP_BG)
        btn_frame.pack(fill="x", padx=18, pady=(4, 4))
        for label, cmd in [
            ("Add Zip Files",   self._ds_add_files),
            ("Remove Selected", self._ds_remove_selected),
            ("Clear All",       self._ds_clear_files),
        ]:
            tk.Button(
                btn_frame, text=label, command=cmd,
                bg="#e0e8f0", fg="#333", font=("Segoe UI", 9),
                relief="flat", padx=12, pady=4, cursor="hand2",
            ).pack(side="left", padx=(0, 6))

        # Options. The AI-naming checkbox stays disabled until the AI backend
        # is live (placeholders replaced) — the feature is a shell for now.
        opt_frame = tk.Frame(parent, bg=APP_BG)
        opt_frame.pack(fill="x", padx=18, pady=(2, 0))
        ai_ready = _ai_backend_ready()
        if not ai_ready:
            self._ds_rename.set(False)
            self._ds_date.set(False)
        self._ds_rename_check = _ToggleSwitch(
            opt_frame,
            text="AI filename suggestions" + ("" if ai_ready else "  (coming soon)"),
            variable=self._ds_rename,
            state=("normal" if ai_ready else "disabled"),
        )
        self._ds_rename_check.pack(anchor="w", pady=2)
        self._ds_date_check = _ToggleSwitch(
            opt_frame,
            text="Add the document's inferred date as a filename prefix (Date_Name)"
                 + ("" if ai_ready else "  (coming soon)"),
            variable=self._ds_date,
            state=("normal" if ai_ready else "disabled"),
        )
        self._ds_date_check.pack(anchor="w", pady=2)
        _ToggleSwitch(
            opt_frame, text="Organize into folders by file type",
            variable=self._ds_organize,
        ).pack(anchor="w", pady=2)

        # Output folder row (optional — defaults to each zip's own folder). The
        # folder need NOT be empty: only the extracted files are written/organized
        # here, existing files are never moved, swept into subfolders, or overwritten.
        out_frame = tk.Frame(parent, bg=APP_BG)
        out_frame.pack(fill="x", padx=18, pady=(4, 2))
        tk.Label(
            out_frame, text="Output folder:",
            bg=APP_BG, fg="#555", font=("Segoe UI", 9),
        ).pack(side="left")
        self._ds_out_var = tk.StringVar(value="Same as source file")
        tk.Label(
            out_frame, textvariable=self._ds_out_var,
            bg=APP_BG, fg=ACCENT, font=("Segoe UI", 9, "italic"),
        ).pack(side="left", padx=4)
        tk.Button(
            out_frame, text="Choose…", command=self._ds_choose_output,
            bg="#e0e8f0", fg="#333", font=("Segoe UI", 9), relief="flat",
            padx=8, pady=2, cursor="hand2",
        ).pack(side="left", padx=4)

        # Progress + status
        self._ds_progress = ttk.Progressbar(parent, mode="determinate")
        self._ds_progress.pack(fill="x", padx=18, pady=(6, 2))
        self._ds_tracker = _ProgressTracker(self.root, self._ds_progress)
        self._ds_status_var = tk.StringVar(
            value="Ready — add .zip files to extract and organize.")
        tk.Label(parent, textvariable=self._ds_status_var, bg=APP_BG, fg="#555",
                 font=("Segoe UI", 8), anchor="w").pack(fill="x", padx=20, pady=(0, 4))
        self._ds_timer = self._build_timer(parent)

        # Run + Cancel
        ds_btn_row = tk.Frame(parent, bg=APP_BG)
        ds_btn_row.pack(pady=(4, 14))
        self._ds_btn = tk.Button(
            ds_btn_row, text="Structure Documents", command=self._start_ds,
            bg=ACCENT, fg=BTN_FG, font=("Segoe UI", 12, "bold"),
            relief="flat", padx=30, pady=8, cursor="hand2",
            activebackground="#005a9e", activeforeground=BTN_FG,
        )
        self._ds_btn.pack(side="left")
        self._ds_cancel_btn = tk.Button(
            ds_btn_row, text="Cancel", command=self._cancel_ds_op,
            font=("Segoe UI", 10, "bold"), relief="flat", padx=16, pady=6,
            activebackground=CANCEL_BG_ACTIVE, activeforeground=BTN_FG,
        )
        self._ds_cancel_btn.pack(side="left", padx=(10, 0))
        self._set_cancel_state(self._ds_cancel_btn, False)

    def _ds_add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select ZIP files",
            filetypes=[("ZIP archives", "*.zip"), ("All files", "*.*")],
        )
        for p in paths:
            self._ds_add_path(p)

    def _ds_on_drop(self, event):
        for p in self.root.tk.splitlist(event.data):
            self._ds_add_path(p)

    def _ds_add_path(self, p):
        path = Path(p)
        if not path.is_file() or path.suffix.lower() != ".zip":
            return   # only accept .zip archives; ignore anything else
        if path not in self._ds_files:
            self._ds_files.append(path)
            self._ds_file_list.insert("end", path.name)
        self._ds_update_drop_label()

    def _ds_remove_selected(self):
        for i in sorted(self._ds_file_list.curselection(), reverse=True):
            self._ds_file_list.delete(i)
            del self._ds_files[i]
        self._ds_update_drop_label()

    def _ds_clear_files(self):
        self._ds_files.clear()
        self._ds_file_list.delete(0, "end")
        self._ds_update_drop_label()
        # Return the page to its ready state (clear any "Done!"/error status).
        self._ds_status_var.set("Ready — add .zip files to extract and organize.")
        self._ds_tracker.reset()

    def _ds_update_drop_label(self):
        if self._ds_files:
            self._ds_drop_label.config(text=f"{len(self._ds_files)} zip file(s) queued")
        else:
            self._ds_drop_label.config(text="Drop .zip files here\nor click 'Add Zip Files'")

    def _ds_choose_output(self):
        d = filedialog.askdirectory(title="Select output folder")
        if d:
            self._ds_out_dir = Path(d)
            self._ds_out_var.set(str(self._ds_out_dir))
        else:
            self._ds_out_dir = None
            self._ds_out_var.set("Same as source file")

    def _cancel_ds_op(self):
        # Stop the batch. Work is copies + short network calls, so the loop
        # exits between files; there are no Office processes to kill.
        self._cancel_ds.set()
        self._set_cancel_state(self._ds_cancel_btn, False)
        self._ds_status_var.set("Stopping…")

    def _start_ds(self):
        if not self._ds_files:
            messagebox.showwarning("No Files", "Please add ZIP files first.")
            return
        do_rename = self._ds_rename.get()
        do_date = self._ds_date.get()
        do_organize = self._ds_organize.get()

        cfg = None
        if do_rename or do_date:
            # Hard gate: the AI backend is unavailable unless its credentials are
            # present in the environment (never baked into this build). The
            # checkboxes are disabled too; this guard is belt-and-braces.
            if not _ai_backend_ready():
                messagebox.showinfo(
                    "AI Features Not Available Yet",
                    "The AI features (filename suggestions and date prefixing) "
                    "aren't available in this version yet. Organizing into folders "
                    "by file type still works.")
                return
            cfg = _load_ai_backend()

        # Snapshot the job on the main thread — the worker must not read Tk vars.
        self._ds_job = {
            "rename": do_rename,
            "date": do_date,
            "organize": do_organize,
            "cfg": cfg,
            "out_dir": self._ds_out_dir,
        }
        self._cancel_ds.clear()
        self._ds_btn.config(state="disabled")
        self._set_cancel_state(self._ds_cancel_btn, True)
        self._ds_tracker.start(len(self._ds_files))
        self._ds_timer.start()
        threading.Thread(target=self._ds_worker, daemon=True).start()

    def _ds_worker(self):
        job = self._ds_job
        do_rename, do_date, do_organize = job["rename"], job["date"], job["organize"]
        cfg, out_base = job["cfg"], job["out_dir"]
        zips = list(self._ds_files)
        total = len(zips)
        saved = 0
        renamed = 0
        failed = []        # (name, reason) — files that couldn't be saved
        notes = []         # (name, info) — kept original name, with the reason why
        protected = []     # password-protected archives we skipped
        bad_zips = []      # (name, reason) — archives that couldn't be extracted

        # Each ZIP is extracted into its own temp dir (nested zips expanded in
        # place), then every extracted file is structured into the output folder.
        for i, zip_path in enumerate(zips):
            if self._cancel_ds.is_set():
                break
            self._ds_timer.set_file(zip_path.name)

            # Skip password-protected archives — they can't be extracted here.
            if _is_password_protected(zip_path):
                protected.append(zip_path.name)
                self._ds_tracker.complete_item()
                continue

            self._set_ds_status(f"Extracting: {zip_path.name}  ({i + 1}/{total})")
            self._ds_tracker.begin_item()
            tmp = Path(tempfile.mkdtemp())
            try:
                with zipfile.ZipFile(zip_path, "r") as zf:
                    zf.extractall(tmp)
                _recursive_unzip_in_place(tmp)   # expand nested zips at any depth
                extracted = sorted(p for p in tmp.rglob("*") if p.is_file())
            except zipfile.BadZipFile:
                bad_zips.append((zip_path.name, "not a valid ZIP file or corrupted"))
                shutil.rmtree(tmp, ignore_errors=True)
                self._ds_tracker.complete_item()
                continue
            except Exception as exc:
                bad_zips.append((zip_path.name, str(exc)))
                shutil.rmtree(tmp, ignore_errors=True)
                self._ds_tracker.complete_item()
                continue

            # No output folder chosen → save beside the source zip ("same as
            # source"). The folder may already contain unrelated files; that's
            # fine — only the extracted files are written here.
            base = out_base if out_base else zip_path.parent
            try:
                ntotal = len(extracted)
                for j, src in enumerate(extracted):
                    if self._cancel_ds.is_set():
                        break
                    s, r = self._ds_structure_file(
                        src, zip_path.name, base, do_rename, do_date,
                        do_organize, cfg, i, total, j, ntotal, failed, notes)
                    saved += s
                    renamed += r
            finally:
                shutil.rmtree(tmp, ignore_errors=True)

            self._ds_tracker.complete_item()

        self.root.after(
            0, lambda: self._ds_finish(saved, renamed, failed, notes, protected, bad_zips))

    def _ds_structure_file(self, src, zip_name, out_base, do_rename, do_date,
                           do_organize, cfg, i, total, j, ntotal, failed, notes):
        """Optionally AI-rename and/or AI-date-prefix one extracted file, then
        copy it into the output folder. When both AI options are on the result is
        'Date_AISuggestedName'; date-only gives 'Date_OriginalName'. Mutates
        `failed`/`notes`; returns (saved_count, renamed_count) as 0/1."""
        new_stem = src.stem
        name_changed = False
        date_prefix = None
        where = f"{zip_name} ({i + 1}/{total}) — {j + 1}/{ntotal}"

        # Both AI features read the same opening text, so extract it just once.
        if do_rename or do_date:
            if _is_password_protected(src):
                notes.append((src.name, "password-protected — kept original name "
                                        "(use Password Removal first)"))
            else:
                self._set_ds_status(f"Reading: {src.name}  [{where}]")
                text = _extract_document_text(src)
                if not text:
                    notes.append((src.name, "no readable text (or unsupported type) — "
                                            "kept original name, no date added"))
                else:
                    if do_rename and not self._cancel_ds.is_set():
                        self._set_ds_status(f"Asking AI for a name: {src.name}  [{where}]")
                        suggestion = _ai_suggest_filename(text, cfg)
                        if suggestion and suggestion != src.stem:
                            new_stem = suggestion
                            name_changed = True
                        elif not suggestion:
                            notes.append((src.name, "no AI name suggestion (API "
                                                    "unavailable) — kept original name"))
                    if do_date and not self._cancel_ds.is_set():
                        self._set_ds_status(f"Asking AI for the date: {src.name}  [{where}]")
                        date_prefix = _ai_infer_date(text, cfg)
                        if not date_prefix:
                            notes.append((src.name, "no date could be inferred — "
                                                    "no date prefix added"))
                if self._cancel_ds.is_set():
                    return (0, 0)

        # Date goes first, then '_', then the (possibly AI-renamed) base name.
        if date_prefix:
            new_stem = f"{date_prefix}_{new_stem}"
            name_changed = True

        self._set_ds_status(f"Saving: {new_stem}{src.suffix}  [{where}]")
        target_dir = out_base
        if do_organize:
            target_dir = target_dir / _filetype_folder(src.suffix)
        try:
            # Safe to share a non-empty output folder: we only ever CREATE the
            # filetype subfolder (if missing) and COPY this one extracted file in.
            # Files already in the output folder are never moved into the new
            # subfolders, and _unique_output_path guarantees we never overwrite an
            # existing file (it suffixes _2, _3, …). Originals are copied, not moved.
            target_dir.mkdir(parents=True, exist_ok=True)
            out_path = _unique_output_path(target_dir, new_stem, src.suffix)
            shutil.copy2(src, out_path)
            return (1, 1 if name_changed else 0)
        except Exception as exc:
            failed.append((src.name, str(exc)))
            return (0, 0)

    def _set_ds_status(self, msg):
        self.root.after(0, lambda: self._ds_status_var.set(msg))

    def _ds_finish(self, saved, renamed, failed, notes, protected, bad_zips):
        self._ds_timer.stop()
        self._ds_tracker.finish()
        elapsed = self._ds_timer.elapsed_str()
        self._ds_timer.reset()
        self._ds_btn.config(state="normal")
        self._set_cancel_state(self._ds_cancel_btn, False)

        extra = ""
        if renamed:
            extra += f"\n{renamed} file(s) renamed."
        if protected:
            extra += (f"\n{len(protected)} archive(s) skipped (password-protected): "
                      + ", ".join(protected))
        if bad_zips:
            bad_lines = "\n".join(f"• {n}: {e}" for n, e in bad_zips)
            extra += f"\n\nArchives that couldn't be extracted:\n{bad_lines}"
        if notes:
            note_lines = "\n".join(f"• {n}: {info}" for n, info in notes)
            extra += f"\n\nKept original names:\n{note_lines}"

        if self._cancel_ds.is_set():
            self._ds_status_var.set(
                f"⏹  Cancelled. {saved} file(s) saved before stopping.")
            messagebox.showinfo(
                "Document Structuring Cancelled",
                f"Operation cancelled.\n{saved} file(s) saved before stopping."
                f"{extra}\n\nTotal time elapsed: {elapsed}",
            )
            return
        if not failed:
            self._ds_status_var.set(f"✅  Done! {saved} file(s) saved.")
            messagebox.showinfo(
                "Document Structuring Complete",
                f"{saved} file(s) extracted and saved to the output folder.{extra}"
                f"\n\nTotal time elapsed: {elapsed}",
            )
        else:
            err_lines = "\n".join(f"• {n}: {e}" for n, e in failed)
            self._ds_status_var.set(f"⚠️  {saved} saved, {len(failed)} failed.")
            messagebox.showerror(
                "Document Structuring — Some Files Failed",
                f"{saved} saved, {len(failed)} failed:{extra}\n\n{err_lines}"
                f"\n\nTotal time elapsed: {elapsed}",
            )

    # ── About & Updates ──────────────────────
    def _build_update_page(self, parent):
        tk.Label(
            parent, text="About & Updates",
            bg=APP_BG, fg="#1a1a1a", font=("Segoe UI", 14, "bold"),
            anchor="w", padx=18,
        ).pack(fill="x", pady=(14, 4))

        tk.Label(
            parent,
            text=("See which version you're running and check whether a newer one is "
                  "available. Updates download from the project's GitHub page and "
                  "install themselves — Box of Scraps closes and reopens on the new "
                  "version when it's done."),
            bg=APP_BG, fg="#555", font=("Segoe UI", 9),
            anchor="w", padx=18, justify="left", wraplength=620,
        ).pack(fill="x", pady=(0, 6))

        # Installed-version card
        card = tk.Frame(parent, bg=DROP_BG, highlightbackground=DROP_BD,
                        highlightthickness=2, relief="flat")
        card.pack(fill="x", padx=18, pady=(6, 8))
        row = tk.Frame(card, bg=DROP_BG)
        row.pack(fill="x", padx=14, pady=(10, 2))
        tk.Label(row, text="Installed version", bg=DROP_BG, fg="#4a6fa5",
                 font=("Segoe UI", 9)).pack(side="left")
        tk.Label(row, text=APP_VERSION, bg=DROP_BG, fg="#1a1a1a",
                 font=("Segoe UI", 11, "bold")).pack(side="left", padx=(8, 0))

        exe = _current_exe_path()
        if exe is None:
            where = ("Running from source (File.py) — in-app updating applies to the "
                     "packaged .exe only.")
        else:
            where = str(exe)
            provider = _cloud_sync_provider(exe.parent)
            if provider:
                where += f"\nStored in {provider} — updating still works normally."
        tk.Label(card, text=where, bg=DROP_BG, fg="#6b7f96", font=("Segoe UI", 8),
                 anchor="w", justify="left", wraplength=620,
                 ).pack(fill="x", padx=14, pady=(0, 10))

        # Release notes for whatever the check turns up
        notes_frame = tk.Frame(parent, bg=DROP_BG, highlightbackground=DROP_BD,
                               highlightthickness=2, relief="flat")
        notes_frame.pack(fill="both", expand=True, padx=18, pady=(0, 6))
        self._upd_notes_title = tk.Label(
            notes_frame, text="Press “Check for Updates” to see what's available.",
            bg=DROP_BG, fg="#4a6fa5", font=("Segoe UI", 9, "bold"),
            anchor="w", justify="left")
        self._upd_notes_title.pack(fill="x", padx=10, pady=(8, 4))
        inner = tk.Frame(notes_frame, bg=DROP_BG)
        inner.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        nscroll = tk.Scrollbar(inner, orient="vertical")
        self._upd_notes = tk.Text(
            inner, yscrollcommand=nscroll.set, wrap="word", height=6,
            bg="#ffffff", fg="#1a1a1a", font=("Segoe UI", 9),
            relief="flat", bd=1, highlightthickness=0, state="disabled")
        nscroll.config(command=self._upd_notes.yview)
        nscroll.pack(side="right", fill="y")
        self._upd_notes.pack(fill="both", expand=True)

        # Progress + status
        self._upd_progress = ttk.Progressbar(parent, mode="determinate", maximum=1000)
        self._upd_progress.pack(fill="x", padx=18, pady=(6, 2))
        self._upd_status_var = tk.StringVar(value="Ready — check GitHub for a newer version.")
        tk.Label(parent, textvariable=self._upd_status_var, bg=APP_BG, fg="#555",
                 font=("Segoe UI", 8), anchor="w").pack(fill="x", padx=20, pady=(0, 4))
        self._upd_timer = self._build_timer(parent)

        # Run + Cancel. One accent button that flips between checking and
        # installing, so there is never a dead "Install" button on screen.
        upd_btn_row = tk.Frame(parent, bg=APP_BG)
        upd_btn_row.pack(pady=(4, 14))
        self._upd_btn = tk.Button(
            upd_btn_row, text="Check for Updates", command=self._start_update_check,
            bg=ACCENT, fg=BTN_FG, font=("Segoe UI", 12, "bold"),
            relief="flat", padx=30, pady=8, cursor="hand2",
            activebackground="#005a9e", activeforeground=BTN_FG,
        )
        self._upd_btn.pack(side="left")
        self._upd_cancel_btn = tk.Button(
            upd_btn_row, text="Cancel", command=self._cancel_update_op,
            font=("Segoe UI", 10, "bold"), relief="flat", padx=16, pady=6,
            activebackground=CANCEL_BG_ACTIVE, activeforeground=BTN_FG,
        )
        self._upd_cancel_btn.pack(side="left", padx=(10, 0))
        self._set_cancel_state(self._upd_cancel_btn, False)

    def _set_upd_status(self, msg):
        self.root.after(0, lambda: self._upd_status_var.set(msg))

    def _upd_show_notes(self, title, body):
        """Fill the release-notes panel (Tk thread)."""
        self._upd_notes_title.config(text=title)
        self._upd_notes.config(state="normal")
        self._upd_notes.delete("1.0", "end")
        self._upd_notes.insert("1.0", body or "(No release notes were published.)")
        self._upd_notes.config(state="disabled")

    def _upd_set_mode(self, mode):
        """Point the accent button at checking or installing (Tk thread).

        This is the one button on the page: finding an update turns it into the
        install action in place, so there is never a dead 'Install' sitting on
        screen with nothing to install."""
        if mode == "install":
            self._upd_cancel_cooldown()     # installing is never rate-limited
            label = (self._upd_release.get("tag_name") or "").lstrip("vV")
            self._upd_btn.config(text=f"Download & Install {label}",
                                 command=self._start_update_install,
                                 state="normal", cursor="hand2")
        else:
            # Every route back to 'check' mode has just spent a GitHub request,
            # so the cooldown starts here — it also owns the button's label and
            # state until it expires.
            self._upd_btn.config(command=self._start_update_check)
            self._upd_begin_cooldown()

    def _upd_set_busy(self, busy, cancellable=True):
        """Lock the page while an update operation is in flight (Tk thread)."""
        self._upd_busy = busy
        if busy:
            self._upd_btn.config(state="disabled", cursor="arrow")
        elif time.monotonic() < self._upd_cooldown_until:
            self._upd_tick_cooldown()      # still cooling down — keep it locked
        else:
            self._upd_btn.config(state="normal", cursor="hand2")
        self._set_cancel_state(self._upd_cancel_btn, busy and cancellable)

    def _cancel_update_op(self):
        # Stops the download at its next chunk boundary; the version check is a
        # single short request, so there it just takes effect on return.
        self._cancel_update.set()
        self._set_cancel_state(self._upd_cancel_btn, False)
        self._upd_status_var.set("Stopping…")

    def _other_operation_running(self) -> bool:
        """True if any other tool page has an operation in flight.

        Every page enables its Cancel button for exactly the duration of its
        operation (`_set_cancel_state`), so an enabled Cancel is an existing,
        reliable busy signal — no new bookkeeping needed. Used to refuse an
        update that would otherwise close the app mid-conversion."""
        for attr in ("cancel_btn", "_uz_cancel_btn", "_ai_cancel_btn",
                     "_office_cancel_btn", "_pwd_cancel_btn", "_ds_cancel_btn"):
            btn = getattr(self, attr, None)
            if btn is None:
                continue
            try:
                if str(btn.cget("state")) == "normal":
                    return True
            except Exception:
                pass
        return False

    # ── update: check ────────────────────────
    def _upd_begin_cooldown(self):
        """Lock the Check button for UPDATE_CHECK_COOLDOWN seconds after a check
        (successful or not — a failed one still spent a request).

        Only ever applies while the button is in 'check' mode: once an update has
        been found the same button becomes Download & Install, which must stay
        clickable immediately."""
        self._upd_cooldown_until = time.monotonic() + UPDATE_CHECK_COOLDOWN
        self._upd_tick_cooldown()

    def _upd_tick_cooldown(self):
        """Count the cooldown down in the button's own label, so the wait is
        visible rather than the button just being mysteriously dead."""
        self._upd_cooldown_after = None
        if self._upd_busy:
            return                    # an operation took the button over
        remaining = self._upd_cooldown_until - time.monotonic()
        if remaining <= 0:
            self._upd_btn.config(text="Check for Updates", state="normal",
                                 cursor="hand2")
            return
        self._upd_btn.config(text=f"Check again in {int(remaining) + 1}s",
                             state="disabled", cursor="arrow")
        self._upd_cooldown_after = self.root.after(250, self._upd_tick_cooldown)

    def _upd_cancel_cooldown(self):
        """Drop a pending countdown tick (the install path takes the button)."""
        if self._upd_cooldown_after is not None:
            try:
                self.root.after_cancel(self._upd_cooldown_after)
            except Exception:
                pass
            self._upd_cooldown_after = None

    def _start_update_check(self):
        if self._upd_busy or time.monotonic() < self._upd_cooldown_until:
            return
        self._cancel_update.clear()
        self._upd_release = None
        self._upd_asset = None
        self._upd_set_busy(True)
        self._upd_progress["value"] = 0
        self._upd_status_var.set("Checking GitHub for a newer version…")
        threading.Thread(target=self._upd_check_worker, daemon=True).start()

    def _upd_check_worker(self):
        try:
            release = _fetch_latest_release()
        except RuntimeError as e:
            self.root.after(0, self._upd_check_failed, str(e))
            return
        if self._cancel_update.is_set():
            self.root.after(0, self._upd_cancelled)
            return
        self.root.after(0, self._upd_check_done, release)

    def _upd_check_done(self, release):
        self._upd_set_busy(False)
        tag = (release.get("tag_name") or "").strip()
        latest = tag.lstrip("vV") or "unknown"
        notes = (release.get("body") or "").strip()

        if _parse_version(tag) <= _parse_version(APP_VERSION):
            self._upd_status_var.set(f"You're up to date — {APP_VERSION} is the latest version.")
            self._upd_show_notes(
                f"Up to date — latest release is {latest}",
                notes)
            self._upd_set_mode("check")
            return

        asset = _find_release_asset(release, UPDATE_ASSET_NAME)
        if asset is None:
            self._upd_status_var.set(
                f"Version {latest} is available, but it has no "
                f"{UPDATE_ASSET_NAME} download attached.")
            self._upd_show_notes(f"Version {latest} is available", notes)
            self._upd_set_mode("check")
            return

        self._upd_release = release
        self._upd_asset = asset
        size_mb = int(asset.get("size") or 0) / 1048576
        self._upd_status_var.set(
            f"Version {latest} is available ({size_mb:,.0f} MB) — "
            f"you have {APP_VERSION}.")
        self._upd_show_notes(f"What's new in {latest}", notes)
        self._upd_set_mode("install")

    def _upd_check_failed(self, msg):
        self._upd_set_busy(False)
        self._upd_set_mode("check")
        self._upd_status_var.set("Couldn't check for updates.")
        self._upd_show_notes("Update check failed", msg)
        if messagebox.askyesno(
                "Update Check Failed",
                f"{msg}\n\nOpen the releases page in your browser instead?"):
            _open_releases_page()

    def _upd_cancelled(self):
        self._upd_set_busy(False)
        self._upd_timer.stop()
        self._upd_progress["value"] = 0
        self._upd_status_var.set("Update cancelled.")
        if self._upd_release is None:
            # A cancelled *check* still spent a GitHub request; a cancelled
            # download did not, and leaves the install button live.
            self._upd_begin_cooldown()

    # ── update: install ──────────────────────
    def _start_update_install(self):
        """Run every precheck that can fail cheaply *before* committing to a
        ~120 MB download, then confirm and hand off to the worker."""
        if self._upd_busy or not self._upd_release or not self._upd_asset:
            return

        exe = _current_exe_path()
        if exe is None:
            messagebox.showinfo(
                "Not Available in Development",
                "In-app updating only works in the packaged Box of Scraps .exe.\n\n"
                "You're running from source, where the running program is "
                "python.exe — replacing that is not something this tool will do.")
            return

        if self._other_operation_running():
            messagebox.showwarning(
                "Operation in Progress",
                "Another tool is still running.\n\nInstalling an update closes and "
                "reopens the app, which would interrupt it. Please wait for it to "
                "finish (or cancel it) and try again.")
            return

        # Writability decides whether this can work at all — check it before the
        # download rather than discovering it at the moment of the swap.
        if not _dir_is_writable(exe.parent):
            if messagebox.askyesno(
                    "Administrator Rights Needed",
                    f"Box of Scraps can't write to its own folder:\n\n{exe.parent}\n\n"
                    "That usually means it's installed somewhere only an "
                    "administrator can change, such as Program Files.\n\n"
                    "Restart Box of Scraps as an administrator and try the update "
                    "again?"):
                if _relaunch_as_admin(exe):
                    self.root.destroy()
                else:
                    messagebox.showerror(
                        "Couldn't Restart",
                        "The elevation prompt was declined or failed.\n\n"
                        "You can also update by downloading the new version "
                        "manually from the releases page.")
            return

        size = int(self._upd_asset.get("size") or 0)
        # Room for the download plus the backup copy kept until the next launch.
        if size and not _has_room_for_update(exe.parent, size * 2 + 64 * 1024 * 1024):
            messagebox.showerror(
                "Not Enough Disk Space",
                f"Updating needs about {(size * 2) / 1048576:,.0f} MB free on "
                f"{exe.drive or exe.parent} — the new version plus a backup of the "
                "current one.\n\nFree up some space and try again.")
            return

        latest = (self._upd_release.get("tag_name") or "").lstrip("vV")
        detail = (f"Update from {APP_VERSION} to {latest}?\n\n"
                  f"Download size: {size / 1048576:,.0f} MB\n\n"
                  "Box of Scraps will close and reopen on the new version when the "
                  "update finishes. Your current version is kept until the next "
                  "launch in case anything goes wrong.")
        provider = _cloud_sync_provider(exe.parent)
        if provider:
            detail += (f"\n\nNote: this app lives in a {provider} folder. The update "
                       f"works normally — just let {provider} finish syncing "
                       "afterwards before copying it elsewhere.")
        if not messagebox.askokcancel("Install Update", detail):
            return

        self._cancel_update.clear()
        self._upd_last_pct = 0.0
        self._upd_set_busy(True)
        self._upd_progress["value"] = 0
        self._upd_status_var.set("Starting download…")
        self._upd_timer.start()
        self._upd_timer.set_file(UPDATE_ASSET_NAME)
        threading.Thread(target=self._upd_install_worker, args=(exe,),
                         daemon=True).start()

    def _upd_on_progress(self, got, total):
        """Download progress, called from the worker thread."""
        pct = (got / total) if total else 0.0
        # Throttle to ~one UI update per half-percent: a 120 MB download would
        # otherwise queue thousands of `after` callbacks onto the Tk thread.
        if total and (pct - self._upd_last_pct) < 0.005 and got < total:
            return
        self._upd_last_pct = pct
        self.root.after(0, self._upd_render_progress, pct,
                        got / 1048576, (total or 0) / 1048576)

    def _upd_render_progress(self, pct, got_mb, total_mb):
        self._upd_progress["value"] = pct * 1000
        if total_mb:
            self._upd_status_var.set(
                f"Downloading… {got_mb:,.0f} MB of {total_mb:,.0f} MB  "
                f"({pct * 100:.0f}%)")
        else:
            self._upd_status_var.set(f"Downloading… {got_mb:,.0f} MB")

    def _upd_install_worker(self, exe: Path):
        """Download, validate, and swap in the new exe, then relaunch.

        Nothing touches the installed application until the download has been
        fully verified, so a failure during the slow, failure-prone part leaves
        the current version completely untouched."""
        release, asset = self._upd_release, self._upd_asset
        latest = (release.get("tag_name") or "").lstrip("vV")
        size = int(asset.get("size") or 0)
        tmpdir = None
        staged = None
        try:
            # Download into the system temp folder, NOT next to the exe: if the
            # app lives in a synced folder, staging ~120 MB there would push the
            # whole download up to the cloud and then delete it again.
            tmpdir = Path(tempfile.mkdtemp(prefix="BoxOfScraps-update-"))
            download = tmpdir / UPDATE_ASSET_NAME
            _download_release_asset(
                asset.get("browser_download_url"), download,
                expected_size=size, progress_cb=self._upd_on_progress,
                cancel_event=self._cancel_update)

            self._set_upd_status("Verifying the download…")
            if download.stat().st_size < UPDATE_MIN_EXE_BYTES:
                raise RuntimeError(
                    "The downloaded file is far too small to be a real build — "
                    "the download was probably intercepted or truncated.")
            if not _looks_like_windows_exe(download):
                raise RuntimeError(
                    "The downloaded file isn't a Windows program. A network "
                    "filter or sign-in page may have replaced it.")

            # Verify against a published checksum when the release carries one.
            # Releases without it still install — the PE check above is the floor.
            checksum_asset = _find_release_asset(
                release, UPDATE_ASSET_NAME + ".sha256")
            if checksum_asset is not None:
                sums = tmpdir / "expected.sha256"
                try:
                    _download_release_asset(
                        checksum_asset.get("browser_download_url"), sums,
                        cancel_event=self._cancel_update)
                    expected = sums.read_text(
                        encoding="utf-8", errors="ignore").split()[0].strip().lower()
                except _UpdateCancelled:
                    raise
                except Exception:
                    expected = ""
                if expected:
                    if _sha256_file(download).lower() != expected:
                        raise RuntimeError(
                            "The download doesn't match the checksum published "
                            "with the release, so it was not installed. Try "
                            "again; if it keeps failing, download the new "
                            "version manually.")

            if self._cancel_update.is_set():
                raise _UpdateCancelled()

            # Belt-and-braces: urllib doesn't apply the Mark-of-the-Web tag, so
            # there is normally nothing here to remove — do it anyway so the
            # no-Unblock guarantee doesn't depend on that.
            _strip_mark_of_the_web(download)

            self._set_upd_status("Installing the new version…")
            staged = _stage_beside(download, exe)
            _strip_mark_of_the_web(staged)
            _swap_in_new_exe(staged, exe)
            staged = None                      # consumed by the swap
            _strip_mark_of_the_web(exe)

            self._set_upd_status(f"Restarting on {latest}…")
            _relaunch(exe)
        except _UpdateCancelled:
            self.root.after(0, self._upd_cancelled)
            return
        except PermissionError as e:
            self.root.after(
                0, self._upd_install_failed,
                "Windows wouldn't let the new version take the place of the "
                f"running one:\n\n{e}\n\nThis is usually a cloud-sync client or "
                "an antivirus scanner holding the file open. Your current "
                "version is untouched — please try again in a moment.")
            return
        except Exception as e:
            self.root.after(0, self._upd_install_failed,
                            f"{type(e).__name__}: {e}")
            return
        finally:
            for leftover in (staged,):
                if leftover is not None:
                    try:
                        os.remove(leftover)
                    except OSError:
                        pass
            if tmpdir is not None:
                shutil.rmtree(tmpdir, ignore_errors=True)

        self.root.after(0, self._upd_restart, latest)

    def _upd_install_failed(self, msg):
        self._upd_set_busy(False)
        self._upd_timer.stop()
        self._upd_progress["value"] = 0
        self._upd_status_var.set("The update didn't install — your current version is unchanged.")
        self._upd_show_notes("Update failed", msg)
        if messagebox.askyesno(
                "Update Failed",
                f"{msg}\n\nYour current version is unchanged and still works.\n\n"
                "Open the releases page to download the update manually?"):
            _open_releases_page()

    def _upd_restart(self, version):
        """The swap succeeded and the new exe is already starting — close this
        one. The old build is left on disk as BoxOfScraps.exe.old and the new
        process deletes it at startup."""
        self._upd_timer.stop()
        self._upd_progress["value"] = 1000
        self._upd_status_var.set(f"Updated to {version} — restarting…")
        self.root.update_idletasks()
        self.root.after(700, self.root.destroy)

    def _allowed_extensions(self):
        m = self._mode.get()
        if m.startswith("HTML"):
            return (".html", ".htm", ".mht", ".mhtml")
        if m.startswith("Visio"):
            return (".vsd", ".vsdx")
        if m.startswith("Excel"):
            return (".xls", ".xlsx", ".xlsm", ".xlsb", ".csv")
        if m.startswith("Word"):
            return (".doc", ".docx", ".txt", ".rtf", ".xml")
        if m.startswith("PowerPoint"):
            return (".ppt", ".pptx", ".pps", ".ppsx")
        if m.startswith("Image"):
            return (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif")
        if m.startswith("Text Report"):
            return TEXT_REPORT_EXTS
        return (".eml", ".msg")

    def _on_mode_change(self, event=None):
        # Clear files when mode changes — they may be wrong type
        self._clear_files()
        self._update_drop_label()

    # ── file management ──────────────────────
    def _add_files(self):
        m = self._mode.get()
        if m.startswith("HTML"):
            filetypes = [("HTML files", "*.html *.htm *.mht *.mhtml"), ("All files", "*.*")]
            title = "Select HTML files"
        elif m.startswith("Visio"):
            filetypes = [("Visio files", "*.vsd *.vsdx"), ("All files", "*.*")]
            title = "Select Visio files"
        elif m.startswith("Excel"):
            filetypes = [("Excel / CSV files", "*.xls *.xlsx *.xlsm *.xlsb *.csv"), ("All files", "*.*")]
            title = "Select Excel / CSV files"
        elif m.startswith("Word"):
            filetypes = [("Word / Text files", "*.doc *.docx *.txt *.rtf *.xml"), ("All files", "*.*")]
            title = "Select Word / Text files"
        elif m.startswith("PowerPoint"):
            filetypes = [("PowerPoint files", "*.ppt *.pptx *.pps *.ppsx"), ("All files", "*.*")]
            title = "Select PowerPoint files"
        elif m.startswith("Image"):
            filetypes = [("Image files", "*.jpg *.jpeg *.png *.bmp *.gif *.tiff *.tif"), ("All files", "*.*")]
            title = "Select image files"
        elif m.startswith("Text Report"):
            filetypes = [("Text report files", "*.rep *.rpt *.prn *.lst"), ("All files", "*.*")]
            title = "Select text report files"
        else:
            filetypes = [("Email files", "*.eml *.msg"), ("All files", "*.*")]
            title = "Select email files"
        paths = filedialog.askopenfilenames(title=title, filetypes=filetypes)
        for p in paths:
            self._add_path(p)

    def _on_drop(self, event):
        for p in self.root.tk.splitlist(event.data):
            self._add_path(p)

    def _add_path(self, p):
        path = Path(p)
        if path.suffix.lower() not in self._allowed_extensions():
            return
        if path not in self.files:
            self.files.append(path)
            self.file_list.insert("end", path.name)
        self._update_drop_label()

    def _remove_selected(self):
        for i in sorted(self.file_list.curselection(), reverse=True):
            self.file_list.delete(i)
            del self.files[i]
        self._update_drop_label()

    def _clear_files(self):
        self.files.clear()
        self.file_list.delete(0, "end")
        self._update_drop_label()
        # Return the page to its ready state (clear any "Done!"/error status).
        self.status_var.set("Ready — add files to get started.")
        self._pdf_tracker.reset()

    def _update_drop_label(self):
        if self.files:
            self.drop_label.config(text=f"{len(self.files)} file(s) queued")
        elif self._mode.get().startswith("HTML"):
            self.drop_label.config(text="Drop .html / .htm / .mht files here\nor click 'Add Files'")
        elif self._mode.get().startswith("Visio"):
            self.drop_label.config(text="Drop .vsd / .vsdx files here\nor click 'Add Files'")
        elif self._mode.get().startswith("Excel"):
            self.drop_label.config(text="Drop .xls / .xlsx / .csv files here\nor click 'Add Files'")
        elif self._mode.get().startswith("Word"):
            self.drop_label.config(text="Drop .doc / .docx / .txt / .rtf files here\nor click 'Add Files'")
        elif self._mode.get().startswith("PowerPoint"):
            self.drop_label.config(text="Drop .ppt / .pptx files here\nor click 'Add Files'")
        elif self._mode.get().startswith("Image"):
            self.drop_label.config(text="Drop image files here\nor click 'Add Files'")
        elif self._mode.get().startswith("Text Report"):
            self.drop_label.config(text="Drop .rep / .rpt / .prn / .lst files here\nor click 'Add Files'")
        else:
            self.drop_label.config(text="Drop .eml / .msg files here\nor click 'Add Files'")

    def _choose_output(self):
        d = filedialog.askdirectory(title="Select output folder")
        if d:
            self._out_dir = Path(d)
            self.out_var.set(str(self._out_dir))
        else:
            self._out_dir = None
            self.out_var.set("Same as source file")

    # ── conversion ───────────────────────────
    def _set_cancel_state(self, btn, enabled):
        """Grey + disabled, or red + clickable."""
        if enabled:
            btn.config(state="normal", bg=CANCEL_BG, fg=BTN_FG, cursor="hand2")
        else:
            btn.config(state="disabled", bg=CANCEL_OFF_BG, fg=CANCEL_OFF_FG, cursor="arrow")

    def _cancel_conversion(self):
        """Force-stop the running conversion. Kills the dedicated Office
        instances we launched (Word/Excel/PowerPoint/Visio) so a stuck file is
        aborted immediately. Outlook is the user's email client and is never
        killed — in email mode this falls back to a graceful stop."""
        self._cancel_convert.set()
        self._set_cancel_state(self.cancel_btn, False)
        self.status_var.set("Stopping… force-closing the current conversion.")
        for pid in list(getattr(self, "_pdf_app_pids", [])):
            _terminate_pid(pid)

    def _start_convert(self):
        if not self.files:
            m = self._mode.get()
            if m.startswith("HTML"):         label = "HTML"
            elif m.startswith("Visio"):      label = "Visio"
            elif m.startswith("Excel"):      label = "Excel / CSV"
            elif m.startswith("Word"):       label = "Word / text"
            elif m.startswith("PowerPoint"): label = "PowerPoint"
            elif m.startswith("Image"):      label = "image"
            elif m.startswith("Text Report"): label = "text report"
            else:                            label = "email"
            messagebox.showwarning("No Files", f"Please add {label} files first.")
            return
        self._cancel_convert.clear()
        self.convert_btn.config(state="disabled")
        self._set_cancel_state(self.cancel_btn, True)
        self._pdf_tracker.start(len(self.files))
        self._pdf_timer.start()
        threading.Thread(target=self._convert_worker, daemon=True).start()

    def _convert_worker(self):
        m = self._mode.get()
        html_mode   = m.startswith("HTML")
        visio_mode  = m.startswith("Visio")
        excel_mode  = m.startswith("Excel")
        word_mode   = m.startswith("Word")
        ppt_mode    = m.startswith("PowerPoint")
        image_mode  = m.startswith("Image")
        report_mode = m.startswith("Text Report")
        email_mode  = not any([html_mode, visio_mode, excel_mode, word_mode,
                               ppt_mode, image_mode, report_mode])

        pythoncom.CoInitialize()
        self._pdf_app_pids = []
        outlook = word = visio = excel = powerpoint = None
        success, failed, protected = 0, [], []

        def _launch(family):
            # Dedicated, isolated Office instance + track its PID so Cancel can
            # force-kill only ours (never the user's open Office windows).
            try:
                app, pids = _launch_office_isolated(family)
            except Exception:
                app, pids = None, set()
            self._pdf_app_pids.extend(pids)
            return app

        def _abort(title, msg):
            self.root.after(0, lambda: (
                self.convert_btn.config(state="normal"),
                self._set_cancel_state(self.cancel_btn, False),
                self._pdf_timer.reset(),
                messagebox.showerror(title, msg),
            ))

        try:
            # Outlook — email only. This is the user's email client, so it is
            # launched shared and NEVER force-killed (email cancel is graceful).
            if email_mode:
                try:
                    outlook = _ensure_outlook()
                except Exception as exc:
                    _abort("Outlook Error", str(exc))
                    return

            # Word — email, HTML, Word doc, image, and text report modes
            if email_mode or html_mode or word_mode or image_mode or report_mode:
                word = _launch("word")
                if word is None and (html_mode or word_mode or image_mode or report_mode):
                    if html_mode:     label = "HTML"
                    elif image_mode:  label = "Image"
                    elif report_mode: label = "text report"
                    else:             label = "Word document"
                    _abort("Word Not Found",
                           f"Microsoft Word is required for {label} conversion.\n"
                           "Please ensure Word is installed.")
                    return

            if visio_mode:
                visio = _launch("visio")
                if visio is None:
                    _abort("Visio Not Found",
                           "Microsoft Visio is required for Visio conversion.\n"
                           "Please ensure Visio is installed.")
                    return

            if excel_mode:
                excel = _launch("excel")
                if excel is None:
                    _abort("Excel Not Found",
                           "Microsoft Excel is required for Excel conversion.\n"
                           "Please ensure Excel is installed.")
                    return

            if ppt_mode:
                powerpoint = _launch("powerpoint")
                if powerpoint is None:
                    _abort("PowerPoint Not Found",
                           "Microsoft PowerPoint is required for PowerPoint conversion.\n"
                           "Please ensure PowerPoint is installed.")
                    return

            total = len(self.files)
            if visio_mode:      mode_key = "visio"
            elif html_mode:     mode_key = "html"
            elif excel_mode:    mode_key = "excel"
            elif word_mode:     mode_key = "word_doc"
            elif ppt_mode:      mode_key = "powerpoint"
            elif image_mode:    mode_key = "image"
            elif report_mode:   mode_key = "text_report"
            else:               mode_key = "email"

            # The one isolated Office app this mode drives — the app the
            # per-file watchdog may have to kill and replace. (Email mode also
            # uses Outlook, but that is the user's mail client: shared, never
            # tracked, never killed.)
            family = {"visio": "visio", "excel": "excel",
                      "powerpoint": "powerpoint"}.get(mode_key, "word")

            for i, src in enumerate(self.files):
                if self._cancel_convert.is_set():
                    break
                # Skip password-protected files — point the user at Password Removal.
                if _is_password_protected(src):
                    protected.append(src.name)
                    self._pdf_tracker.complete_item()
                    continue
                self._set_status(f"Converting: {src.name}  ({i + 1}/{total})")
                self._pdf_timer.set_file(src.name)
                self._pdf_tracker.begin_item()
                out_dir = self._out_dir if self._out_dir else src.parent
                out_path = _unique_pdf_path(out_dir, src.stem)
                # Bound this file so one stuck conversion can't hang the batch.
                wd = _FileWatchdog(self._pdf_app_pids)
                try:
                    with wd:
                        convert_file(src, out_dir, outlook, word, visio,
                                     excel, powerpoint, mode_key)
                    success += 1
                except Exception as exc:
                    # Drop any partial PDF a force-kill may have left behind.
                    try:
                        if out_path.exists():
                            out_path.unlink()
                    except Exception:
                        pass
                    if self._cancel_convert.is_set():
                        pass                    # user cancelled — not a failure
                    elif wd.fired:
                        failed.append((src.name, "timed out — appeared stuck, skipped"))
                    else:
                        failed.append((src.name, str(exc)))

                # The watchdog killed our instance — stand up a fresh one so the
                # rest of the batch still has a healthy app to work with.
                if wd.fired and not self._cancel_convert.is_set():
                    self._pdf_app_pids = []
                    replacement = _launch(family)
                    if   family == "excel":      excel = replacement
                    elif family == "visio":      visio = replacement
                    elif family == "powerpoint": powerpoint = replacement
                    else:                        word = replacement

                self._pdf_tracker.complete_item()
        finally:
            # Quit our dedicated Office instances, then make sure none linger
            # (e.g. after a kill). Outlook is intentionally left alone.
            for app in (word, visio, excel, powerpoint):
                if app is not None:
                    try:
                        app.Quit()
                    except Exception:
                        pass
            for pid in list(self._pdf_app_pids):
                _terminate_pid(pid)
            self._pdf_app_pids = []
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

        self.root.after(0, lambda: self._finish(success, failed, protected))

    def _set_status(self, msg):
        self.root.after(0, lambda: self.status_var.set(msg))

    def _finish(self, success, failed, protected=None):
        self._pdf_timer.stop()
        self._pdf_tracker.finish()
        elapsed = self._pdf_timer.elapsed_str()
        self._pdf_timer.reset()
        self.convert_btn.config(state="normal")
        self._set_cancel_state(self.cancel_btn, False)
        protected = protected or []
        pw_note = _password_skip_note(protected)
        skip_status = f"  {len(protected)} skipped (password-protected)." if protected else ""
        if self._cancel_convert.is_set():
            self.status_var.set(f"⏹  Cancelled. {success} file(s) converted before stopping.")
            messagebox.showinfo(
                "Conversion Cancelled",
                f"Conversion was cancelled.\n{success} file(s) converted before stopping."
                f"\n\nTotal time elapsed: {elapsed}{pw_note}",
            )
        elif not failed:
            self.status_var.set(f"✅  Done! {success} file(s) converted successfully.{skip_status}")
            messagebox.showinfo(
                "Conversion Complete",
                f"{success} file(s) converted to PDF successfully."
                f"\n\nTotal time elapsed: {elapsed}{pw_note}",
            )
        else:
            err_lines = "\n".join(f"• {n}: {e}" for n, e in failed)
            self.status_var.set(f"⚠️  {success} converted, {len(failed)} failed.{skip_status}")
            messagebox.showerror(
                "Conversion Errors",
                f"{success} succeeded, {len(failed)} failed:\n\n{err_lines}"
                f"\n\nTotal time elapsed: {elapsed}{pw_note}",
            )

    # ── window ───────────────────────────────
    def _center_window(self, w, h):
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    def run(self):
        self.root.mainloop()


# ─────────────────────────────────────────────
def main():
    app = ConverterApp()
    app.run()


if __name__ == "__main__":
    main()
