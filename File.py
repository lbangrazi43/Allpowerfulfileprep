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
import zipfile

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

# Outlook constants
olFolderInbox      = 6
olMsg              = 3      # .msg save format
PR_ATTACH_METHOD   = 0x37200003

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

            # PR_ATTACH_METHOD: 5 = OLE object (skip), 6 = embedded message
            # Type 5/6 are rarely useful as files; skip them
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


def _eml_to_pdf(src_path: Path, out_path: Path, outlook, word):
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


def _ensure_word():
    """
    Launch a hidden Word Application COM instance.
    Returns the COM object, or None if Word is not installed.
    """
    try:
        w = win32com.client.Dispatch("Word.Application")
        w.Visible = False
        w.DisplayAlerts = False
        return w
    except Exception:
        return None


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


def _ensure_visio():
    """
    Launch a hidden Visio Application COM instance.
    Returns the COM object, or None if Visio is not installed.
    """
    try:
        v = win32com.client.Dispatch("Visio.Application")
        v.Visible = False
        return v
    except Exception:
        return None


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


def _ensure_excel():
    """
    Launch a hidden Excel Application COM instance.
    Returns the COM object, or None if Excel is not installed.
    """
    try:
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        return xl
    except Exception:
        return None


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


def _ensure_powerpoint():
    """
    Launch a hidden PowerPoint Application COM instance.
    Returns the COM object, or None if PowerPoint is not installed.
    """
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        # PowerPoint doesn't support Visible=False the same way;
        # minimise the window instead so it doesn't steal focus
        ppt.WindowState = 2   # ppWindowMinimized
        return ppt
    except Exception:
        return None


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
                         ("SaveNormalPrompt", False)):
            try:
                setattr(app.Options, opt, val)
            except Exception:
                pass
    elif family == "excel":
        for prop, val in (("AskToUpdateLinks", False),
                          ("AlertBeforeOverwriting", False)):
            try:
                setattr(app, prop, val)
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


def convert_file(src_path: Path, out_dir: Path, outlook, word, visio, excel, powerpoint, mode: str) -> Path:
    """
    Convert a single file to PDF.
    mode: 'email', 'html', 'visio', 'excel', 'word_doc', 'powerpoint', 'image'
    Returns the path of the created PDF. Never overwrites an existing PDF.
    """
    ext = src_path.suffix.lower()
    out_path = _unique_pdf_path(out_dir, src_path.stem)

    if mode == "email":
        if ext == ".msg":
            _msg_to_pdf(src_path, out_path, outlook, word)
        elif ext == ".eml":
            _eml_to_pdf(src_path, out_path, outlook, word)
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
    COM applications are initialized lazily and quit when done.
    """
    if not HAS_COM:
        status_cb("Auto-PDF skipped: pywin32 is not installed.")
        return []

    pdf_failed = []
    pythoncom.CoInitialize()
    word = excel = visio = powerpoint = outlook = None

    try:
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

                if ext == "":
                    # No extension — treat as plain text via a temp .txt copy so
                    # Word opens it without a format-detection dialog.
                    status_cb(f"Auto-PDF ({i + 1}/{total}): {src.name}  (no extension, treating as text)")
                    tmp_txt = src.parent / (src.name + ".txt")
                    out_path = _unique_pdf_path(src.parent, src.name)
                    try:
                        if word is None:
                            word = _ensure_word()
                        shutil.copy2(str(src), str(tmp_txt))
                        _word_to_pdf(tmp_txt, out_path, word)
                        if out_path.exists() and out_path.stat().st_size > 0:
                            src.unlink()
                            converted += 1
                        else:
                            pdf_failed.append(src.name)
                    except Exception:
                        pdf_failed.append(src.name)
                    finally:
                        try:
                            tmp_txt.unlink()
                        except Exception:
                            pass
                    continue

                mode = AUTO_PDF_EXTS[ext]
                status_cb(f"Auto-PDF ({i + 1}/{total}): {src.name}")
                try:
                    if mode == "image":
                        if word is None:
                            word = _ensure_word()
                        out_path = _unique_pdf_path(src.parent, src.stem)
                        _image_to_pdf(src, out_path, word)
                    else:
                        if mode in ("word_doc", "html") and word is None:
                            word = _ensure_word()
                        if mode == "excel" and excel is None:
                            excel = _ensure_excel()
                        if mode == "visio" and visio is None:
                            visio = _ensure_visio()
                        if mode == "powerpoint" and powerpoint is None:
                            powerpoint = _ensure_powerpoint()
                        if mode == "email":
                            if outlook is None:
                                outlook = _ensure_outlook()
                            if word is None:
                                word = _ensure_word()
                        out_path = convert_file(
                            src, src.parent, outlook, word, visio, excel, powerpoint, mode
                        )

                    # Only delete the original if the PDF was actually written
                    if out_path.exists() and out_path.stat().st_size > 0:
                        src.unlink()
                        converted += 1
                    else:
                        pdf_failed.append(src.name)
                except Exception:
                    pdf_failed.append(src.name)

            status_cb(f"Auto-PDF done: {converted}/{total} file(s) converted.")

    finally:
        for app in (word, excel, visio, powerpoint):
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
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
    extracted_files = []   # explicit list of files extracted from the zip(s)
    total = len(zip_paths)

    for i, zip_path in enumerate(zip_paths):
        if cancelled():
            break
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

    finish_cb(success, failed, pdf_failed)


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
        from PIL import Image, ImageTk
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
            return
        now = time.monotonic()
        if self._start is not None:
            self.total_var.set("Total Time Elapsed:  " + self._fmt(now - self._start))
        if self._file_name is not None and self._file_start is not None:
            self.file_var.set(
                f"Current file ({self._file_name}):  " + self._fmt(now - self._file_start))
        self.root.after(250, self._tick)

    def stop(self):
        if not self._running and self._start is None:
            return
        self._running = False
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
        self._start = None
        self._file_name = None
        self._file_start = None
        self.total_var.set("")
        self.file_var.set("")
        if self._shown:
            self.total_lbl.pack_forget()
            self.file_lbl.pack_forget()
            self._shown = False


class ConverterApp:
    def __init__(self):
        # Give the process its own AppUserModelID *before* the window exists so
        # Windows ties the taskbar button to our icon, not python/Tk's default.
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "AllPowerfulFilePrep.App")
        except Exception:
            pass

        if HAS_DND:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()

        # Hide the window *immediately* — before any geometry/icon work forces a
        # paint — so the blank Tk window never flashes on screen at launch.
        self.root.withdraw()

        self.root.title("File Preparation Toolkit")
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

        # Navigation state
        self._active_page = None
        self._page_frames = {}
        self._nav_buttons = {}

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
            text="All Powerful\nFile Prep",
            bg=SIDEBAR_BG, fg="#ffffff",
            font=("Segoe UI", 11, "bold"),
            justify="center", pady=20,
        ).pack(fill="x")

        tk.Frame(self._sidebar, bg="#3a5068", height=1).pack(fill="x", padx=12, pady=(0, 8))

        for page_key, label in [("pdf", "  PDF Conversion"),
                                ("unzip", "  Folder Unzipping"),
                                ("aiprep", "  Markdown Conversion"),
                                ("office", "  Office Modernizer"),
                                ("pwd", "  Password Removal")]:
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

    def _show_page(self, key: str):
        """Switch the visible content page, building it lazily on first visit."""
        for btn in self._nav_buttons.values():
            btn.config(bg=SIDEBAR_BG, fg=SIDEBAR_FG)
        self._nav_buttons[key].config(bg=SIDEBAR_SEL, fg="#ffffff")
        self._active_page = key

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

        self._page_frames[key].pack(fill="both", expand=True)

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
        mode_menu = ttk.Combobox(
            mode_frame,
            textvariable=self._mode,
            values=["Email (.eml, .msg)", "HTML / MHT (.html, .htm, .mht)", "Visio (.vsd, .vsdx)", "Excel (.xls, .xlsx, .csv)", "Word (.doc, .docx, .txt, .rtf, .xml)", "PowerPoint (.ppt, .pptx)", "Image (.jpg, .png, .bmp, .gif, .tiff)"],
            state="readonly",
            width=26,
            font=("Segoe UI", 9),
        )
        mode_menu.pack(side="left", padx=8)
        mode_menu.bind("<<ComboboxSelected>>", self._on_mode_change)

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
            tk.Radiobutton(
                uz_org_frame, text=text, variable=self._uz_separate, value=val,
                bg=APP_BG, fg="#333", activebackground=APP_BG,
                font=("Segoe UI", 9),
            ).pack(side="left", padx=(0, 14))

        # Auto-PDF checkbox
        uz_autopdf_frame = tk.Frame(parent, bg=APP_BG)
        uz_autopdf_frame.pack(fill="x", padx=18, pady=(6, 2))

        self._uz_auto_pdf = tk.BooleanVar(value=False)
        tk.Checkbutton(
            uz_autopdf_frame,
            text="Auto-convert files to PDF after unzipping  (originals deleted on success)",
            variable=self._uz_auto_pdf,
            bg=APP_BG, fg="#333", activebackground=APP_BG,
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(side="left")

        # Progress bar
        self._uz_progress = ttk.Progressbar(parent, mode="determinate")
        self._uz_progress.pack(fill="x", padx=18, pady=(6, 2))

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
        self._uz_progress["value"] = 0
        self._uz_progress["maximum"] = len(self._zip_files)
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
                lambda: self.root.after(0, lambda: self._uz_progress.step(1)),
                lambda s, f, pf: self.root.after(0, lambda: self._uz_finish(s, f, pf)),
                self._cancel_unzip,
                self._uz_timer.set_file,
            ),
            daemon=True,
        ).start()

    def _cancel_unzip_op(self):
        """Request that the running unzip/auto-convert stop as soon as possible."""
        self._cancel_unzip.set()
        self._set_cancel_state(self._uz_cancel_btn, False)
        self._uz_status_var.set("Cancelling… finishing the current item, then stopping.")

    def _uz_finish(self, success, failed, pdf_failed=None):
        self._uz_timer.stop()
        elapsed = self._uz_timer.elapsed_str()
        self._uz_timer.reset()
        self._uz_btn.config(state="normal")
        self._set_cancel_state(self._uz_cancel_btn, False)
        if self._cancel_unzip.is_set():
            self._uz_status_var.set(f"⏹  Cancelled. {success} archive(s) extracted before stopping.")
            messagebox.showinfo(
                "Unzip Cancelled",
                f"Operation was cancelled.\n{success} archive(s) extracted before stopping."
                f"\n\nTotal time elapsed: {elapsed}",
            )
            return
        if not failed:
            self._uz_status_var.set(f"Done! {success} archive(s) extracted successfully.")
            messagebox.showinfo(
                "Unzip Complete",
                f"{success} archive(s) extracted successfully."
                f"\n\nTotal time elapsed: {elapsed}",
            )
        else:
            err_lines = "\n".join(f"• {n}: {e}" for n, e in failed)
            self._uz_status_var.set(f"{success} extracted, {len(failed)} failed.")
            messagebox.showerror(
                "Unzip Errors",
                f"{success} succeeded, {len(failed)} failed:\n\n{err_lines}"
                f"\n\nTotal time elapsed: {elapsed}",
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
        tk.Checkbutton(
            opt_frame, text="Delete original files after successful conversion",
            variable=self._ai_delete_orig, bg=APP_BG, fg="#333",
            activebackground=APP_BG, font=("Segoe UI", 9), anchor="w",
        ).pack(side="left")

        # Progress + status
        self._ai_progress = ttk.Progressbar(parent, mode="determinate")
        self._ai_progress.pack(fill="x", padx=18, pady=(6, 2))
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
        self._ai_progress["value"] = 0
        self._ai_progress["maximum"] = len(self._ai_files)
        self._ai_timer.start()
        threading.Thread(target=self._aiprep_worker, daemon=True).start()

    def _aiprep_worker(self):
        files = list(self._ai_files)
        delete_orig = self._ai_delete_orig.get()
        total = len(files)
        converted = 0
        failed = []

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
            self.root.after(0, lambda s=src, n=i: self._ai_status_var.set(
                f"Converting: {s.name}  ({n + 1}/{total})"))
            self._ai_timer.set_file(src.name)
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
                        try:
                            src.unlink()
                        except Exception:
                            pass
                else:
                    failed.append(src.name)
            except Exception as exc:
                failed.append(f"{src.name}: {exc}")
            self.root.after(0, lambda: self._ai_progress.step(1))

        self.root.after(0, lambda: self._ai_finish(converted, failed))

    def _ai_finish(self, converted, failed):
        self._ai_timer.stop()
        elapsed = self._ai_timer.elapsed_str()
        self._ai_timer.reset()
        self._ai_btn.config(state="normal")
        self._set_cancel_state(self._ai_cancel_btn, False)
        if self._cancel_ai.is_set():
            self._ai_status_var.set(f"⏹  Cancelled. {converted} file(s) converted before stopping.")
            messagebox.showinfo(
                "AI Preparation Cancelled",
                f"Conversion cancelled.\n{converted} file(s) converted before stopping."
                f"\n\nTotal time elapsed: {elapsed}",
            )
            return
        if not failed:
            self._ai_status_var.set(f"✅  Done! {converted} file(s) converted to Markdown.")
            messagebox.showinfo(
                "AI Preparation Complete",
                f"{converted} file(s) converted to Markdown (.md)."
                f"\n\nTotal time elapsed: {elapsed}",
            )
        else:
            err_lines = "\n".join(f"• {n}" for n in failed)
            self._ai_status_var.set(f"⚠️  {converted} converted, {len(failed)} failed.")
            messagebox.showerror(
                "AI Preparation — Some Files Failed",
                f"{converted} converted, {len(failed)} failed:\n\n{err_lines}"
                f"\n\nTotal time elapsed: {elapsed}",
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
        tk.Checkbutton(
            opt_frame, text="Delete original files after successful conversion",
            variable=self._office_delete_orig, bg=APP_BG, fg="#333",
            activebackground=APP_BG, font=("Segoe UI", 9), anchor="w",
        ).pack(side="left")

        # Progress + status
        self._office_progress = ttk.Progressbar(parent, mode="determinate")
        self._office_progress.pack(fill="x", padx=18, pady=(6, 2))
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
        self._office_progress["value"] = 0
        self._office_progress["maximum"] = len(self._office_files)
        self._office_timer.start()
        threading.Thread(target=self._office_worker, daemon=True).start()

    def _office_worker(self):
        files = list(self._office_files)
        delete_orig = self._office_delete_orig.get()
        total = len(files)
        converted = 0
        failed = []

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
                self._set_office_status(f"Converting: {src.name}  ({i + 1}/{total})")
                self._office_timer.set_file(src.name)
                info = LEGACY_OFFICE.get(src.suffix.lower())
                if info is None:
                    failed.append((src.name, "unsupported file type"))
                    self.root.after(0, lambda: self._office_progress.step(1))
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

                # Watchdog: if this file isn't done within OFFICE_FILE_TIMEOUT,
                # assume its dedicated instance is stuck on a hidden dialog and
                # force-kill it so the batch can skip this file and continue.
                pids = set(self._office_app_pids.get(family) or ())
                timed_out = {"v": False}
                done = threading.Event()

                def _watchdog(pids=pids, timed_out=timed_out, done=done):
                    if not pids:
                        return   # nothing we can safely kill (e.g. shared PowerPoint)
                    if not done.wait(OFFICE_FILE_TIMEOUT):
                        timed_out["v"] = True
                        for p in pids:
                            _terminate_pid(p)

                wd = threading.Thread(target=_watchdog, daemon=True)
                wd.start()
                try:
                    _modernize_office_file(src, out_path, family, fmt, apps[family])
                    converted += 1
                    if delete_orig:
                        try:
                            src.unlink()
                        except Exception:
                            pass
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
                    elif timed_out["v"]:
                        failed.append((src.name, "timed out — appeared stuck, skipped"))
                    else:
                        failed.append((src.name, str(exc)))
                finally:
                    done.set()
                    wd.join(timeout=2)

                # If the watchdog killed the instance, drop it so the next file
                # of this family gets a fresh, healthy one.
                if timed_out["v"]:
                    apps[family] = None
                    ensured[family] = False
                    self._office_app_pids.pop(family, None)

                self.root.after(0, lambda: self._office_progress.step(1))
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

        self.root.after(0, lambda: self._office_finish(converted, failed))

    def _set_office_status(self, msg):
        self.root.after(0, lambda: self._office_status_var.set(msg))

    def _office_finish(self, converted, failed):
        self._office_timer.stop()
        elapsed = self._office_timer.elapsed_str()
        self._office_timer.reset()
        self._office_btn.config(state="normal")
        self._set_cancel_state(self._office_cancel_btn, False)
        if self._cancel_office.is_set():
            self._office_status_var.set(
                f"⏹  Cancelled. {converted} file(s) modernized before stopping.")
            messagebox.showinfo(
                "Office Modernizer Cancelled",
                f"Conversion cancelled.\n{converted} file(s) modernized before stopping."
                f"\n\nTotal time elapsed: {elapsed}",
            )
            return
        if not failed:
            self._office_status_var.set(f"✅  Done! {converted} file(s) modernized.")
            messagebox.showinfo(
                "Office Modernizer Complete",
                f"{converted} file(s) converted to their modern Office formats."
                f"\n\nTotal time elapsed: {elapsed}",
            )
        else:
            err_lines = "\n".join(f"• {n}: {e}" for n, e in failed)
            self._office_status_var.set(f"⚠️  {converted} modernized, {len(failed)} failed.")
            messagebox.showerror(
                "Office Modernizer — Some Files Failed",
                f"{converted} modernized, {len(failed)} failed:\n\n{err_lines}"
                f"\n\nTotal time elapsed: {elapsed}",
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
        self._pwd_progress["value"] = 0
        self._pwd_progress["maximum"] = len(self._pwd_files)
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
        tk.Checkbutton(dlg, text="Show password", variable=show_var,
                       command=_toggle, bg=APP_BG, fg="#555",
                       activebackground=APP_BG, font=("Segoe UI", 8)).pack(
            anchor="w", padx=14)

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
                self.root.after(0, lambda: self._pwd_progress.step(1))
                continue

            # Skip files that aren't actually encrypted — no point prompting.
            if not _office_is_encrypted(src):
                failed.append((src.name, "not password-protected (no open-password encryption)"))
                self.root.after(0, lambda: self._pwd_progress.step(1))
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

            self.root.after(0, lambda: self._pwd_progress.step(1))

        self.root.after(0, lambda: self._pwd_finish(unlocked, failed, skipped))

    def _set_pwd_status(self, msg):
        self.root.after(0, lambda: self._pwd_status_var.set(msg))

    def _pwd_finish(self, unlocked, failed, skipped):
        self._pwd_timer.stop()
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
        return (".eml", ".msg")

    def _on_mode_change(self, event=None):
        # Clear files when mode changes — they may be wrong type
        self._clear_files()
        self._update_drop_label()

    # ── file management ──────────────────────
    def _add_files(self):
        m = self._mode.get()
        if m.startswith("HTML"):
            filetypes = [("HTML / MHT files", "*.html *.htm *.mht *.mhtml"), ("All files", "*.*")]
            title = "Select HTML / MHT files"
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
            if m.startswith("HTML"):         label = "HTML / MHT"
            elif m.startswith("Visio"):      label = "Visio"
            elif m.startswith("Excel"):      label = "Excel / CSV"
            elif m.startswith("Word"):       label = "Word / text"
            elif m.startswith("PowerPoint"): label = "PowerPoint"
            elif m.startswith("Image"):      label = "image"
            else:                            label = "email"
            messagebox.showwarning("No Files", f"Please add {label} files first.")
            return
        self._cancel_convert.clear()
        self.convert_btn.config(state="disabled")
        self._set_cancel_state(self.cancel_btn, True)
        self.progress["value"] = 0
        self.progress["maximum"] = len(self.files)
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
        email_mode  = not any([html_mode, visio_mode, excel_mode, word_mode, ppt_mode, image_mode])

        pythoncom.CoInitialize()
        self._pdf_app_pids = []
        outlook = word = visio = excel = powerpoint = None
        success, failed = 0, []

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

            # Word — email, HTML, Word doc, and image modes
            if email_mode or html_mode or word_mode or image_mode:
                word = _launch("word")
                if word is None and (html_mode or word_mode or image_mode):
                    if html_mode:    label = "HTML / MHT"
                    elif image_mode: label = "Image"
                    else:            label = "Word document"
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
            else:               mode_key = "email"

            for i, src in enumerate(self.files):
                if self._cancel_convert.is_set():
                    break
                self._set_status(f"Converting: {src.name}  ({i + 1}/{total})")
                self._pdf_timer.set_file(src.name)
                out_dir = self._out_dir if self._out_dir else src.parent
                out_path = _unique_pdf_path(out_dir, src.stem)
                try:
                    convert_file(src, out_dir, outlook, word, visio, excel, powerpoint, mode_key)
                    success += 1
                except Exception as exc:
                    # Drop any partial PDF a force-kill may have left behind.
                    try:
                        if out_path.exists():
                            out_path.unlink()
                    except Exception:
                        pass
                    if not self._cancel_convert.is_set():
                        failed.append((src.name, str(exc)))
                self.root.after(0, lambda: self.progress.step(1))
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

        self.root.after(0, lambda: self._finish(success, failed))

    def _set_status(self, msg):
        self.root.after(0, lambda: self.status_var.set(msg))

    def _finish(self, success, failed):
        self._pdf_timer.stop()
        elapsed = self._pdf_timer.elapsed_str()
        self._pdf_timer.reset()
        self.convert_btn.config(state="normal")
        self._set_cancel_state(self.cancel_btn, False)
        if self._cancel_convert.is_set():
            self.status_var.set(f"⏹  Cancelled. {success} file(s) converted before stopping.")
            messagebox.showinfo(
                "Conversion Cancelled",
                f"Conversion was cancelled.\n{success} file(s) converted before stopping."
                f"\n\nTotal time elapsed: {elapsed}",
            )
        elif not failed:
            self.status_var.set(f"✅  Done! {success} file(s) converted successfully.")
            messagebox.showinfo(
                "Conversion Complete",
                f"{success} file(s) converted to PDF successfully."
                f"\n\nTotal time elapsed: {elapsed}",
            )
        else:
            err_lines = "\n".join(f"• {n}: {e}" for n, e in failed)
            self.status_var.set(f"⚠️  {success} converted, {len(failed)} failed.")
            messagebox.showerror(
                "Conversion Errors",
                f"{success} succeeded, {len(failed)} failed:\n\n{err_lines}"
                f"\n\nTotal time elapsed: {elapsed}",
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
