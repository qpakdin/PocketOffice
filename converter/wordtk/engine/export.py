# -*- coding: utf-8 -*-
"""PDF conversion and printing of DOCX files (uses what is installed: LibreOffice, or MS Word on Windows)."""
import os
import shutil
import subprocess
import sys
from pathlib import Path


class ExportError(RuntimeError):
    pass


def _soffice():
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(p):
            return p
    return None


def _word_com():
    if not sys.platform.startswith("win"):
        return None
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return None


def available_backend():
    """'word' | 'libreoffice' | None"""
    if _word_com():
        return "word"
    return "libreoffice" if _soffice() else None


def docx_to_pdf(src, out_dir=None, backend=None):
    """Convert one .docx to PDF next to it (or into ``out_dir``). Returns the PDF path."""
    src = Path(src).resolve()
    out_dir = Path(out_dir).resolve() if out_dir else src.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    dst = out_dir / (src.stem + ".pdf")
    backend = backend or available_backend()
    if backend == "word":
        import win32com.client
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        try:
            doc = word.Documents.Open(str(src), ReadOnly=True)
            doc.SaveAs(str(dst), FileFormat=17)         # wdFormatPDF
            doc.Close(False)
        finally:
            word.Quit()
    elif backend == "libreoffice":
        exe = _soffice()
        r = subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(src)],
                           capture_output=True, text=True, timeout=300)
        if r.returncode != 0:
            raise ExportError("LibreOffice failed: %s" % (r.stderr or r.stdout).strip()[:300])
    else:
        raise ExportError("No PDF converter found. Install LibreOffice (free) or use Windows with MS Word + pywin32.")
    if not dst.exists():
        raise ExportError("PDF was not created for %s" % src.name)
    return dst


def print_file(path, printer=None):
    """Send a .docx/.pdf to the printer. Windows: shell 'print' verb (default app) ; else CUPS 'lp'."""
    path = Path(path).resolve()
    if sys.platform.startswith("win"):
        if printer:
            raise ExportError("Choosing a printer is not supported by the Windows shell verb; set the default printer.")
        os.startfile(str(path), "print")                # noqa: S606 (Windows only)
        return "sent to the default printer"
    lp = shutil.which("lp")
    if not lp:
        raise ExportError("No 'lp' command found; cannot print from this system.")
    target = path
    if path.suffix.lower() != ".pdf":
        import tempfile
        target = docx_to_pdf(path, tempfile.mkdtemp(prefix="wordtk_print_"))
    cmd = [lp] + (["-d", printer] if printer else []) + [str(target)]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        raise ExportError(r.stderr.strip() or "lp failed")
    return r.stdout.strip() or "sent to the printer"
