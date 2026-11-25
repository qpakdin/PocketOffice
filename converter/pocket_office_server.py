#!/usr/bin/env python3
"""
Pocket Office converter server.

Turns .docx / .xlsx / .pptx into PDF with LibreOffice (a real page-layout engine) for the
Pocket Office Android app. Runs best on your Windows PC: LibreOffice there automatically uses
EVERY font installed in Windows, so nothing has to be copied to the phone.

  Windows :  python pocket_office_server.py --lan
  Linux / Termux / macOS : python3 pocket_office_server.py [--lan] [--fonts /path/to/extra/fonts]

Endpoints (used by the app)
  GET  /ping             server info (the app's "Test" button)
  GET  /fonts            list of font families LibreOffice can use on this machine
  POST /convert?ext=docx office file -> PDF  (header X-PO-Missing-Fonts lists fonts the file uses that are not installed)
  POST /text?ext=docx    office file -> plain text in logical reading order (good for Persian / RTL)
  POST /pdftext          PDF -> text (needs pdftotext, PyMuPDF or pypdf)

Security: by default only loopback and private-network clients (192.168.x.x, 10.x.x.x, 172.16-31.x.x) are
served, and the server listens on 127.0.0.1 unless you pass --lan / --host. There is no password.
"""
import argparse, hashlib, html, ipaddress, json, os, platform, posixpath, re, shutil, socket, struct, subprocess, sys, tempfile, threading, time, zipfile
import xml.etree.ElementTree as ET
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs, quote
import uuid

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
try:                                           # optional: the editors live in po_docx_edit.py / po_office_edit.py / po_new_file.py (same folder)
    import po_docx_edit as PDE
    import po_office_edit as POE
    import po_new_file as PNF
    PDE_ERROR = POE.OFFICE_EDIT_ERROR
    if PDE_ERROR:
        POE = None
except Exception as _e:  # noqa
    PDE = POE = PNF = None
    PDE_ERROR = "%s: %s" % (type(_e).__name__, _e)
EDIT_ERRORS = tuple(e for e in ((PDE.EditError if PDE else None), (POE.OfficeEditError if POE else None)) if e)
PANEL_FILES = {"index.html": "text/html; charset=utf-8", "panel.js": "application/javascript; charset=utf-8", "panel.css": "text/css; charset=utf-8"}

IS_WIN = os.name == "nt"
IS_MAC = sys.platform == "darwin"
VERSION = "4.0"

if IS_WIN:
    WORK = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "PocketOffice")
else:
    WORK = os.path.join(os.path.expanduser("~"), ".pocket_office")
PROFILE = os.path.join(WORK, "lo_profile")
FONTCONF = os.path.join(WORK, "fonts.conf")
LOCK = threading.Lock()

# ---- supported formats -------------------------------------------------------------------------
OOXML_WORD = {"docx", "docm", "dotx", "dotm"}
OOXML_SHEET = {"xlsx", "xlsm", "xltx", "xltm"}
OOXML_SLIDES = {"pptx", "pptm", "ppsx", "ppsm", "potx", "potm"}
# everything below is opened by LibreOffice itself (it recognises the format by content and extension)
OTHER_FORMATS = {"doc", "dot", "odt", "ott", "fodt", "rtf", "sxw", "stw",              # text documents
                 "xls", "xlt", "xlsb", "ods", "ots", "fods", "sxc", "csv", "tsv",      # spreadsheets
                 "ppt", "pps", "pot", "odp", "otp", "fodp", "sxi",                     # presentations
                 "html", "htm"}
SUPPORTED_EXT = OOXML_WORD | OOXML_SHEET | OOXML_SLIDES | OTHER_FORMATS


def ooxml_kind(ext):
    return "docx" if ext in OOXML_WORD else "xlsx" if ext in OOXML_SHEET else "pptx" if ext in OOXML_SLIDES else None


def clean_ext(raw):
    e = (raw or "").strip().lower()
    return e if re.fullmatch(r"[a-z0-9]{2,5}", e) else ""


CFG = {"soffice": None, "library": [], "fonts_all": False, "allow_any": False}
INDEX_FILE = os.path.join(WORK, "font_index.json")
FONTSETS = os.path.join(WORK, "fontsets")


# ----------------------------------------------------------------------------- LibreOffice / fonts

def find_soffice(explicit=None):
    cands = []
    if explicit:
        cands.append(explicit)
    if os.environ.get("PO_SOFFICE"):
        cands.append(os.environ["PO_SOFFICE"])
    if IS_WIN:
        for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)"),
                     r"C:\Program Files", r"C:\Program Files (x86)"):
            if base:
                for exe in ("soffice.com", "soffice.exe"):   # .com is the console version that waits until finished
                    cands.append(os.path.join(base, "LibreOffice", "program", exe))
    for name in (("soffice.com", "soffice", "libreoffice") if IS_WIN else ("soffice", "libreoffice")):
        p = shutil.which(name)
        if p:
            cands.append(p)
    if IS_MAC:
        cands.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    for c in cands:
        if c and os.path.isfile(c):
            return c
    return None


def prepare():
    os.makedirs(WORK, exist_ok=True)
    os.makedirs(PROFILE, exist_ok=True)
    os.makedirs(FONTSETS, exist_ok=True)


def write_conf(path, extra_dirs):
    """fontconfig file = the machine's normal fonts + the given extra folders."""
    dirs = "".join("  <dir>%s</dir>\n" % html.escape(d) for d in extra_dirs)
    with open(path, "w", encoding="utf-8") as f:
        f.write('<?xml version="1.0"?>\n<!DOCTYPE fontconfig SYSTEM "fonts.dtd">\n<fontconfig>\n'
                '  <include ignore_missing="yes">/etc/fonts/fonts.conf</include>\n%s'
                '  <cachedir>%s</cachedir>\n</fontconfig>\n' % (dirs, html.escape(os.path.join(WORK, "fccache"))))


def tool_env(conf=None):
    e = dict(os.environ)
    if conf and not IS_WIN and not IS_MAC:
        e["FONTCONFIG_FILE"] = conf
    return e


def norm_variants(raw):
    """Lower-case family-name variants for matching (strips '(TrueType)' and style words)."""
    s = re.sub(r"\s*\((truetype|opentype|all res|type 1)\)\s*$", "", (raw or "").strip(), flags=re.I)
    out = set()
    for part in s.split(" & "):
        p = part.strip().lower()
        if not p:
            continue
        out.add(p)
        for suf in (" bold italic", " bold oblique", " bold", " italic", " oblique", " regular"):
            if p.endswith(suf):
                out.add(p[: -len(suf)].strip())
    return out


_font_cache = {"t": 0.0, "names": set()}


def installed_fonts(force=False):
    """Font families available to LibreOffice on this machine (empty set = could not be determined)."""
    if not force and time.time() - _font_cache["t"] < 60 and _font_cache["names"]:
        return _font_cache["names"]
    names = set()
    try:
        if IS_WIN:
            import winreg
            for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                try:
                    with winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts") as k:
                        i = 0
                        while True:
                            try:
                                n = winreg.EnumValue(k, i)[0]
                            except OSError:
                                break
                            i += 1
                            names |= norm_variants(n)
                except OSError:
                    pass
        else:
            fc = shutil.which("fc-list")
            if fc:
                out = subprocess.run([fc, ":", "family"], env=tool_env(), capture_output=True, text=True, timeout=60).stdout
                for line in out.splitlines():
                    for fam in line.split(","):
                        names |= norm_variants(fam)
    except Exception:
        pass
    _font_cache["t"], _font_cache["names"] = time.time(), names
    return names



# ----------------------------------------------------------------------------- font library (load only what a document needs)

FONT_EXT = (".ttf", ".otf", ".ttc", ".otc")
_index = {"files": {}, "dirs": {}}
_fam_map = {}
_index_lock = threading.Lock()


def _parse_name_table(data):
    out = set()
    if len(data) < 6:
        return out
    count, str_off = struct.unpack(">HH", data[2:6])
    for i in range(count):
        r = data[6 + 12 * i: 18 + 12 * i]
        if len(r) < 12:
            break
        plat, _enc, _lang, nid, ln, off = struct.unpack(">6H", r)
        if nid not in (1, 4, 16):          # family, full name, typographic family
            continue
        raw = data[str_off + off: str_off + off + ln]
        try:
            txt = raw.decode("utf-16-be") if plat in (0, 3) else raw.decode("latin-1")
        except Exception:
            continue
        txt = txt.strip().lower()
        if txt:
            out.add(txt)
    return out


def read_font_names(path):
    """Family names inside a TTF / OTF / TTC file (reads only the small 'name' table)."""
    names = set()
    with open(path, "rb") as f:
        head = f.read(12)
        if len(head) < 12:
            return names
        offsets = [0]
        if head[:4] == b"ttcf":
            n = min(struct.unpack(">I", head[8:12])[0], 64)
            offsets = list(struct.unpack(">%dI" % n, f.read(4 * n)))
        for off in offsets:
            f.seek(off)
            h = f.read(12)
            if len(h) < 12 or h[:4] not in (b"\x00\x01\x00\x00", b"OTTO", b"true", b"typ1"):
                continue
            num = struct.unpack(">H", h[4:6])[0]
            tabdir = f.read(16 * num)
            for i in range(num):
                rec = tabdir[i * 16:(i + 1) * 16]
                if rec[:4] == b"name":
                    toff, tlen = struct.unpack(">II", rec[8:16])
                    f.seek(toff)
                    names |= _parse_name_table(f.read(min(tlen, 262144)))
                    break
    return names


def refresh_index(force=False, verbose=False):
    """Keep a cached list 'font file -> family names'. Only new or changed files are read."""
    global _fam_map
    if not CFG["library"]:
        return
    with _index_lock:
        if not _index["files"] and os.path.exists(INDEX_FILE):
            try:
                with open(INDEX_FILE, "r", encoding="utf-8") as f:
                    _index.update(json.load(f))
            except Exception:
                pass
        listing, cur_dirs, changed = {}, {}, force
        for lib in CFG["library"]:
            for root, _d, fnames in os.walk(lib):
                try:
                    cur_dirs[root] = os.stat(root).st_mtime
                except OSError:
                    continue
                listing[root] = [n for n in fnames if n.lower().endswith(FONT_EXT)]
        if cur_dirs != _index["dirs"] or not _fam_map:
            changed = True
        if not changed:
            return
        files, new_files = _index["files"], {}
        todo = [os.path.join(r, n) for r, names in listing.items() for n in names]
        done = 0
        for path in todo:
            done += 1
            try:
                st = os.stat(path)
            except OSError:
                continue
            old = files.get(path)
            if old and old["m"] == st.st_mtime and old["s"] == st.st_size:
                new_files[path] = old
            else:
                try:
                    new_files[path] = {"m": st.st_mtime, "s": st.st_size, "f": sorted(read_font_names(path))}
                except Exception:
                    new_files[path] = {"m": st.st_mtime, "s": st.st_size, "f": []}
                if verbose and done % 200 == 0:
                    print("  indexing fonts: %d / %d" % (done, len(todo)), flush=True)
        _index["files"], _index["dirs"] = new_files, cur_dirs
        fam = {}
        for path, e in new_files.items():
            for n in e["f"]:
                fam.setdefault(n, []).append(path)
        _fam_map = fam
        try:
            with open(INDEX_FILE, "w", encoding="utf-8") as f:
                json.dump(_index, f)
        except Exception:
            pass


def library_files_for(families):
    """Font files (all styles) of the requested families."""
    refresh_index()
    paths = set()
    for fam in families:
        paths.update(_fam_map.get(fam.strip().lower(), []))
    return sorted(paths)


def build_fontset(paths):
    """A small folder holding links to just these font files + its own fontconfig file. Re-used when the same set comes again."""
    key = hashlib.sha1("\n".join(paths).encode("utf-8")).hexdigest()[:16]
    d = os.path.join(FONTSETS, key)
    conf = os.path.join(d, "fonts.conf")
    if not os.path.exists(conf):
        fdir = os.path.join(d, "fonts")
        os.makedirs(fdir, exist_ok=True)
        for i, p in enumerate(paths):
            dst = os.path.join(fdir, "%03d_%s" % (i, os.path.basename(p)))
            try:
                os.symlink(p, dst)
            except (OSError, NotImplementedError):
                shutil.copy2(p, dst)
        write_conf(conf, [fdir])
    try:
        os.utime(d, None)
        sets = sorted((os.path.getmtime(os.path.join(FONTSETS, n)), n) for n in os.listdir(FONTSETS))
        for _t, n in sets[:-40]:                      # keep the 40 most recently used sets
            shutil.rmtree(os.path.join(FONTSETS, n), ignore_errors=True)
    except OSError:
        pass
    return conf


def fonts_conf_for(src, ext):
    """-> (fontconfig file or None, fonts the document wants, fonts loaded from the library)."""
    wanted = doc_fonts(src, ext)
    if CFG["library"] and not IS_WIN and not IS_MAC:
        if CFG["fonts_all"]:
            conf = os.path.join(WORK, "fonts_all.conf")
            write_conf(conf, CFG["library"])
            return conf, wanted, []
        paths = library_files_for(wanted)
        if paths:
            return build_fontset(paths), wanted, paths
    return None, wanted, []


FONT_PATTERNS = [
    re.compile(r'w:(?:ascii|hAnsi|cs|eastAsia)="([^"]+)"'),
    re.compile(r'<w:font w:name="([^"]+)"'),
    re.compile(r'<name val="([^"]+)"'),
    re.compile(r'<a:(?:latin|ea|cs) typeface="([^"]*)"'),
]


def doc_fonts(path, ext):
    """Font families a document asks for."""
    kind = ooxml_kind(ext)
    if kind is None:
        return set()                      # binary / OpenDocument formats: font list not read, LibreOffice still converts them
    prefixes = {"docx": ("word/",), "xlsx": ("xl/styles.xml", "xl/theme/"), "pptx": ("ppt/slides/", "ppt/slideLayouts/", "ppt/slideMasters/", "ppt/theme/")}[kind]
    found = set()
    try:
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                if n.endswith(".xml") and n.startswith(prefixes):
                    txt = z.read(n).decode("utf-8", "ignore")
                    for pat in FONT_PATTERNS:
                        for m in pat.findall(txt):
                            m = html.unescape(m).strip()
                            if m and not m.startswith("+") and len(m) < 60:
                                found.add(m)
    except Exception:
        pass
    return found


def missing_fonts(wanted):
    """Fonts the document asks for that exist neither on this machine nor in the font library."""
    names = installed_fonts()
    if not names and not _fam_map:
        return None
    return sorted(f for f in wanted if f.lower() not in names and f.lower() not in _fam_map)


# ----------------------------------------------------------------------------- text extraction (logical order)

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PML = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
SML = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _para_text(p):
    parts = []

    def rec(e):
        for c in e:
            t = c.tag
            if t == W + "t":
                parts.append(c.text or "")
            elif t == W + "tab":
                parts.append("\t")
            elif t in (W + "br", W + "cr"):
                parts.append("\n")
            elif t == W + "noBreakHyphen":
                parts.append("-")
            elif t in (W + "txbxContent", MC + "Fallback", W + "instrText", W + "delText"):
                continue
            else:
                rec(c)
    rec(p)
    return "".join(parts)


def _textboxes(e):
    for c in e:
        if c.tag == MC + "Fallback":
            continue
        if c.tag == W + "txbxContent":
            yield c
        else:
            yield from _textboxes(c)


def _walk_block(parent, lines):
    for c in parent:
        t = c.tag
        if t == W + "p":
            lines.append(_para_text(c))
            for tb in _textboxes(c):
                _walk_block(tb, lines)
        elif t == W + "tbl":
            for tr in c.findall(W + "tr"):
                cells = []
                for tc in tr.findall(W + "tc"):
                    cl = []
                    _walk_block(tc, cl)
                    cells.append(" ".join(x for x in cl if x))
                lines.append("\t".join(cells))
        elif t == W + "sdt":
            sc = c.find(W + "sdtContent")
            if sc is not None:
                _walk_block(sc, lines)
        elif t == MC + "AlternateContent":
            ch = c.find(MC + "Choice")
            if ch is not None:
                _walk_block(ch, lines)


def docx_text(path):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    lines = []
    body = root.find(W + "body")
    _walk_block(body if body is not None else root, lines)
    return "\n".join(lines)


def _col_index(ref):
    n = 0
    for ch in re.match(r"[A-Za-z]+", ref).group(0).upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _part_path(base_dir, target):
    return target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base_dir, target))


def xlsx_text(path):
    out = []
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        shared = []
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(SML + "si"):
                shared.append("".join(t.text or "" for t in si.iter(SML + "t")))
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
        for sh in wb.find(SML + "sheets"):
            if sh.get("state") in ("hidden", "veryHidden"):
                continue
            target = rels.get(sh.get(REL + "id"))
            if not target:
                continue
            p = _part_path("xl", target)
            if p not in names:
                continue
            out.append("=== %s ===" % sh.get("name"))
            for row in ET.fromstring(z.read(p)).iter(SML + "row"):
                cells = {}
                for c in row.findall(SML + "c"):
                    t, v = c.get("t"), c.find(SML + "v")
                    if t == "s":
                        val = shared[int(v.text)] if v is not None and v.text and int(v.text) < len(shared) else ""
                    elif t == "inlineStr":
                        val = "".join(x.text or "" for x in c.iter(SML + "t"))
                    elif t == "b":
                        val = "TRUE" if (v is not None and v.text == "1") else "FALSE"
                    else:
                        val = v.text if (v is not None and v.text is not None) else ""
                    if c.get("r"):
                        cells[_col_index(c.get("r"))] = val
                if cells:
                    out.append("\t".join(cells.get(i, "") for i in range(1, max(cells) + 1)).rstrip("\t"))
            out.append("")
    return "\n".join(out)


def pptx_text(path):
    out = []
    with zipfile.ZipFile(path) as z:
        pres = ET.fromstring(z.read("ppt/presentation.xml"))
        rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
        lst = pres.find(PML + "sldIdLst")
        for i, sid in enumerate(list(lst) if lst is not None else [], 1):
            target = rels.get(sid.get(REL + "id"))
            if not target:
                continue
            out.append("--- Slide %d ---" % i)
            for p in ET.fromstring(z.read(_part_path("ppt", target))).iter(A + "p"):
                s = []
                for c in p:
                    if c.tag in (A + "r", A + "fld"):
                        s.extend(t.text or "" for t in c.findall(A + "t"))
                    elif c.tag == A + "br":
                        s.append("\n")
                line = "".join(s)
                if line.strip():
                    out.append(line)
            out.append("")
    return "\n".join(out)


def office_text(path, ext):
    kind = ooxml_kind(ext)
    if kind is None:
        raise ValueError("no direct text reader for .%s - use /pdftext" % ext)
    return {"docx": docx_text, "xlsx": xlsx_text, "pptx": pptx_text}[kind](path)


def pdf_to_text(pdf_path, workdir):
    exe = shutil.which("pdftotext")
    if exe:
        out = os.path.join(workdir, "out.txt")
        subprocess.run([exe, "-enc", "UTF-8", "-layout", pdf_path, out], capture_output=True, timeout=300)
        if os.path.exists(out):
            with open(out, "rb") as f:
                return f.read().decode("utf-8", "replace")
    try:
        import fitz  # PyMuPDF
        with fitz.open(pdf_path) as d:
            return "\n".join(p.get_text() for p in d)
    except Exception:
        pass
    try:
        from pypdf import PdfReader
        return "\n".join((p.extract_text() or "") for p in PdfReader(pdf_path).pages)
    except Exception:
        pass
    raise RuntimeError("no PDF text tool found (install poppler's pdftotext, or: pip install pymupdf)")


# ----------------------------------------------------------------------------- HTTP


def convert_file_to_pdf(src, ext, d, target="pdf"):
    """Convert one file with LibreOffice (one at a time). -> (output_path, fonts_wanted, loaded_files, missing_fonts)"""
    if not CFG["soffice"]:
        raise RuntimeError("LibreOffice was not found on the server machine (install it, or pass --soffice PATH)")
    conf, wanted, loaded = fonts_conf_for(src, ext)
    miss = missing_fonts(wanted)
    with LOCK:
        cmd = [CFG["soffice"], "--headless", "--norestore", "--nolockcheck", "--nodefault", "--nologo",
               "-env:UserInstallation=" + Path(PROFILE).as_uri(),
               "--convert-to", target, "--outdir", d, src]
        r = subprocess.run(cmd, env=tool_env(conf), capture_output=True, timeout=900)
        out = os.path.splitext(src)[0] + "." + target
        t0, last = time.time(), -1
        while time.time() - t0 < 40:               # the launcher can return before the file is complete
            if os.path.exists(out):
                size = os.path.getsize(out)
                if size > 0 and size == last:
                    break
                last = size
            time.sleep(0.6)
    if not os.path.exists(out) or os.path.getsize(out) == 0:
        tail = ((r.stderr or b"") + (r.stdout or b"")).decode("utf-8", "replace")[-500:]
        raise RuntimeError("conversion failed. " + tail)
    return out, wanted, loaded, miss


def count_pdf_pages(path):
    exe = shutil.which("pdfinfo")
    if exe:
        try:
            m = re.search(r"Pages:\s+(\d+)", subprocess.run([exe, path], capture_output=True, text=True, timeout=60).stdout)
            if m:
                return int(m.group(1))
        except Exception:
            pass
    with open(path, "rb") as f:
        return len(re.findall(rb"/Type\s*/Page(?![s\w])", f.read()))


# ----------------------------------------------------------------------------- Word editing sessions (copy only, everything is logged)
SESSIONS = {}
EDIT_LOG = os.path.join(WORK, "edit_log.txt")


def elog(msg):
    line = "[EDIT %s] %s" % (time.strftime("%Y-%m-%d %H:%M:%S"), msg)
    print(line, flush=True)
    try:
        os.makedirs(WORK, exist_ok=True)
        with open(EDIT_LOG, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass


def get_session(sid):
    ent = SESSIONS.get(sid)
    if not ent:
        raise KeyError("editing session not found (the server was restarted or the session expired) - open the file for editing again")
    ent["t"] = time.time()
    return ent["s"]


def session_meta(sid):
    return SESSIONS[sid]["meta"]


RENDERS = {}


def drop_session(sid):
    SESSIONS.pop(sid, None)
    r = RENDERS.pop(sid, None)
    if r:
        shutil.rmtree(r["dir"], ignore_errors=True)


def prune_sessions():
    now = time.time()
    for k in [k for k, v in SESSIONS.items() if now - v["t"] > 6 * 3600]:
        drop_session(k)
    while len(SESSIONS) > 12:
        drop_session(min(SESSIONS, key=lambda k: SESSIONS[k]["t"]))


def describe_op(op):
    k = op.get("op")
    if k == "set_text":
        return "paragraph %s: text replaced (new length %d characters)" % (op.get("p"), len(str(op.get("text", ""))))
    if k == "format":
        props = ", ".join("%s=%s" % (a, op[a]) for a in ("font", "size", "bold", "italic", "underline", "strike", "color", "highlight") if op.get(a) is not None)
        rng = "whole paragraph" if op.get("start") is None else "characters %s-%s" % (op.get("start"), op.get("end"))
        return "paragraph %s: font formatting (%s) on %s" % (op.get("p"), props or "-", rng)
    if k == "para":
        props = ", ".join("%s=%s" % (a, op[a]) for a in ("align", "rtl", "before", "after", "line", "left", "right", "first") if op.get(a) is not None)
        return "paragraph %s: paragraph settings (%s)" % (op.get("p"), props or "-")
    if k == "page":
        return "page setup (%s): %.2f x %.2f cm, margins T%.2f B%.2f L%.2f R%.2f cm" % (op.get("scope", "last"), float(op.get("w_cm", 0)), float(op.get("h_cm", 0)),
                                                                                   float(op.get("top", 0)), float(op.get("bottom", 0)), float(op.get("left", 0)), float(op.get("right", 0)))
    if k == "set_cell":
        return "cell %s!%s set to %s" % (op.get("sheet"), op.get("ref"), (str(op.get("value", ""))[:40] + ("..." if len(str(op.get("value", ""))) > 40 else "")) or "(empty)")
    return "%s %s" % (k, {a: op[a] for a in ("p", "t", "r", "where") if a in op})


def glob_first(prefix):
    import glob
    hits = sorted(glob.glob(prefix + "*.png"))
    return hits[0] if hits else None



def install_hint():
    """Exact command that fixes a missing 'lxml' on the machine the server runs on."""
    if IS_WIN:
        return "pip install lxml"
    if os.environ.get("PREFIX", "").startswith("/data/data/com.termux"):
        return "pkg install python-lxml     (or, inside Debian proot: apt install python3-lxml)"
    if os.path.exists("/etc/debian_version"):
        return "apt install -y python3-lxml"
    return "pip install lxml"


def editors_error_text():
    """Readable reason why the editors are off (shown in the panel, in the app and in the banner)."""
    e = PDE_ERROR or ""
    if "lxml" in e:
        return "The editors need the Python module 'lxml', which is not installed for the Python that runs this server.\nFix (run on the machine where the server runs, then restart the server):\n    %s\nTo check everything at once: python3 po_doctor.py" % install_hint()
    return "The editors are not available (%s). Copy the whole 'converter' folder (po_docx_edit.py, po_office_edit.py, po_new_file.py, wordtk/, panel/) next to the server and run: python3 po_doctor.py" % e


def is_allowed(addr):
    if CFG["allow_any"]:
        return True
    try:
        ip = ipaddress.ip_address(addr.split("%")[0])
        if getattr(ip, "ipv4_mapped", None):
            ip = ip.ipv4_mapped
        return ip.is_loopback or ip.is_private or ip.is_link_local
    except ValueError:
        return False


class Handler(BaseHTTPRequestHandler):
    server_version = "PocketOffice/" + VERSION

    def log_message(self, fmt, *args):
        sys.stderr.write("[%s] %s\n" % (self.log_date_time_string(), fmt % args))

    def reply(self, code, body, ctype="text/plain; charset=utf-8", headers=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def read_body_to(self, path):
        left = int(self.headers.get("Content-Length", "0"))
        if left <= 0:
            raise ValueError("empty request body")
        with open(path, "wb") as f:
            while left > 0:
                chunk = self.rfile.read(min(1 << 20, left))
                if not chunk:
                    raise ValueError("connection closed early")
                f.write(chunk)
                left -= len(chunk)


    def jreply(self, obj, code=200, headers=None):
        self.reply(code, json.dumps(obj, ensure_ascii=False), "application/json; charset=utf-8", headers)

    def _edit_guard(self):
        if POE is None:
            self.jreply({"ok": False, "error": editors_error_text()}, 501)
            return False
        return True

    def _open_payload(self, sid, sess, meta):
        out = {"ok": True, "sid": sid, "kind": sess.kind, "name": SESSIONS[sid]["name"], "meta": meta, "report": sess.report, "blocks": sess.blocks(),
               "tables": sess.table_rows(), "page": sess.page_info(), "ops": len(sess.history) if hasattr(sess, "history") else len(sess.ops)}
        if sess.kind == "xlsx":
            out["sheets"] = sess.sheet_names()
        return out

    def _register(self, sess, name, meta):
        prune_sessions()
        sid = uuid.uuid4().hex[:12]
        SESSIONS[sid] = {"s": sess, "t": time.time(), "name": name, "meta": meta}
        return sid

    def handle_edit_post(self, u, ext, d):
        if not self._edit_guard():
            return
        q = parse_qs(u.query)
        if u.path == "/edit/open":
            if ooxml_kind(ext) is None:
                return self.jreply({"ok": False, "error": "Editing is available for Word (.docx .docm .dotx .dotm), Excel (.xlsx .xlsm .xltx .xltm) and PowerPoint (.pptx .pptm .ppsx .ppsm .potx .potm). Other formats would have to be re-written by LibreOffice, which could damage features of the document, so they are view-only."}, 400)
            src = os.path.join(d, "in." + ext)
            self.read_body_to(src)
            with open(src, "rb") as f:
                data = f.read()
            name = q.get("name", ["document." + ext])[0]
            try:
                sess = POE.open_session(ext, data, name, elog)
            except EDIT_ERRORS as e:
                return self.jreply({"ok": False, "error": str(e)}, 400)
            meta = {"ext": ext, "save_ext": ext, "new": False}
            sid = self._register(sess, name, meta)
            rep = sess.report
            elog("=== EDIT PREFLIGHT  session %s | %s | file '%s' | %d bytes | sha256 %s" % (sid, sess.kind, name, len(data), rep["sha256"]))
            for it in rep["items"]:
                elog("    %-5s %s" % (it["level"], it["text"]))
            if rep.get("stats"):
                elog("    stats: %s" % json.dumps(rep["stats"]))
            elog("    the ORIGINAL file is never modified; edits go to a copy that is saved by the user")
            return self.jreply(self._open_payload(sid, sess, meta))
        if u.path == "/edit/new":
            try:
                blank, edit_ext, save_ext = PNF.create(ext)
                name = q.get("name", ["Untitled." + save_ext])[0]
                sess = POE.open_session(edit_ext, blank, name, elog)
            except (ValueError,) + EDIT_ERRORS as e:
                return self.jreply({"ok": False, "error": str(e)}, 400)
            meta = {"ext": edit_ext, "save_ext": save_ext, "new": True}
            sid = self._register(sess, name, meta)
            elog("=== NEW FILE  session %s | %s | '%s' (edited as .%s, saved as .%s)" % (sid, sess.kind, name, edit_ext, save_ext))
            return self.jreply(self._open_payload(sid, sess, meta))
        sid = q.get("sid", [""])[0]
        try:
            sess = get_session(sid)
        except KeyError as e:
            return self.jreply({"ok": False, "error": str(e.args[0])}, 404)
        src = os.path.join(d, "body.json")
        self.read_body_to(src) if int(self.headers.get("Content-Length", "0")) > 0 else open(src, "wb").write(b"{}")
        with open(src, "rb") as f:
            body = json.loads(f.read().decode("utf-8") or "{}")
        if u.path == "/edit/apply":
            ops = body.get("ops", [])
            try:
                for op in ops:
                    elog("[%s] requested: %s" % (sid, describe_op(op)))
                res = sess.apply(ops)
                elog("[%s] APPLIED %d operation(s); changes in this session: %d" % (sid, len(ops), len(sess.history)))
            except EDIT_ERRORS as e:
                elog("[%s] REFUSED, nothing was changed: %s" % (sid, e))
                return self.jreply({"ok": False, "error": str(e)})
            touched = {op.get("p") for op in ops if isinstance(op.get("p"), int)}
            allb = sess.blocks()
            out = {"ok": True, "structure_changed": res.get("structure_changed", False), "tables": sess.table_rows(), "page": sess.page_info(), "ops": len(sess.history),
                   "formula_results_cleared": res.get("formula_results_cleared")}
            out["blocks"] = allb if out["structure_changed"] else [b for b in allb if b["i"] in touched]
            return self.jreply(out)
        if u.path == "/edit/settings":
            group, dry = body.get("group", ""), bool(body.get("dry"))
            try:
                elog("[%s] settings %s: %s %s" % (sid, "PREVIEW" if dry else "APPLY", group, json.dumps(body.get("values", {}), ensure_ascii=False)[:300]))
                res = sess.settings(group, body.get("values", {}), body.get("scope"), dry)
            except EDIT_ERRORS as e:
                elog("[%s] settings REFUSED, nothing was changed: %s" % (sid, e))
                return self.jreply({"ok": False, "error": str(e)})
            elog("[%s]   -> %s | changes %s | parts changed %s | warnings %d" % (sid, "nothing to change" if res["nothing_changed"] else ("dry run" if dry else "APPLIED"), json.dumps(res["changes"]), res["parts_changed"], len(res["warnings"])))
            for w in res["warnings"][:8]:
                elog("[%s]   warning: %s" % (sid, w))
            out = {"ok": True, "result": res, "dry": dry, "ops": len(sess.history)}
            if not dry:
                out.update({"blocks": sess.blocks(), "tables": sess.table_rows(), "page": sess.page_info()})
            return self.jreply(out)
        if u.path == "/edit/undo":
            try:
                sess.undo()
                elog("[%s] UNDO; changes left: %d" % (sid, len(sess.history)))
            except EDIT_ERRORS as e:
                return self.jreply({"ok": False, "error": str(e)})
            return self.jreply({"ok": True, "blocks": sess.blocks(), "tables": sess.table_rows(), "page": sess.page_info(), "ops": len(sess.history)})
        if u.path == "/edit/close":
            drop_session(sid)
            elog("[%s] session closed (the original file was never modified)" % sid)
            return self.jreply({"ok": True})
        return self.jreply({"ok": False, "error": "unknown edit command"}, 404)

    def _render(self, sid, sess):
        """PDF of the current state (cached until the document changes). -> dict(dir, pdf, pages, key)"""
        key = hashlib.sha1(sess.cur).hexdigest()[:12]
        r = RENDERS.get(sid)
        if r and r["key"] == key:
            return r
        if r:
            shutil.rmtree(r["dir"], ignore_errors=True)
        os.makedirs(WORK, exist_ok=True)
        d = tempfile.mkdtemp(dir=WORK, prefix="render_")
        ext = session_meta(sid)["ext"]
        src = os.path.join(d, "edited." + ext)
        with open(src, "wb") as f:
            f.write(sess.export())
        pdf, _w, _l, _m = convert_file_to_pdf(src, ext, d)
        r = {"dir": d, "pdf": pdf, "pages": count_pdf_pages(pdf), "key": key}
        RENDERS[sid] = r
        return r

    def handle_edit_get(self, path, q):
        if not self._edit_guard():
            return
        if path == "/edit/sessions":
            return self.jreply({"ok": True, "sessions": [{"sid": k, "name": v["name"], "kind": v["s"].kind, "changes": len(v["s"].history), "new": v["meta"].get("new", False),
                                                          "age": int(time.time() - v["t"])} for k, v in sorted(SESSIONS.items(), key=lambda kv: -kv[1]["t"])]})
        if path == "/edit/catalog":
            return self.jreply({"ok": True, **PNF.catalog()})
        sid = q.get("sid", [""])[0]
        try:
            sess = get_session(sid)
        except KeyError as e:
            return self.jreply({"ok": False, "error": str(e.args[0])}, 404)
        meta = session_meta(sid)
        try:
            return self._edit_get(path, q, sid, sess, meta)
        except EDIT_ERRORS as e:
            return self.jreply({"ok": False, "error": str(e)})
        except RuntimeError as e:
            return self.jreply({"ok": False, "error": str(e)}, 500)

    def _edit_get(self, path, q, sid, sess, meta):
        flag = lambda k: q.get(k, ["0"])[0] in ("1", "true", "on")
        if path == "/edit/open_payload":
            return self.jreply(self._open_payload(sid, sess, meta))
        if path == "/edit/blocks":
            return self.jreply({"ok": True, "blocks": sess.blocks(), "tables": sess.table_rows(), "page": sess.page_info(), "ops": len(sess.history)})
        if path == "/edit/schema":
            return self.jreply({"ok": True, "kind": sess.kind, "groups": sess.schema()})
        if path == "/edit/find":
            return self.jreply({"ok": True, **sess.find(q.get("q", [""])[0], flag("case"), flag("word"), flag("regex"))})
        if path == "/edit/cells":
            return self.jreply({"ok": True, **sess.cells(int(q.get("sheet", ["0"])[0]), int(q.get("offset", ["0"])[0]))})
        if path == "/edit/grid":
            g = lambda k, d: int(q.get(k, [str(d)])[0])
            return self.jreply({"ok": True, **sess.grid(g("sheet", 0), max(1, g("r0", 1)), max(1, g("c0", 1)), min(200, g("rows", 40)), min(60, g("cols", 12)))})
        if path == "/edit/render":
            r = self._render(sid, sess)
            elog("[%s] preview rendered: %d page(s)" % (sid, r["pages"]))
            return self.jreply({"ok": True, "pages": r["pages"], "key": r["key"]})
        if path == "/edit/page":
            r = self._render(sid, sess)
            n = max(1, min(r["pages"], int(q.get("n", ["1"])[0])))
            dpi = max(40, min(200, int(q.get("dpi", ["90"])[0])))
            exe = shutil.which("pdftoppm")
            if not exe:
                return self.reply(500, "pdftoppm is not installed (Debian: apt install poppler-utils)")
            prefix = os.path.join(r["dir"], "pg%d_%d" % (n, dpi))
            if not glob_first(prefix):
                subprocess.run([exe, "-png", "-f", str(n), "-l", str(n), "-r", str(dpi), r["pdf"], prefix], capture_output=True, timeout=120)
            f = glob_first(prefix)
            if not f:
                return self.reply(500, "could not render the page")
            with open(f, "rb") as fh:
                return self.reply(200, fh.read(), "image/png", {"Cache-Control": "max-age=60"})
        if path == "/edit/preview":
            r = self._render(sid, sess)
            with open(r["pdf"], "rb") as f:
                return self.reply(200, f.read(), "application/pdf", {"X-PO-Pages": str(r["pages"])})
        if path == "/edit/verify":
            v = sess.verify()
            try:
                r = self._render(sid, sess)
                before = int(q.get("orig_pages", ["0"])[0])
                v["checks"].append({"name": "LibreOffice can open and lay out the edited file", "ok": True, "detail": "%d page(s)%s" % (r["pages"], (" (original: %d)" % before) if before else "")})
            except RuntimeError as e:
                v["ok"] = False
                v["checks"].append({"name": "LibreOffice can open and lay out the edited file", "ok": False, "detail": str(e)[:200]})
            elog("[%s] VERIFICATION %s" % (sid, "PASSED" if v["ok"] else "FAILED"))
            for c in v["checks"]:
                elog("    %-4s %s | %s" % ("ok" if c["ok"] else "BAD", c["name"], c["detail"]))
            for ch in v["changes"][:20]:
                elog("    change: %s '%s' -> '%s'" % (ch["tag"], ch["old"][:60], ch["new"][:60]))
            return self.jreply({"ok": True, "verify": v})
        if path == "/edit/export":
            native = meta["ext"]
            target = (q.get("as", [meta["save_ext"]])[0] or native).lower()
            data = sess.export()
            if target != native:
                if not re.fullmatch(r"[a-z0-9]{2,5}", target):
                    return self.jreply({"ok": False, "error": "unknown format"}, 400)
                d = tempfile.mkdtemp(dir=WORK)
                try:
                    src = os.path.join(d, "edited." + native)
                    with open(src, "wb") as f:
                        f.write(data)
                    out, _w, _l, _m = convert_file_to_pdf(src, native, d, target)
                    with open(out, "rb") as f:
                        data = f.read()
                finally:
                    shutil.rmtree(d, ignore_errors=True)
            base = os.path.splitext(SESSIONS[sid]["name"])[0] or "document"
            fname = base + ("" if meta.get("new") else "_edited") + "." + target
            sha = hashlib.sha256(data).hexdigest()
            elog("[%s] EXPORT '%s' %d bytes, sha256 %s | original sha256 %s | %d change(s)" % (sid, fname, len(data), sha, sess.sha, len(sess.history)))
            ascii_name = re.sub(r'[^A-Za-z0-9._-]', "_", fname)
            return self.reply(200, data, "application/octet-stream", {"X-PO-SHA256": sha, "Content-Disposition": "attachment; filename=\"%s\"; filename*=UTF-8''%s" % (ascii_name, quote(fname))})
        return self.jreply({"ok": False, "error": "unknown edit command"}, 404)

    def _host_ok(self):
        """Blocks DNS-rebinding and cross-site requests: only IP addresses / localhost may be used as host name, and POSTs from a web page must come from the same origin."""
        host = self.headers.get("Host", "")
        name = host[1:host.index("]")] if host.startswith("[") and "]" in host else host.rsplit(":", 1)[0] if ":" in host else host
        if name != "localhost":
            try:
                ipaddress.ip_address(name)
            except ValueError:
                return False
        origin = self.headers.get("Origin")
        if origin and urlparse(origin).netloc != host:
            return False
        return True

    def serve_panel(self, path):
        if path in ("/", "/panel"):
            self.send_response(302)
            self.send_header("Location", "/panel/")
            self.send_header("Content-Length", "0")
            self.end_headers()
            return
        name = path[len("/panel/"):] or "index.html"
        if name not in PANEL_FILES:
            return self.reply(404, "not found")
        try:
            with open(os.path.join(HERE, "panel", name), "rb") as f:
                return self.reply(200, f.read(), PANEL_FILES[name], {"Cache-Control": "no-store"})
        except OSError:
            return self.reply(404, "the panel files are missing: copy the 'panel' folder next to this server")

    def do_GET(self):
        if not is_allowed(self.client_address[0]):
            return self.reply(403, "only private-network clients are allowed")
        if not self._host_ok():
            return self.reply(403, "forbidden host / origin")
        path = urlparse(self.path).path
        if path == "/favicon.ico":
            self.send_response(204)
            self.send_header("Content-Length", "0")
            self.end_headers()
            return
        if path == "/" or path.startswith("/panel"):
            return self.serve_panel(path)
        if path == "/ping":
            fonts = installed_fonts(force=True)
            lib = ("%d font files in library (only the fonts a document needs are loaded)" % len(_index["files"])) if CFG["library"] else "no font library"
            self.reply(200, "pocket-office-server v%s | os=%s | libreoffice=%s | system font families=%s | %s" % (
                VERSION, platform.system(), CFG["soffice"] or "NOT FOUND", len(fonts) if fonts else "unknown", lib) + " | formats=%d | editors=%s | panel=/panel/" % (len(SUPPORTED_EXT), "ready" if POE else "MISSING (%s)" % PDE_ERROR))
        elif path.startswith("/edit/"):
            self.handle_edit_get(path, parse_qs(urlparse(self.path).query))
        elif path == "/formats":
            self.reply(200, " ".join(sorted(SUPPORTED_EXT)))
        elif path == "/reindex":
            refresh_index(force=True)
            self.reply(200, "indexed %d font files" % len(_index["files"]))
        elif path == "/fonts":
            self.reply(200, "\n".join(sorted(installed_fonts(force=True))))
        else:
            self.reply(404, "not found")

    def do_POST(self):
        if not is_allowed(self.client_address[0]):
            return self.reply(403, "only private-network clients are allowed")
        if not self._host_ok():
            return self.reply(403, "forbidden host / origin")
        u = urlparse(self.path)
        ext = clean_ext(parse_qs(u.query).get("ext", ["docx"])[0])
        os.makedirs(WORK, exist_ok=True)
        d = tempfile.mkdtemp(dir=WORK)
        try:
            if u.path == "/convert":
                if ext not in SUPPORTED_EXT:
                    return self.reply(400, "unsupported type .%s" % (ext or "?"))
                if not CFG["soffice"]:
                    return self.reply(500, "LibreOffice was not found on the server machine (install it, or pass --soffice PATH)")
                src = os.path.join(d, "in." + ext)
                self.read_body_to(src)
                try:
                    out, wanted, loaded, miss = convert_file_to_pdf(src, ext, d)
                except RuntimeError as e:
                    return self.reply(500, str(e))
                hdr = {}
                if miss:
                    hdr["X-PO-Missing-Fonts"] = quote(",".join(miss[:25]), safe=",")
                self.log_message("fonts wanted: %s | loaded from library: %d | missing: %s", ", ".join(sorted(wanted)) or "-", len(loaded), ", ".join(miss or []) or "-")
                with open(out, "rb") as f:
                    return self.reply(200, f.read(), "application/pdf", hdr)
            elif u.path == "/text":
                if ooxml_kind(ext) is None:
                    return self.reply(400, "direct text reading only for OOXML files; use /pdftext")
                src = os.path.join(d, "in." + ext)
                self.read_body_to(src)
                return self.reply(200, office_text(src, ext))
            elif u.path == "/pdftext":
                src = os.path.join(d, "in.pdf")
                self.read_body_to(src)
                return self.reply(200, pdf_to_text(src, d))
            elif u.path.startswith("/edit/"):
                return self.handle_edit_post(u, ext, d)
            else:
                return self.reply(404, "not found")
        except subprocess.TimeoutExpired:
            self.reply(500, "conversion timed out")
        except Exception as e:  # noqa
            try:
                self.reply(500, "error: %s" % e)
            except Exception:
                pass
        finally:
            shutil.rmtree(d, ignore_errors=True)


def local_ips():
    ips = set()
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ips.add(info[4][0])
    except Exception:
        pass
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("10.255.255.255", 1))
        ips.add(s.getsockname()[0])
        s.close()
    except Exception:
        pass
    return sorted(i for i in ips if not i.startswith("127."))


def main():
    ap = argparse.ArgumentParser(description="Pocket Office converter server")
    ap.add_argument("--host", default=os.environ.get("PO_HOST", "127.0.0.1"), help="listen address (default 127.0.0.1)")
    ap.add_argument("--lan", action="store_true", help="listen on all interfaces so your phone can reach this machine")
    ap.add_argument("--port", type=int, default=int(os.environ.get("PO_PORT", "8585")))
    ap.add_argument("--fonts", action="append", default=[], help="font LIBRARY folder (Linux/Termux). Can hold thousands of fonts; only the ones a document needs are loaded. May be repeated.")
    ap.add_argument("--fonts-all", action="store_true", help="load the whole font library for every file (slow; for troubleshooting)")
    ap.add_argument("--soffice", default=None, help="full path of soffice / soffice.com")
    ap.add_argument("--allow-any", action="store_true", help="also serve clients that are not on a private network (unsafe)")
    a = ap.parse_args()
    host = "0.0.0.0" if a.lan else a.host
    CFG["soffice"] = find_soffice(a.soffice)
    CFG["allow_any"] = a.allow_any
    libs = list(a.fonts)
    if os.environ.get("PO_FONTS"):
        libs.append(os.environ["PO_FONTS"])
    if not libs and not IS_WIN and os.path.isdir("/sdcard/PocketOffice/Fonts"):
        libs.append("/sdcard/PocketOffice/Fonts")
    CFG["library"] = [os.path.abspath(os.path.expanduser(p)) for p in libs if os.path.isdir(os.path.expanduser(p))]
    CFG["fonts_all"] = a.fonts_all
    prepare()

    print("Pocket Office converter v%s" % VERSION)
    if POE:
        print("  Editors     : Word + Excel + PowerPoint ready (edit a COPY only, everything is logged to %s)" % EDIT_LOG)
    else:
        print("  Editors     : NOT AVAILABLE - viewing/conversion still works, editing does not")
        for ln in editors_error_text().split("\n"):
            print("                " + ln)
    print("  LibreOffice :", CFG["soffice"] or "NOT FOUND - install LibreOffice (https://www.libreoffice.org) or use --soffice PATH")
    fonts = installed_fonts(force=True)
    if IS_WIN or IS_MAC:
        print("  Fonts       : all fonts installed in %s are used automatically (%s families found)" % (
            "Windows" if IS_WIN else "macOS", len(fonts) if fonts else "?"))
        if CFG["library"]:
            print("                (--fonts is ignored here: install extra fonts into the system instead)")
    else:
        print("  Fonts       : system fonts (%s families found)" % (len(fonts) if fonts else "?"))
        if CFG["library"]:
            print("  Font library: %s" % ", ".join(CFG["library"]))
            t0 = time.time()
            refresh_index(verbose=True)
            print("                %d font files indexed in %.1fs - %s" % (len(_index["files"]), time.time() - t0,
                  "ALL are loaded for every file (--fonts-all)" if CFG["fonts_all"] else "only the fonts a document needs are loaded"))
    print("  Listening   : %s:%d%s" % (host, a.port, "" if host != "127.0.0.1" else "   (this machine only - use --lan for your phone)"))
    print("  Web panel   : http://%s:%d/   (open it in the phone's browser: editor with all tools; every browser tab is its own window)" % ("127.0.0.1" if host in ("127.0.0.1", "0.0.0.0") else host, a.port))
    if host in ("0.0.0.0", "::"):
        for ip in local_ips():
            print("  In the app  : http://%s:%d" % (ip, a.port))
        print("  Only private-network clients are served. Allow Python in the Windows firewall (Private networks) when asked.")
    print("  Stop with Ctrl+C\n")
    try:
        ThreadingHTTPServer((host, a.port), Handler).serve_forever()
    except KeyboardInterrupt:
        print("bye")


if __name__ == "__main__":
    main()
