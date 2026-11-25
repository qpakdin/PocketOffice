#!/usr/bin/env python3
"""
Pocket Office doctor: checks that this machine can run the converter + editors, and tells you the exact fix for anything missing.
Run it with the SAME python that starts the server:   python3 po_doctor.py
"""
import importlib, os, shutil, socket, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
IS_WIN = os.name == "nt"
TERMUX = os.environ.get("PREFIX", "").startswith("/data/data/com.termux")
DEBIAN = os.path.exists("/etc/debian_version")
bad = 0


def line(ok, what, fix=""):
    global bad
    print(("  OK    " if ok else "  FIX   ") + what)
    if not ok:
        bad += 1
        if fix:
            for l in fix.split("\n"):
                print("          -> " + l)


def apt(pkgs):
    if IS_WIN:
        return ""
    if TERMUX:
        return "(inside Debian proot)  apt install -y " + pkgs
    return "apt install -y " + pkgs


print("Pocket Office doctor  |  python %s  |  %s" % (sys.version.split()[0], sys.executable))
print("Environment: %s\n" % ("Windows" if IS_WIN else "Termux (outer shell - the server normally runs INSIDE Debian proot)" if TERMUX else "Debian/Ubuntu Linux" if DEBIAN else "Linux/macOS"))

print("Server files (the whole 'converter' folder must stay together):")
for f in ("pocket_office_server.py", "po_docx_edit.py", "po_office_edit.py", "po_new_file.py", "wordtk/__init__.py", "panel/index.html", "panel/panel.js", "panel/panel.css"):
    line(os.path.exists(os.path.join(HERE, f)), f, "copy the complete converter folder (not single files) to the machine that runs the server")

print("\nPython modules:")
try:
    importlib.import_module("lxml.etree")
    line(True, "lxml (needed by ALL editors)")
except Exception as e:
    line(False, "lxml (needed by ALL editors): %s" % e, ("pip install lxml" if IS_WIN else apt("python3-lxml") + "\n   (Termux python instead of Debian: pkg install python-lxml)\n   If you run the server with a venv/pip python: pip install lxml"))

sys.path.insert(0, HERE)
try:
    import po_office_edit as POE
    line(not POE.OFFICE_EDIT_ERROR, "editing engine loads" + ("" if not POE.OFFICE_EDIT_ERROR else ": " + str(POE.OFFICE_EDIT_ERROR)), "fix the item above, then run this doctor again")
except Exception as e:
    line(False, "editing engine loads: %s: %s" % (type(e).__name__, e), "fix the items above, then run this doctor again")

print("\nPrograms:")
so = shutil.which("soffice") or shutil.which("libreoffice")
if not so and IS_WIN:
    for p in (r"C:\Program Files\LibreOffice\program\soffice.com", r"C:\Program Files (x86)\LibreOffice\program\soffice.com"):
        if os.path.exists(p):
            so = p
line(bool(so), "LibreOffice (soffice)" + (": " + so if so else ""), "install LibreOffice (Windows: libreoffice.org, then add ...\\LibreOffice\\program to PATH)\n" + apt("libreoffice-writer libreoffice-calc libreoffice-impress"))
for prog in ("pdftotext", "pdfinfo", "pdftoppm"):
    line(bool(shutil.which(prog)), prog + " (poppler-utils)" + ("" if shutil.which(prog) else " - needed for text export / page previews"),
         "Windows: install poppler and add its bin folder to PATH\n" + apt("poppler-utils"))

print("\nNetwork (for using the phone app with a PC):")
ips = []
try:
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM); s.connect(("10.255.255.255", 1)); ips.append(s.getsockname()[0]); s.close()
except Exception:
    pass
if ips:
    print("  This machine's address: %s   ->  in the phone app use  http://%s:8585  and start the server with  --lan" % (ips[0], ips[0]))
else:
    print("  could not detect a LAN address (not on Wi-Fi?)")
port_free = True
try:
    t = socket.socket(); t.settimeout(0.5)
    port_free = t.connect_ex(("127.0.0.1", 8585)) != 0; t.close()
except Exception:
    pass
print("  Port 8585: %s" % ("free" if port_free else "already in use (a server is probably running - that is fine, or use --port)"))
if IS_WIN:
    print("  Windows firewall: when Python asks, allow it on PRIVATE networks; the phone and the PC must be on the same Wi-Fi.")

print("\n%s" % ("Everything needed is present. Start:  python3 pocket_office_server.py%s" % (" --lan" if ips else "") if not bad else "%d thing(s) to fix (see 'FIX' above). Then run the doctor again." % bad))
sys.exit(1 if bad else 0)
