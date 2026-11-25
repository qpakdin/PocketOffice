import subprocess, time, urllib.request, urllib.error, urllib.parse, os, sys, signal, json, hashlib, io, zipfile
S = "/home/claude/po/PocketOffice/converter/pocket_office_server.py"
log = open("srv_edit.log", "w")
p = subprocess.Popen([sys.executable, S, "--port", "8821"], stdout=log, stderr=subprocess.STDOUT, preexec_fn=os.setsid)
for _ in range(100):
    if "Stop with Ctrl+C" in open("srv_edit.log").read(): break
    time.sleep(0.3)
B = "http://127.0.0.1:8821"
def call(path, data=None):
    try:
        r = urllib.request.urlopen(urllib.request.Request(B + path, data=data, method="POST" if data is not None else "GET"), timeout=180)
        return r.status, r.read(), dict(r.headers)
    except urllib.error.HTTPError as e:
        return e.code, e.read(), dict(e.headers)
def J(path, obj=None):
    raw = obj if isinstance(obj, (bytes, bytearray)) else (None if obj is None else json.dumps(obj).encode())
    st, body, h = call(path, raw)
    return st, json.loads(body.decode())
fails = []
def check(n, c, x=""):
    print(("PASS " if c else "FAIL ") + n + ((" -> " + str(x)) if (x and not c) else "")); 
    if not c: fails.append(n)

orig = open("/tmp/ed/complex.docx", "rb").read()
print(call("/ping")[1].decode()[:170])
st, r = J("/edit/open?ext=docx&name=" + urllib.parse.quote("گزارش complex.docx"), orig)
check("open returns session, blocks and a preflight report", st == 200 and r["ok"] and len(r["blocks"]) == 10 and r["report"]["stats"]["locked"] == 5, r.get("error"))
sid = r["sid"]
st, r = J("/edit/apply?sid=" + sid, {"ops": [{"op": "set_text", "p": 1, "text": "Edited through the server"}, {"op": "format", "p": 1, "start": 0, "end": 6, "bold": True}]})
check("apply ok, returns only changed blocks", r["ok"] and not r["structure_changed"] and [b["i"] for b in r["blocks"]] == [1] and r["blocks"][0]["text"] == "Edited through the server", r)
st, r = J("/edit/apply?sid=" + sid, {"ops": [{"op": "set_text", "p": 2, "text": "hack a locked paragraph"}]})
check("locked paragraph refused with a clear message", st == 200 and not r["ok"] and "locked" in r["error"], r)
st, r = J("/edit/apply?sid=" + sid, {"ops": [{"op": "insert_paragraph_after", "p": 1}]})
check("structural op returns all blocks", r["ok"] and r["structure_changed"] and len(r["blocks"]) == 11)
st, r = J("/edit/undo?sid=" + sid, {})
check("undo removes the last op", r["ok"] and len(r["blocks"]) == 10 and r["ops"] == 2)
st, pdf, h = call("/edit/preview?sid=" + sid)
check("preview is a real PDF with a page count header", st == 200 and pdf[:5] == b"%PDF-" and int(h.get("X-PO-Pages", 0)) >= 1, h)
st, r = J("/edit/verify?sid=" + sid + "&orig_pages=1")
v = r["verify"]
check("verify passes incl. LibreOffice opening the file", st == 200 and v["ok"] and any("LibreOffice can open" in c["name"] and c["ok"] for c in v["checks"]), v)
st, data, h = call("/edit/export?sid=" + sid)
z = zipfile.ZipFile(io.BytesIO(data)); zo = zipfile.ZipFile(io.BytesIO(orig))
check("export is a valid docx, sha header matches", st == 200 and h.get("X-PO-SHA256") == hashlib.sha256(data).hexdigest() and z.testzip() is None)
check("export: only document.xml differs from the original", [n for n in zo.namelist() if zo.read(n) != z.read(n)] == ["word/document.xml"])
check("original upload bytes were never modified", hashlib.sha256(orig).hexdigest() == r_sha if (r_sha := hashlib.sha256(open('/tmp/ed/complex.docx','rb').read()).hexdigest()) else False)
st, r = J("/edit/blocks?sid=nope"); check("unknown session gives a clear 404", st == 404 and not r["ok"], r)
st, r = J("/edit/open?ext=doc&name=x.doc", b"x"); check(".doc editing refused with explanation", st == 400 and "re-written" in r["error"], r)
prot = io.BytesIO(); zi = zipfile.ZipFile("/tmp/ed/src.docx"); zp = zipfile.ZipFile(prot, "w")
for n in zi.namelist():
    d = zi.read(n)
    if n == "word/settings.xml": d = d.replace(b"</w:settings>", b'<w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>')
    zp.writestr(n, d)
zp.close()
st, r = J("/edit/open?ext=docx&name=p.docx", prot.getvalue()); sid2 = r["sid"]
st, r2 = J("/edit/apply?sid=" + sid2, {"ops": [{"op": "set_text", "p": 1, "text": "x"}]})
check("protected document refuses every edit", not r2["ok"] and "protect" in r2["error"].lower(), r2)
st, r = J("/edit/close?sid=" + sid, {}); check("close", r["ok"])
st, _, _ = call("/ping"); 
os.killpg(os.getpgid(p.pid), signal.SIGTERM)
print("\nFAILED:", fails or "none")
print("\n===== what the server printed (this is what appears in Termux) =====")
for line in open("srv_edit.log").read().splitlines():
    if "[EDIT" in line: print(line[:190])
