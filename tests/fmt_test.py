import subprocess, time, urllib.request, urllib.error, os, sys, signal, glob
S = "/home/claude/po/PocketOffice/converter/pocket_office_server.py"
p = subprocess.Popen([sys.executable, S, "--port", "8811"], stdout=open("srv_fm.log","w"), stderr=subprocess.STDOUT, preexec_fn=os.setsid)
for _ in range(100):
    if "Stop with Ctrl+C" in open("srv_fm.log").read(): break
    time.sleep(0.3)
def req(path, data=None):
    try:
        r = urllib.request.urlopen(urllib.request.Request("http://127.0.0.1:8811"+path, data=data, method="POST" if data is not None else "GET"), timeout=120)
        return r.status, r.read()
    except urllib.error.HTTPError as e:
        return e.code, e.read()
print(req("/ping")[1].decode()[:140])
ok, bad = [], []
for f in sorted(glob.glob("*.*")):
    ext = f.rsplit(".",1)[1]
    if ext in ("fodp",): continue
    code, body = req("/convert?ext=%s" % ext, open(f,"rb").read())
    (ok if code == 200 and body[:4] == b"%PDF" else bad).append((f, code))
print("CONVERTED OK  :", " ".join(f for f,_ in ok))
print("FAILED        :", bad)
print("ext=exe       :", req("/convert?ext=exe", b"x")[0], "(expected 400)")
print("ext=../etc    :", req("/convert?ext=../etc", b"x")[0], "(expected 400)")
print("text docm/doc :", req("/text?ext=docx", open("test.docx","rb").read())[0], req("/text?ext=doc", open("test.doc","rb").read())[0], "(expected 200 400)")
os.killpg(os.getpgid(p.pid), signal.SIGTERM)
