import subprocess, time, urllib.request, urllib.error, urllib.parse, os, sys, signal, json, hashlib, io, zipfile, http.client
S = "/home/claude/po/PocketOffice/converter/pocket_office_server.py"
log = open("srv8.log", "w")
p = subprocess.Popen([sys.executable, S, "--port", "8831"], stdout=log, stderr=subprocess.STDOUT, preexec_fn=os.setsid)
for _ in range(120):
    if "Stop with Ctrl+C" in open("srv8.log").read(): break
    time.sleep(0.3)
B = "127.0.0.1:8831"
def call(method, path, data=None, headers=None):
    c = http.client.HTTPConnection("127.0.0.1", 8831, timeout=240)
    c.request(method, path, body=data, headers=headers or {})
    r = c.getresponse(); body = r.read(); h = dict((k.lower(), v) for k, v in r.getheaders()); c.close(); return r.status, body, h
def J(method, path, obj=None):
    raw = obj if isinstance(obj, (bytes, bytearray)) else (None if obj is None else json.dumps(obj).encode())
    st, body, h = call(method, path, raw); return st, json.loads(body.decode())
fails = []
def check(n, c, x=""):
    print(("PASS " if c else "FAIL ") + n + ((" -> " + str(x)[:300]) if (x and not c) else ""))
    if not c: fails.append(n)
print(call("GET", "/ping")[1].decode()[:230])
st, b, h = call("GET", "/"); check("/ redirects to the panel", st == 302 and h.get("location") == "/panel/")
st, b, h = call("GET", "/panel/"); check("panel html served", st == 200 and b"panel.js" in b)
st, b, h = call("GET", "/panel/panel.js"); check("panel js served", st == 200 and b"Pocket Office web panel" in b)
check("unknown panel file is 404", call("GET", "/panel/../pocket_office_server.py")[0] == 404)
# host / origin guard
c = http.client.HTTPConnection("127.0.0.1", 8831); c.request("GET", "/ping", headers={"Host": "evil.example.com"}); check("DNS-rebinding host name refused", c.getresponse().status == 403)
st, _, _ = call("POST", "/edit/new?ext=docx", b"{}", {"Origin": "http://evil.example.com"}); check("cross-site POST refused", st == 403)
st, _, _ = call("POST", "/edit/new?ext=docx&name=a.docx", b"{}", {"Origin": "http://127.0.0.1:8831"}); check("same-origin POST accepted", st == 200)
st, cat = J("GET", "/edit/catalog"); check("catalog lists formats", cat["groups"]["Word"][0] == "docx" and "xlsb" in cat["not_creatable"])
# new files of every group
sids = {}
for ext, nm in (("docx", "Untitled.docx"), ("xlsx", "Book.xlsx"), ("pptx", "Deck.pptx"), ("odt", "Note.odt"), ("xls", "Old.xls"), ("odp", "Old.odp"), ("dotx", "T.dotx")):
    st, r = J("POST", "/edit/new?" + urllib.parse.urlencode({"ext": ext, "name": nm}), {}); check("new .%s" % ext, st == 200 and r["ok"], r)
    sids[ext] = r["sid"]
st, r = J("POST", "/edit/new?ext=xlsb&name=x.xlsb", {}); check("xlsb refused with a clear reason", st == 400 and "xlsb" in r["error"], r)
st, r = J("POST", "/edit/new?ext=sxw&name=x.sxw", {}); check("sxw refused with a clear reason", st == 400 and "OpenOffice 1.x" in r["error"], r)
# Word
sid = sids["docx"]
st, r = J("POST", "/edit/apply?sid=" + sid, {"ops": [{"op": "set_text", "p": 0, "text": "Typed into a new document"}, {"op": "insert_paragraph_after", "p": 0}, {"op": "set_text", "p": 1, "text": "سلام"}]}); check("word apply", r["ok"] and len(r["blocks"]) == 2, r)
st, r = J("GET", "/edit/schema?sid=" + sid); check("word schema 13 groups", len(r["groups"]) == 13)
st, r = J("POST", "/edit/settings?sid=" + sid, {"group": "text", "values": {"replacements": json.dumps([{"find": "new", "replace": "NEW"}])}, "dry": True}); check("settings preview", r["ok"] and r["dry"] and r["result"]["paragraphs_text_changed"] == 1, r)
st, r = J("POST", "/edit/settings?sid=" + sid, {"group": "text", "values": {"replacements": json.dumps([{"find": "new", "replace": "NEW"}])}}); check("settings apply returns blocks", r["ok"] and "NEW" in r["blocks"][0]["text"], r)
st, r = J("POST", "/edit/settings?sid=" + sid, {"group": "paragraph", "values": {"alignment": "center"}, "scope": {"parts": ["body"], "styles": None, "include_cover": False}}); check("settings with scope", r["ok"], r)
st, r = J("POST", "/edit/settings?sid=" + sid, {"group": "paragraph", "values": {"alignment": "diagonal"}}); check("bad value refused clearly", not r["ok"], r)
st, r = J("GET", "/edit/find?sid=" + sid + "&q=NEW"); check("find", r["matches"] == 1 and r["results"][0]["i"] == 0, r)
st, pdf, h = call("GET", "/edit/render?sid=" + sid); rr = json.loads(pdf); check("render pages", rr["ok"] and rr["pages"] >= 1, rr)
st, png, h = call("GET", "/edit/page?sid=%s&n=1&dpi=70&k=%s" % (sid, rr["key"])); check("page image is a PNG", st == 200 and png[:4] == b"\x89PNG" and len(png) > 1000, (st, png[:50]))
st, r = J("GET", "/edit/verify?sid=" + sid); check("verify passes incl. LibreOffice", r["verify"]["ok"], r)
for fmt in ("docx", "pdf", "odt", "rtf", "doc"):
    st, data, h = call("GET", "/edit/export?sid=%s&as=%s" % (sid, fmt)); ok = st == 200 and len(data) > 200 and h.get("content-disposition", "").startswith("attachment") and ("filename*=UTF-8''Untitled.%s" % fmt) in h.get("content-disposition", "")
    check("export as .%s (%d bytes)" % (fmt, len(data)), ok, (st, h.get("content-disposition")))
# Excel
sid = sids["xlsx"]
ops = [{"op": "set_cell", "sheet": 0, "ref": "A1", "value": "Item"}, {"op": "set_cell", "sheet": 0, "ref": "B1", "value": "Qty"}, {"op": "set_cell", "sheet": 0, "ref": "A2", "value": "Pen"}, {"op": "set_cell", "sheet": 0, "ref": "B2", "value": "12"}, {"op": "set_cell", "sheet": 0, "ref": "B3", "value": "=B2*3"}]
st, r = J("POST", "/edit/apply?sid=" + sid, {"ops": ops}); check("excel set cells", r["ok"], r)
st, g = J("GET", "/edit/grid?sid=%s&sheet=0&r0=1&c0=1&rows=5&cols=4" % sid); check("grid window", g["rows"][1][1][0] == "12" and g["rows"][2][1][1] == "=B2*3" and g["sheets"] == ["Sheet1"], g)
st, r = J("POST", "/edit/settings?sid=" + sid, {"group": "format", "values": {"sheets": "Sheet1", "range": "A1:B1", "bold": True, "fill": "FFFF00"}}); check("excel format", r["ok"], r)
st, r = J("POST", "/edit/settings?sid=" + sid, {"group": "page", "values": {"orientation": "landscape", "fit_width": "1", "footer_text": "Page &P"}}); check("excel page setup", r["ok"], r)
st, r = J("GET", "/edit/find?sid=" + sid + "&q=pen"); check("excel find", r["matches"] == 1 and r["results"][0]["ref"] == "A2", r)
st, data, h = call("GET", "/edit/export?sid=%s&as=xls" % sid); check("excel export as .xls", st == 200 and len(data) > 1000, st)
st, r = J("GET", "/edit/verify?sid=" + sid); check("excel verify", r["verify"]["ok"], r)
# PowerPoint
sid = sids["pptx"]
st, r = J("POST", "/edit/apply?sid=" + sid, {"ops": [{"op": "set_text", "p": 0, "text": "Deck title"}]}); check("pptx set_text", r["ok"] and r["blocks"][0]["text"] == "Deck title", r)
st, r = J("POST", "/edit/settings?sid=" + sid, {"group": "font", "values": {"size": "40", "bold": True}}); check("pptx font settings", r["ok"], r)
st, data, h = call("GET", "/edit/export?sid=%s&as=pdf" % sid); check("pptx export as pdf", st == 200 and data[:4] == b"%PDF", st)
# sessions + close
st, r = J("GET", "/edit/sessions"); check("sessions list", len(r["sessions"]) >= 7 and all("sid" in s for s in r["sessions"]))
st, r = J("POST", "/edit/close?sid=" + sids["docx"], {}); st2, r2 = J("GET", "/edit/blocks?sid=" + sids["docx"]); check("close removes the session", r["ok"] and st2 == 404)
os.killpg(os.getpgid(p.pid), signal.SIGTERM)
print("\nFAILED:", fails or "none")
