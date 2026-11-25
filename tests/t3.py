import sys, io, zipfile, random, re, subprocess, os, shutil
sys.path.insert(0, '/home/claude/po/PocketOffice/converter')
import po_docx_edit as E
fails = []
def check(name, cond, extra=""):
    print(("PASS " if cond else "FAIL ") + name + ((" -> " + str(extra)) if (extra and not cond) else ""))
    if not cond: fails.append(name)

# ---- negative tests: does verify() really catch damage?
s = E.DocxSession(open('/tmp/ed/complex.docx', 'rb').read(), 'c')
s.apply([{"op": "set_text", "p": 1, "text": "changed"}])
good = s.cur
x = good.decode('latin-1').replace("Inline picture:", "Inline PICTURE:")          # silently damage a locked paragraph
s.cur = x.encode('latin-1'); s._scan = None
v = s.verify(); check("verify detects damage to a locked paragraph", not v['ok'] and any(("Locked" in c['name'] and not c['ok']) for c in v['checks']), [c for c in v['checks'] if not c['ok']])
s.cur = good.decode('latin-1').replace("</w:body>", "<w:p><w:r><w:t>sneaky</w:t></w:r></w:p></w:body>").encode('latin-1'); s._scan = None
v = s.verify(); check("verify detects an unexpected extra paragraph", not v['ok'], [c['name'] for c in v['checks'] if not c['ok']])
s.cur = good.decode('latin-1').replace("<w:body>", "<w:body><w:p><w:r><w:t>x</w:r></w:p>").encode('latin-1'); s._scan = None
v = s.verify(); check("verify detects broken XML", not v['ok'] and any(("well-formed" in c['name'] and not c['ok']) for c in v['checks']))

# ---- crafted risky documents
def make(mod_doc=None, mod_settings=None, extra_files=None, base='/tmp/ed/src.docx'):
    zi = zipfile.ZipFile(base); buf = io.BytesIO(); zo = zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED)
    for n in zi.namelist():
        d = zi.read(n)
        if n == 'word/document.xml' and mod_doc: d = mod_doc(d.decode('utf-8')).encode('utf-8')
        if n == 'word/settings.xml' and mod_settings: d = mod_settings(d.decode('utf-8')).encode('utf-8')
        zo.writestr(n, d)
    for n, d in (extra_files or {}).items(): zo.writestr(n, d)
    zo.close(); return buf.getvalue()
def addp(x, para): return x.replace("</w:body>", para + "</w:body>") if "<w:sectPr" not in x else re.sub(r"(<w:sectPr)", para + r"\1", x, count=1)
protect = make(mod_settings=lambda x: x.replace("</w:settings>", '<w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>'))
ss = E.DocxSession(protect, 'p'); check("document protection blocks editing", bool(ss.report['blocked']))
try: ss.apply([{"op": "set_text", "p": 1, "text": "x"}]); check("protected doc refuses apply", False)
except E.EditError: check("protected doc refuses apply", True)
ss = E.DocxSession(b"not a zip at all", 'x'); check("garbage file is blocked, no crash", bool(ss.report['blocked']) and ss.blocks() == [])
ss = E.DocxSession(make(extra_files={'_xmlsignatures/sig1.xml': b'<x/>', 'word/vbaProject.bin': b'\x00'}), 's')
txt = " ".join(i['text'] for i in ss.report['items']); check("signature and macro warnings", "DIGITAL SIGNATURE" in txt and "MACROS" in txt)
tc = make(mod_doc=lambda x: addp(x, '<w:p><w:ins w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>tracked new</w:t></w:r></w:ins></w:p><w:p><w:r><w:t>normal</w:t></w:r></w:p>'), mod_settings=lambda x: x.replace("</w:settings>", "<w:trackRevisions/></w:settings>"))
ss = E.DocxSession(tc, 't'); bl = ss.blocks(); txt = " ".join(i['text'] for i in ss.report['items'])
check("tracked insertion paragraph locked", [b for b in bl if 'tracked new' in b['text']][0]['locked'] and not [b for b in bl if b['text'] == 'normal'][0]['locked'])
check("trackRevisions-on warning", "NOT recorded as tracked changes" in txt)
sdt = make(mod_doc=lambda x: addp(x, '<w:sdt><w:sdtPr><w:dataBinding w:xpath="/a" w:storeItemID="{1}"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>bound</w:t></w:r></w:p></w:sdtContent></w:sdt>'))
ss = E.DocxSession(sdt, 'b'); txt = " ".join(i['text'] for i in ss.report['items']); check("data-bound content control warning", "bound to custom XML" in txt, txt)
link = make(mod_doc=lambda x: addp(x, '<w:p><w:hyperlink r:id="rId99" w:history="1"><w:r><w:t>click here</w:t></w:r></w:hyperlink><w:r><w:t> tail</w:t></w:r></w:p>'))
ss = E.DocxSession(link, 'l'); i = [b['i'] for b in ss.blocks() if 'click here' in b['text']][0]
check("hyperlink paragraph is editable", not ss.blocks()[i]['locked'])
ss.apply([{"op": "set_text", "p": i, "text": "click HERE now tail!"}]); check("edit inside hyperlink text", ss.blocks()[i]['text'] == "click HERE now tail!" and 'rId99' in ss.view())

# ---- fuzz
def fuzz(path, seed, n):
    random.seed(seed)
    sess = E.DocxSession(open(path, 'rb').read(), os.path.basename(path)); done = 0
    words = ["alpha", "بتا", "gamma ", " δ", "x&y", "<tag>", "a\tb", "l1\nl2", "😀", "  spaced  ", ""]
    for _ in range(n):
        bl = sess.blocks(); ed = [b for b in bl if not b['locked']]
        if not ed: break
        b = random.choice(ed); i = b['i']; t = b['text']; r = random.random()
        try:
            if r < 0.40:
                pos = random.randint(0, len(t)); end = random.randint(pos, len(t)); new = t[:pos] + random.choice(words) + t[end:]
                sess.apply([{"op": "set_text", "p": i, "text": new}])
                assert sess.blocks()[i]['text'] == new, "text mismatch"
            elif r < 0.65 and t:
                a = random.randint(0, len(t) - 1); c = random.randint(a + 1, len(t))
                sess.apply([{"op": "format", "p": i, "start": a, "end": c, "bold": random.choice([True, False, None]), "size": random.choice([None, 9, 14.5, 20]),
                             "color": random.choice([None, "FF0000", "00FF00"]), "font": random.choice([None, "Arial", "B Nazanin"]), "italic": random.choice([None, True]), "underline": random.choice([None, True])}])
                assert sess.blocks()[i]['text'] == t, "format changed text"
            elif r < 0.78:
                sess.apply([{"op": "para", "p": i, "align": random.choice(["left", "center", "right", "justify"]), "rtl": random.choice([None, True, False]),
                             "before": random.choice([None, 0, 6]), "after": random.choice([None, 12]), "line": random.choice([None, 1.0, 1.5])}])
            elif r < 0.86:
                sess.apply([{"op": "insert_paragraph_after", "p": i}])
            elif r < 0.91:
                sess.apply([{"op": "delete_paragraph", "p": i}])
            elif r < 0.95 and sess.table_rows():
                sess.apply([{"op": "insert_row", "t": 0, "r": random.randrange(sess.table_rows()[0]), "where": random.choice(["above", "below"])}])
            elif r < 0.97 and sess.table_rows() and sess.table_rows()[0] > 1:
                sess.apply([{"op": "delete_row", "t": 0, "r": random.randrange(sess.table_rows()[0])}])
            elif r < 0.99:
                sess.apply([{"op": "page", "scope": random.choice(["last", "all"]), "w_cm": random.choice([21, 29.7]), "h_cm": random.choice([29.7, 21]), "top": 2, "bottom": 2, "left": 2, "right": 2}])
            else:
                sess.undo()
            done += 1
        except E.EditError:
            pass
        assert E._wellformed(sess.cur) is None, "ill-formed XML"
    return sess, done
for path, seed in (('/tmp/ed/src.docx', 1), ('/tmp/ed/src.docx', 2), ('/tmp/ed/complex.docx', 3), ('/tmp/ed/complex.docx', 4)):
    try:
        sess, done = fuzz(path, seed, 250)
        v = sess.verify()
        bad = [c['name'] for c in v['checks'] if not c['ok']]
        check("fuzz %s seed %d: %d ops, verify ok" % (os.path.basename(path), seed, done), v['ok'], bad)
        out = '/tmp/ed/fuzz_%d.docx' % seed; open(out, 'wb').write(sess.export())
    except AssertionError as e:
        check("fuzz %s seed %d" % (path, seed), False, e)
print("\nFAILED:", fails or "none")
