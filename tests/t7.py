import sys, io, zipfile, subprocess, os, shutil, glob
sys.path.insert(0, '/home/claude/po/PocketOffice/converter')
import po_new_file as N, po_office_edit as O
from lxml import etree
fails = []
def check(n, c, x=""):
    print(("PASS " if c else "FAIL ") + n + ((" -> " + str(x)[:250]) if (x and not c) else ""))
    if not c: fails.append(n)
shutil.rmtree('/tmp/ed/new', ignore_errors=True); os.makedirs('/tmp/ed/new/pdf')
# 1. every OOXML variant: valid zip, all XML well-formed, LibreOffice opens it, our session opens it
for ext in sorted(N.MAIN_CT):
    data, base, save = N.create(ext)
    z = zipfile.ZipFile(io.BytesIO(data)); bad = []
    for n in z.namelist():
        if n.endswith(('.xml', '.rels')):
            try: etree.fromstring(z.read(n))
            except Exception as e: bad.append(n)
    ct = z.read('[Content_Types].xml').decode()
    path = '/tmp/ed/new/new.%s' % ext; open(path, 'wb').write(data)
    r = subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', '/tmp/ed/new/pdf', path], capture_output=True, text=True, timeout=120)
    ok_lo = os.path.exists('/tmp/ed/new/pdf/new.pdf'); 
    if ok_lo: os.rename('/tmp/ed/new/pdf/new.pdf', '/tmp/ed/new/pdf/%s.pdf' % ext)
    s = O.open_session(ext, data, 'Untitled.' + ext)
    check("new .%s: valid zip, XML ok, content type %s, LibreOffice opens, session opens" % (ext, "ok" if N.MAIN_CT[ext] in ct else "WRONG"),
          not bad and z.testzip() is None and N.MAIN_CT[ext] in ct and ok_lo and not s.blocked, (bad, ok_lo, s.blocked))
# 2. blank docs are editable by our editors
d, _, _ = N.create('docx'); s = O.WordSession(d, 'n.docx')
check("blank docx has one empty editable paragraph", len(s.blocks()) == 1 and s.blocks()[0]['text'] == '' and not s.blocks()[0]['locked'], s.blocks())
s.apply([{"op": "set_text", "p": 0, "text": "Hello from a brand new document"}]); s.apply([{"op": "insert_paragraph_after", "p": 0}]); s.apply([{"op": "set_text", "p": 1, "text": "سلام دنیا"}])
s.apply([{"op": "format", "p": 0, "bold": True, "size": 18, "font": "Arial"}]); s.apply([{"op": "para", "p": 1, "rtl": True, "align": "right"}])
s.settings("page", {"paper": "A4", "margin_top": "2cm"}); print("   texts:", s._dx().full_texts())
check("typing into self-closed <w:p/> works", s._dx().full_texts() == ["Hello from a brand new document", "سلام دنیا"])
v = s.verify(); check("verify ok", v['ok'], v['checks']); open('/tmp/ed/new/edited.docx', 'wb').write(s.export())
d, _, _ = N.create('xlsx'); x = O.XlsxSession(d, 'n.xlsx')
x.apply([{"op": "set_cell", "sheet": 0, "ref": "A1", "value": "Name"}, {"op": "set_cell", "sheet": 0, "ref": "B1", "value": "Qty"}, {"op": "set_cell", "sheet": 0, "ref": "A2", "value": "Pen"}, {"op": "set_cell", "sheet": 0, "ref": "B2", "value": "12"}, {"op": "set_cell", "sheet": 0, "ref": "B3", "value": "=B2*3"}])
x.settings("format", {"range": "A1:B1", "bold": True, "fill": "FFFF00", "font": "Arial"})
c = {k['ref']: k['text'] for k in x.cells(0)['cells']}; print("   cells:", c)
check("blank xlsx: cells entered and formatted", c.get('A1') == 'Name' and c.get('B2') == '12'); check("verify ok (xlsx)", x.verify()['ok']); open('/tmp/ed/new/edited.xlsx', 'wb').write(x.export())
d, _, _ = N.create('pptx'); p = O.PptxSession(d, 'n.pptx'); b = p.blocks()
check("blank pptx has title + subtitle placeholders (empty, editable)", len(b) == 2 and all(not q['locked'] for q in b), b)
p.apply([{"op": "set_text", "p": 0, "text": "My Presentation"}, {"op": "set_text", "p": 1, "text": "subtitle here"}]); check("verify ok (pptx)", p.verify()['ok']); open('/tmp/ed/new/edited.pptx', 'wb').write(p.export())
# 3. save-as conversion of every other format
for base, ed in (('docx', 'edited.docx'), ('xlsx', 'edited.xlsx'), ('pptx', 'edited.pptx')):
    for ext in [e for e, b in N.EDIT_AS.items() if b == base]:
        out = '/tmp/ed/new/conv_%s' % ext; shutil.rmtree(out, ignore_errors=True)
        subprocess.run(['soffice', '--headless', '--convert-to', ext, '--outdir', out, '/tmp/ed/new/' + ed], capture_output=True, timeout=120)
        f = glob.glob(out + '/*.' + ext); ok = bool(f) and os.path.getsize(f[0]) > 100
        re_ok = False
        if ok:
            subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', out + '/pdf', f[0]], capture_output=True, timeout=120); re_ok = bool(glob.glob(out + '/pdf/*.pdf'))
        check("convert %s -> .%s and LibreOffice re-opens it" % (base, ext), ok and re_ok, (f, re_ok))
try: N.create('xlsb'); check("xlsb creation refused with explanation", False)
except ValueError as e: check("xlsb creation refused with explanation", "xlsb" in str(e))
print("\nFAILED:", fails or "none")
