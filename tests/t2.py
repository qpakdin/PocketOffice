import sys, io, zipfile, subprocess, os, re, hashlib
sys.path.insert(0, '/home/claude/po/PocketOffice/converter')
import po_docx_edit as E
fails = []
def check(name, cond, extra=""):
    print(("PASS " if cond else "FAIL ") + name + ((" -> " + str(extra)) if (extra and not cond) else ""))
    if not cond: fails.append(name)

orig = open('/tmp/ed/complex.docx', 'rb').read()
s = E.DocxSession(orig, 'complex.docx')
# locked paragraphs refuse edits
for i in (2, 3, 4, 5, 8):
    try: s.apply([{"op": "set_text", "p": i, "text": "hack"}]); check("locked p%d refuses text edit" % i, False)
    except E.EditError as e: check("locked p%d refuses text edit" % i, "locked" in str(e))
    try: s.apply([{"op": "format", "p": i, "bold": True}]); check("locked p%d refuses format" % i, False)
    except E.EditError: check("locked p%d refuses format" % i, True)
    try: s.apply([{"op": "delete_paragraph", "p": i}]); check("locked p%d cannot be deleted" % i, False)
    except E.EditError: check("locked p%d cannot be deleted" % i, True)
check("failed attempts changed nothing", s.cur == s.doc0)

# real edits on editable paragraphs
sc0 = E.Scan(s.doc0); a = s.doc0.decode('latin-1')
locked_xml0 = {i: a[r['start']:r['end']] for i, r in enumerate([r for r in sc0.paras if not r['in_fallback']]) if i in (2, 3, 4, 5, 8)}
s.apply([{"op": "set_text", "p": 1, "text": "Plain paragraph one, now EDITED with new words."},
         {"op": "format", "p": 1, "start": 0, "end": 5, "bold": True, "color": "0000FF"},
         {"op": "para", "p": 0, "align": "center"},
         {"op": "set_text", "p": 9, "text": "Final paragraph changed."},
         {"op": "insert_paragraph_after", "p": 7},
         {"op": "page", "scope": "last", "w_cm": 21.59, "h_cm": 27.94, "top": 2.54, "bottom": 2.54, "left": 3, "right": 3}])
check("edited text present", s.blocks()[1]['text'].startswith("Plain paragraph one, now EDITED"))
sc1 = E.Scan(s.cur); b = s.view()
L1 = [r for r in sc1.paras if not r['in_fallback']]
# paragraph indexes shifted by one after index 7 (inserted); locked ones before 7 must be byte-identical, 8 -> 9
mapping = {2: 2, 3: 3, 4: 4, 5: 5, 8: 9}
check("locked paragraphs byte-identical after edits", all(b[L1[j]['start']:L1[j]['end']] == locked_xml0[i] for i, j in mapping.items()))
check("fallback copies byte-identical", all(a[r['start']:r['end']] == b[r2['start']:r2['end']] for r, r2 in zip([r for r in sc0.paras if r['in_fallback']], [r for r in sc1.paras if r['in_fallback']])))
v = s.verify()
for c in v['checks']: print("   verify:", "ok " if c['ok'] else "BAD", c['name'], "|", c['detail'])
check("verify passes", v['ok'])
out = s.export()
open('/tmp/ed/complex_edited.docx', 'wb').write(out)
zo, zi = zipfile.ZipFile(io.BytesIO(out)), zipfile.ZipFile(io.BytesIO(orig))
check("only word/document.xml differs", [n for n in zi.namelist() if zi.read(n) != zo.read(n)] == ['word/document.xml'])
check("zip order identical", zi.namelist() == zo.namelist())
check("zip test() ok", zo.testzip() is None)
check("original bytes untouched in memory", hashlib.sha256(orig).hexdigest() == s.sha)
print("\nFAILED:", fails or "none")
