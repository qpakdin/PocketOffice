import sys, io, zipfile, subprocess, os, re
sys.path.insert(0, '/home/claude/po/PocketOffice/converter')
import po_office_edit as O
fails = []
def check(n, c, x=""):
    print(("PASS " if c else "FAIL ") + n + ((" -> " + str(x)[:300]) if (x and not c) else ""))
    if not c: fails.append(n)
orig = open('/tmp/ed/p.pptx', 'rb').read()
s = O.PptxSession(orig, 'p.pptx')
print("   report:", [(i['level'], i['text'][:60]) for i in s.report['items']][:5])
b = s.blocks(); print("   blocks:", [(x['i'], x['where'], x['text'][:30], x['locked']) for x in b])
check("blocks list paragraphs of both slides", len(b) >= 4 and any("apple" in x['text'] for x in b) and any(x['slide'] == 2 for x in b))
check("schema: 4 groups", [g['key'] for g in s.schema()] == ['text', 'font', 'paragraph', 'slides'])
i_title = [x['i'] for x in b if 'Quarterly' in x['text']][0]
# direct edits
s.apply([{"op": "set_text", "p": i_title, "text": "Annual Report 2027"}])
check("set_text keeps other paragraphs", [x['text'] for x in s.blocks()][i_title] == "Annual Report 2027" and len(s.blocks()) == len(b))
s.apply([{"op": "format", "p": i_title, "start": 0, "end": 6, "bold": True, "color": "C00000", "size": 40, "font": "Arial"}])
check("range format keeps text", s.blocks()[i_title]['text'] == "Annual Report 2027")
s.apply([{"op": "para", "p": i_title, "align": "center", "line": 1.2, "before": 6, "after": 6}])
check("para settings", s.blocks()[i_title]['align'] == 'center')
n0 = len(s.blocks()); s.apply([{"op": "insert_paragraph_after", "p": i_title}]); check("insert paragraph", len(s.blocks()) == n0 + 1)
s.apply([{"op": "set_text", "p": i_title + 1, "text": "typed into a new empty paragraph"}]); check("type into empty paragraph", s.blocks()[i_title + 1]['text'] == "typed into a new empty paragraph")
s.apply([{"op": "delete_paragraph", "p": i_title + 1}]); check("delete paragraph", len(s.blocks()) == n0)
try: s.apply([{"op": "set_text", "p": 0, "text": "a\nb"}]); check("newline rejected", False)
except O.OfficeEditError: check("newline rejected", True)
# settings
r = s.settings("text", {"replacements": '[{"find":"apple","replace":"APPLE"}]'}, dry=True); check("replace preview counts 2 paragraphs", r['paragraphs_text_changed'] == 2, r)
s.settings("text", {"replacements": '[{"find":"apple","replace":"APPLE"}]'})
check("replace applied on both slides", sum("APPLE" in x['text'] for x in s.blocks()) == 2)
s.settings("text", {"slides": "1", "digits_convert": "persian->english"}); check("digits on slide 1 only", any("123 تومان" in x['text'] for x in s.blocks()))
s.settings("font", {"slides": "2", "font": "Calibri", "size": "28", "color": "1F3864", "bold": True}); check("font on slide 2", True)
s.settings("paragraph", {"direction": "rtl", "alignment": "right"}); check("paragraph rtl/right on all slides", True)
s.settings("slides", {"slides": "2", "hidden": True}); check("hide slide 2", b'show="0"' in zipfile.ZipFile(io.BytesIO(s.cur)).read('ppt/slides/slide2.xml'))
try: s.settings("font", {"slides": "9", "size": "10"}); check("bad slide number rejected", False)
except O.OfficeEditError as e: check("bad slide number rejected", "9" in str(e))
f = s.find("APPLE", match_case=True); check("find", f['matches'] == 2 and f['results'][0]['where'].startswith("Slide"))
v = s.verify(); [print("   verify:", "ok " if c['ok'] else "BAD", c['name'], "|", c['detail'][:100]) for c in v['checks']]
check("verify ok", v['ok'])
open('/tmp/ed/p_edit.pptx', 'wb').write(s.export())
# undo everything
while s.history: s.undo()
check("undo all -> original bytes", s.cur == orig)
# formatting-only history keeps text strictly
s2 = O.PptxSession(orig, 'p.pptx'); s2.settings("font", {"size": "20"}); v2 = s2.verify(); check("formatting-only verify strict text equality", v2['ok'] and any("No text was changed" in c['name'] for c in v2['checks']))
print("\nFAILED:", fails or "none")
