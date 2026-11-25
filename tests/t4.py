import sys, io, zipfile, subprocess, os, re
sys.path.insert(0, '/home/claude/po/PocketOffice/converter')
import po_office_edit as O, po_docx_edit as PDE
fails = []
def check(n, c, x=""):
    print(("PASS " if c else "FAIL ") + n + ((" -> " + str(x)) if (x and not c) else ""))
    if not c: fails.append(n)
orig = open('/tmp/ed/complex.docx', 'rb').read()
s = O.WordSession(orig, 'complex.docx')
check("schema has 13 groups with sections", len(s.schema()) == 13 and all(g['sections'] for g in s.schema()))
print("   groups:", [g['key'] for g in s.schema()])
# preview (dry) does not change anything
r = s.settings("paragraph", {"alignment": "justify", "line_spacing": "1.5", "space_after": "6pt"}, dry=True)
check("dry-run reports changes but changes nothing", r['changes'] and r['parts_changed'] == ['word/document.xml'] and s.cur == orig and not s.history, r)
r = s.settings("paragraph", {"alignment": "justify", "line_spacing": "1.5", "space_after": "6pt"})
check("paragraph settings applied", s.cur != orig and len(s.history) == 1)
r = s.settings("font", {"font": "Arial", "size": "13", "color": "1F3864"})
check("font settings applied (all runs)", r['changes'], r)
r = s.settings("text", {"replacements": '[{"find":"paragraph","replace":"PARA","match_case":false}]'}, dry=True)
check("find/replace preview shows affected paragraphs", r['paragraphs_text_changed'] >= 2 and r['samples'], r)
r = s.settings("text", {"replacements": '[{"find":"paragraph","replace":"PARA","match_case":false}]'})
t = s._dx().full_texts()
check("replace applied incl. keeping other paragraphs", any("PARA" in x for x in t) and not any("paragraph" in x.lower() for x in t if "PARA" not in x), t)
r = s.settings("page", {"paper": "A4", "margin_top": "2.5cm", "orientation": "landscape"})
check("page setup via wordtk", r['changes'], r)
r = s.settings("header_footer", {"footer": '{"center": "{PAGE} / {NUMPAGES}"}'}); check("footer with page fields", 'word/footer' in " ".join(r['parts_added'] + r['parts_changed']) or r['changes'], r)
# precise edit still works after settings, in the same history
s.apply([{"op": "set_text", "p": 1, "text": "Precise edit after settings"}])
check("precise edit after settings", s._dx().full_texts()[1] == "Precise edit after settings")
n = len(s.history)
s.undo(); s.undo(); check("undo steps back through settings and ops", len(s.history) == n - 2)
# errors
for bad, why in (({"alignment": "sideways"}, "bad choice"), ({"size": "abc"}, "bad number")):
    try: s.settings("font" if "size" in bad else "paragraph", bad); check("rejects " + why, False)
    except O.OfficeEditError as e: check("rejects " + why + " clearly", True)
try: s.settings("font", {}); check("empty settings rejected", False)
except O.OfficeEditError: check("empty settings rejected", True)
# find
f = s.find("e"); check("find returns paragraphs with snippets", f['matches'] >= 1 and f['results'], f)
f = s.find("PARA.*", regex=True, match_case=True); check("regex find", f['matches'] >= 1)
try: s.find("(", regex=True); check("invalid regex rejected", False)
except O.OfficeEditError: check("invalid regex rejected", True)
# verify + LibreOffice
v = s.verify()
for c in v['checks']: print("   verify:", "ok " if c['ok'] else "BAD", c['name'], "|", c['detail'][:110])
check("verify passes (text-changing history)", v['ok'])
open('/tmp/ed/word_settings.docx', 'wb').write(s.export())
