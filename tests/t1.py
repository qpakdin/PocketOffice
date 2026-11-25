import sys, importlib.util, io, zipfile, subprocess, os, re, json
sys.path.insert(0, '/home/claude/po/PocketOffice/converter')
import po_docx_edit as E

def new():
    return E.DocxSession(open('/tmp/ed/src.docx', 'rb').read(), 'src.docx')
def texts(s): return [b['text'] for b in s.blocks()]
fails = []
def check(name, cond, extra=""):
    print(("PASS " if cond else "FAIL ") + name + (("  -> " + str(extra)) if (extra and not cond) else ""))
    if not cond: fails.append(name)

# ---- A. text edits
cases = [
 ("insert at start", 1, lambda t: "NEW " + t),
 ("insert in middle", 1, lambda t: t.replace("first", "very first"),),
 ("append at end", 1, lambda t: t + " More."),
 ("delete a word", 1, lambda t: t.replace(" and a normal ending", "")),
 ("replace all text", 1, lambda t: "Completely different text"),
 ("make empty", 1, lambda t: ""),
 ("special chars & < >", 1, lambda t: "a & b < c > d \"q\" 'x' &amp; &#65;"),
 ("tab and newline", 1, lambda t: "col1\tcol2\nline2"),
 ("persian edit", 3, lambda t: t.replace("فارسی", "فارسی ویرایش‌شده") ),
 ("persian append", 3, lambda t: t + " ادامه"),
 ("table cell", 6, lambda t: "Banana"),
 ("heading", 0, lambda t: "Annual Report 2026"),
 ("emoji / astral char", 2, lambda t: t + " 😀"),
]
for name, i, fn in cases:
    s = new(); before = texts(s)
    want = fn(before[i])
    try:
        s.apply([{"op": "set_text", "p": i, "text": want}])
        after = texts(s)
        ok = after[i] == want and all(after[k] == before[k] for k in range(len(before)) if k != i)
        check("set_text: " + name, ok, (after[i], want))
    except E.EditError as e:
        check("set_text: " + name, False, e)

# empty -> text
s = new(); s.apply([{"op": "set_text", "p": 1, "text": ""}]); s.apply([{"op": "set_text", "p": 1, "text": "back again"}])
check("empty paragraph then typed again", texts(s)[1] == "back again", texts(s)[1])

# untouched paragraphs are byte-identical
s = new(); a = s.view()
sc0 = E.Scan(s.cur); L0 = [a[r['start']:r['end']] for r in sc0.paras]
s.apply([{"op": "set_text", "p": 1, "text": "Changed one paragraph"}])
b = s.view(); sc1 = E.Scan(s.cur); L1 = [b[r['start']:r['end']] for r in sc1.paras]
check("all other paragraphs byte-identical", all(L0[k] == L1[k] for k in range(len(L0)) if k != 1) and L0[1] != L1[1])
pre0 = a[:sc0.paras[0]['start']]; pre1 = b[:sc1.paras[0]['start']]
check("document header/namespaces byte-identical", pre0 == pre1)

# ---- B. formatting
s = new()
s.apply([{"op": "format", "p": 1, "start": 5, "end": 10, "bold": True, "color": "FF0000", "size": 16, "font": "Arial"}])
x = s.view(); sc = E.Scan(s.cur); px = x[sc.paras[1]['start']:sc.paras[1]['end']]
check("range format keeps text", texts(s)[1] == "This is the first paragraph with italic text and a normal ending.", texts(s)[1])
check("range format splits runs & sets props", 'w:val="FF0000"' in px and 'w:val="32"' in px and 'w:ascii="Arial"' in px, px[:400])
s.apply([{"op": "format", "p": 2, "bold": True, "italic": True, "underline": True, "highlight": "yellow"}])
check("whole paragraph format", texts(s)[2] == "Centered line")
v = s.verify(); check("verify after formatting ok", v['ok'], v['checks'])
# ---- C. paragraph props
s = new()
s.apply([{"op": "para", "p": 1, "align": "center", "before": 6, "after": 12, "line": 1.5, "left": 1, "first": 0.5}])
b1 = [b for b in s.blocks() if b['i'] == 1][0]
check("align center read back", b1['align'] == 'center', b1)
s.apply([{"op": "para", "p": 3, "align": "left"}])
b3 = [b for b in s.blocks() if b['i'] == 3][0]
check("rtl paragraph visual left read back", b3['align'] == 'left' and b3['rtl'], b3)
s.apply([{"op": "para", "p": 1, "rtl": True}])
check("rtl switch", [b for b in s.blocks() if b['i'] == 1][0]['rtl'])
# ---- D. structure
s = new()
s.apply([{"op": "insert_paragraph_after", "p": 1}]); t = texts(s)
check("insert paragraph", len(t) == 12 and t[2] == "" and t[1].startswith("This is"), t[:4])
s.apply([{"op": "set_text", "p": 2, "text": "I am new"}]); check("new paragraph editable", texts(s)[2] == "I am new")
s.apply([{"op": "delete_paragraph", "p": 2}]); check("delete paragraph", len(texts(s)) == 11 and texts(s)[2] == "Centered line")
s.apply([{"op": "insert_row", "t": 0, "r": 1, "where": "below"}]); check("insert row", s.table_rows() == [4], s.table_rows())
t = texts(s); check("new row empty, others intact", t[8:10] == ["", ""] and t[6:8] == ["Apple", "5"], t)
s.apply([{"op": "delete_row", "t": 0, "r": 2}]); check("delete row", s.table_rows() == [3] and texts(s)[8] == "Pear", texts(s))
try:
    s.apply([{"op": "delete_paragraph", "p": 4}]); check("cannot delete sole cell paragraph", False)
except E.EditError as e: check("cannot delete sole cell paragraph", True)
# ---- E. page setup
s = new()
s.apply([{"op": "page", "scope": "last", "w_cm": 29.7, "h_cm": 21.0, "top": 2, "bottom": 2, "left": 2.5, "right": 2.5}])
pg = s.page_info()[0]
check("page setup landscape", pg['landscape'] and pg['w'] == 16838 and pg['h'] == 11906 and pg['left'] == 1417, pg)
try:
    s.apply([{"op": "page", "scope": "last", "w_cm": 10, "h_cm": 10, "top": 6, "bottom": 6, "left": 1, "right": 1}]); check("reject silly margins", False)
except E.EditError: check("reject silly margins", True)
# ---- F. undo
s = new(); orig = s.cur
s.apply([{"op": "set_text", "p": 1, "text": "x"}]); s.apply([{"op": "para", "p": 1, "align": "right"}]); s.apply([{"op": "page", "scope": "last", "w_cm": 21, "h_cm": 29.7, "top": 2, "bottom": 2, "left": 2, "right": 2}])
s.undo(); s.undo(); s.undo()
check("undo x3 restores original bytes exactly", s.cur == orig)
# ---- G. rollback on error
s = new(); b = s.cur
try: s.apply([{"op": "set_text", "p": 1, "text": "ok"}, {"op": "set_text", "p": 99, "text": "no"}])
except E.EditError: pass
check("failed batch rolls back completely", s.cur == b and not s.ops)
try: s.apply([{"op": "set_text", "p": 1, "text": "bad \x01 char"}]); check("control char rejected", False)
except E.EditError: check("control char rejected", True)
print("\nFAILED:", fails if fails else "none")
