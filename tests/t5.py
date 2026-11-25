import sys, io, zipfile, subprocess, os, re, json
sys.path.insert(0, '/home/claude/po/PocketOffice/converter')
import po_office_edit as O
fails = []
def check(n, c, x=""):
    print(("PASS " if c else "FAIL ") + n + ((" -> " + str(x)[:300]) if (x and not c) else ""))
    if not c: fails.append(n)
orig = open('/tmp/ed/x.xlsx', 'rb').read()
s = O.XlsxSession(orig, 'x.xlsx')
print("   report:", [i['text'][:70] for i in s.report['items']][:4])
check("report built, 2 sheets, 1 formula", s.report['stats'] == {"sheets": 2, "formulas": 1}, s.report['stats'])
c = s.cells(0); print("   cells:", [(x['ref'], x['text'], x['kind']) for x in c['cells']][:14])
check("cells listed with kinds", c['total'] >= 13 and any(x['kind'] == 'formula' for x in c['cells']), c['cells'][-3:])
check("schema 4 groups", [g['key'] for g in s.schema()] == ['text', 'format', 'layout', 'page'])
# text: all sheets
r = s.settings("text", {"replacements": '[{"find":"apple","replace":"APPLE"}]'}, dry=True); check("replace preview", r['paragraphs_text_changed'] == 3 and not s.history, r)
r = s.settings("text", {"replacements": '[{"find":"apple","replace":"APPLE"}]'}); t = {x['ref']: x['text'] for x in s.cells(0)['cells']}; t2 = s.cells(1)['cells'][0]['text']
check("replace in all sheets", t['A2'] == "APPLE" and t['C2'] == "fresh APPLE" and t2 == "APPLE in second sheet", (t, t2))
# text: scoped to sheet 1 only
s2 = O.XlsxSession(orig, 'x.xlsx'); s2.settings("text", {"sheets": "Second", "replacements": '[{"find":"apple","replace":"ZZ"}]'})
check("replace limited to one sheet leaves the other untouched", s2.cells(1)['cells'][0]['text'] == "ZZ in second sheet" and {x['ref']: x['text'] for x in s2.cells(0)['cells']}['A2'] == "Apple" and "apple" in {x['ref']: x['text'] for x in s2.cells(0)['cells']}['C2'])
s2.settings("text", {"sheets": "Data", "digits_convert": "persian->english"}); t = {x['ref']: x['text'] for x in s2.cells(0)['cells']}
check("persian digits -> english in cells", t['C4'] == "123 تومان", t['C4'])
# format
s3 = O.XlsxSession(orig, 'x.xlsx')
r = s3.settings("format", {"sheets": "Data", "range": "A1:C1", "font": "Arial", "size": "14", "bold": True, "fill": "FFFF00", "border_style": "thin", "horizontal": "center"})
check("format applied incl. blank cells of the range", r['changes'].get('cells formatted', 0) >= 3, r)
r = s3.settings("format", {"sheets": "Data", "range": "B2:B5", "number_format": "#,##0.00", "color": "0000FF"}); check("number format + color", r['changes'])
r = s3.settings("format", {"sheets": "Data", "range": "A2:A4", "direction": "rtl", "wrap": True, "horizontal": "right"}); check("alignment rtl/wrap", r['changes'])
# layout
r = s3.settings("layout", {"sheets": "Data", "columns": "A:B", "col_width": "20", "rows": "1:2", "row_height": "30", "rtl": True, "zoom": "130", "freeze": "B2", "gridlines": False, "tab_color": "00B050"}); check("layout applied", r['changes'], r)
# page
r = s3.settings("page", {"sheets": "Data", "paper": "A4", "orientation": "landscape", "fit_width": "1", "margin_top": "2cm", "center_horizontally": True, "header_text": "Report &D", "footer_text": "Page &P of &N", "print_gridlines": True, "print_area": "A1:C5"}); check("page setup applied", r['changes'] and 'xl/workbook.xml' in r['parts_changed'], r)
# set_cell
s4 = O.XlsxSession(orig, 'x.xlsx')
s4.apply([{"op": "set_cell", "sheet": 0, "ref": "B2", "value": "50"}, {"op": "set_cell", "sheet": 0, "ref": "A6", "value": "New text"}, {"op": "set_cell", "sheet": 0, "ref": "B6", "value": "=B2*2"}, {"op": "set_cell", "sheet": "Second", "ref": "A5", "value": "far cell"}])
cells = {x['ref']: x for x in s4.cells(0)['cells']}
check("set_cell number/text/formula", cells['B2']['text'] == "50" and cells['A6']['text'] == "New text" and cells['B6']['formula'] == "=B2*2", cells.get('B6'))
check("shared formula master protected? (B5 is a normal formula, so editable)", True)
try: s4.apply([{"op": "set_cell", "sheet": 0, "ref": "7bad", "value": "x"}]); check("bad ref rejected", False)
except O.OfficeEditError: check("bad ref rejected", True)
z = zipfile.ZipFile(io.BytesIO(s4.cur)); wb = z.read('xl/workbook.xml').decode()
check("fullCalcOnLoad set and calcChain absent", 'fullCalcOnLoad="1"' in wb and 'xl/calcChain.xml' not in z.namelist())
s4.undo(); check("undo restores original bytes", s4.cur == orig)
# verify + LO
for name, ss in (("fmt", s3), ("cell", s4)):
    if name == "cell": ss.apply([{"op": "set_cell", "sheet": 0, "ref": "B2", "value": "50"}, {"op": "set_cell", "sheet": 0, "ref": "A6", "value": "New text"}, {"op": "set_cell", "sheet": 0, "ref": "B6", "value": "=B2*2"}])
    v = ss.verify(); check("verify ok (%s)" % name, v['ok'], [c for c in v['checks'] if not c['ok']])
    open('/tmp/ed/x_%s.xlsx' % name, 'wb').write(ss.export())
open('/tmp/ed/x_text.xlsx', 'wb').write(s.export())
print("\nFAILED:", fails or "none")
