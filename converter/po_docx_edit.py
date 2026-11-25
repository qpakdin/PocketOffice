#!/usr/bin/env python3
"""
Surgical DOCX editor for Pocket Office.

Principle: a Word file is a zip of XML parts. We NEVER re-save the document through another program.
Only word/document.xml is changed, and inside it only the exact byte ranges of the paragraphs / rows /
section settings that the user edited. Every other part (styles, theme, charts, SmartArt, images, macros,
headers, footers, fonts, signatures ...) is copied byte for byte, and the rest of document.xml stays
byte-identical, so namespaces, mc:Ignorable, rsid values, shapes, anchors ... cannot be damaged.

Paragraphs that contain anything this editor does not fully understand (fields, pictures, shapes, tracked
changes, content controls, footnote marks, equations ...) are LOCKED: they stay viewable but cannot be edited.

Text trick: document.xml is handled as a latin-1 "view" of its UTF-8 bytes, so string offsets are byte offsets
and non-ASCII text is never re-encoded by accident.
"""
import difflib, hashlib, io, json, re, time, zipfile
from xml.parsers import expat
import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MAIN_PART = "word/document.xml"
TWIPS_PER_CM = 566.9291338
MAX_TEXT = 60000

ORDER_PPR = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd",
             "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi",
             "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
             "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]
ORDER_RPR = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
             "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u",
             "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"]
ORDER_SECTPR = ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType",
                "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings",
                "sectPrChange"]
HIGHLIGHTS = {"yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow",
              "darkGray", "lightGray", "black", "white", "none"}
CONTENT_OK = {"p", "r", "t", "tab", "br", "cr", "noBreakHyphen", "softHyphen", "lastRenderedPageBreak", "proofErr", "bookmarkStart", "bookmarkEnd", "hyperlink"}
FRIENDLY = {"drawing": "a picture, chart or shape", "pict": "a picture or shape (VML)", "object": "an embedded object", "fldChar": "a field (page number, TOC, cross-reference ...)",
            "fldSimple": "a field", "instrText": "a field", "footnoteReference": "a footnote mark", "endnoteReference": "an endnote mark",
            "commentReference": "a comment", "commentRangeStart": "a comment", "commentRangeEnd": "a comment", "ins": "a tracked insertion",
            "del": "a tracked deletion", "moveFrom": "tracked moved text", "moveTo": "tracked moved text", "sdt": "a content control", "oMath": "an equation",
            "oMathPara": "an equation", "smartTag": "a smart tag", "customXml": "custom XML", "AlternateContent": "a shape / compatibility block",
            "txbxContent": "a text box", "sym": "a symbol character", "ruby": "ruby text", "pPrChange": "a tracked paragraph-format change",
            "rPrChange": "a tracked run-format change", "footnoteRef": "a footnote mark", "endnoteRef": "an endnote mark", "annotationRef": "a comment mark",
            "separator": "a separator", "continuationSeparator": "a separator", "pgNum": "a page-number field", "control": "a form control"}


# ----------------------------------------------------------------------------- small helpers (latin-1 "byte view" strings)
def b2u(s):
    """byte-view string -> real unicode text"""
    return s.encode("latin-1").decode("utf-8", "replace")


def u2b(s):
    """real unicode text -> byte-view string"""
    return s.encode("utf-8").decode("latin-1")


_ENT = re.compile(r"&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-fA-F]+);")


def unesc(s):
    def f(m):
        g = m.group(1)
        if g == "lt": return "<"
        if g == "gt": return ">"
        if g == "amp": return "&"
        if g == "quot": return '"'
        if g == "apos": return "'"
        return chr(int(g[2:], 16)) if g[1] in "xX" else chr(int(g[1:]))
    return _ENT.sub(f, s)


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _tag_end(data, i):
    """index just after the '>' that closes the tag starting at data[i] (respects quoted attribute values)"""
    q, n, j = None, len(data), i + 1
    while j < n:
        c = data[j]
        if q is not None:
            if c == q:
                q = None
        elif c == 34 or c == 39:
            q = c
        elif c == 62:
            return j + 1
        j += 1
    return n


def _local(name):
    return name.split(":")[-1]


def _attrs(tag):
    return re.findall(r'([\w:.-]+)\s*=\s*"([^"]*)"', tag)


def split_children(inner):
    """top-level child elements of an XML fragment -> [(local_name, xml)]"""
    out, depth, start, cur = [], 0, None, None
    for m in re.finditer(r"<(/?)([^\s/>!?]+)([^>]*?)(/?)>", inner):
        closing, name, selfc = m.group(1), m.group(2), m.group(4)
        if closing:
            depth -= 1
            if depth == 0 and start is not None:
                out.append((_local(cur), inner[start:m.end()]))
                start = None
        else:
            if depth == 0:
                start, cur = m.start(), name
                if selfc:
                    out.append((_local(name), inner[start:m.end()]))
                    start = None
                    continue
            if not selfc:
                depth += 1
    return out


def split_container(xml):
    """'<w:rPr a="1">..</w:rPr>' / '<w:rPr/>' -> (qname, attrs_text, inner)"""
    m = re.match(r"<([^\s/>]+)((?:\s[^>]*?)?)(/?)>", xml)
    qn, attrs, selfc = m.group(1), m.group(2), m.group(3)
    if selfc:
        return qn, attrs, ""
    return qn, attrs, xml[m.end():xml.rindex("</" + qn)]


def set_child(container, local, child_xml, order, qn_default):
    """Insert / replace / remove one child inside a pPr / rPr / sectPr container, keeping the schema order."""
    if container is None:
        if child_xml is None:
            return None
        container = "<%s/>" % qn_default
    qn, attrs, inner = split_container(container)
    kids = split_children(inner)
    rank = {n: i for i, n in enumerate(order)}
    pos = [i for i, (n, _x) in enumerate(kids) if n == local]
    if child_xml is None:
        kids = [k for i, k in enumerate(kids) if i not in pos]
    elif pos:
        kids[pos[0]] = (local, child_xml)
        kids = [k for i, k in enumerate(kids) if i == pos[0] or i not in pos]
    else:
        r = rank.get(local, 10 ** 6)
        at = len(kids)
        for i, (n, _x) in enumerate(kids):
            if rank.get(n, 10 ** 6) > r:
                at = i
                break
        kids.insert(at, (local, child_xml))
    if not kids:
        return "<%s%s/>" % (qn, attrs)
    return "<%s%s>%s</%s>" % (qn, attrs, "".join(x for _n, x in kids), qn)


def get_child(container, local):
    if container is None:
        return None
    _qn, _a, inner = split_container(container)
    for n, x in split_children(inner):
        if n == local:
            return x
    return None


def merge_attrs(elem_xml, qname, updates):
    """update attributes of an empty element (spacing / ind / pgSz ...), keep the others in their order"""
    cur = _attrs(elem_xml) if elem_xml else []
    seen, out = set(), []
    for k, v in cur:
        kl = _local(k)
        if kl in updates:
            seen.add(kl)
            if updates[kl] is not None:
                out.append((k, str(updates[kl])))
        else:
            out.append((k, v))
    pfx = qname.rsplit(":", 1)[0] + ":" if ":" in qname else ""
    for k, v in updates.items():
        if k not in seen and v is not None:
            out.append((pfx + k, str(v)))
    return "<%s%s/>" % (qname, "".join(' %s="%s"' % (k, v) for k, v in out))


# ----------------------------------------------------------------------------- scanning document.xml
class Scan:
    """Byte ranges of paragraphs, tables, rows, cells, section properties."""

    def __init__(self, data):
        self.data = data                       # bytes
        self.pf = self._prefix(data)
        self.paras, self.tables, self.sects, self.bodies = [], [], [], []
        self.sdt_bound = 0
        self.alt_blocks = 0
        self._scan()

    @staticmethod
    def _prefix(data):
        found = {"pf": "w:"}

        class Stop(Exception):
            pass

        def start(name, attrs):
            for k, v in attrs.items():
                if v == W_NS:
                    if k == "xmlns":
                        found["pf"] = ""
                    elif k.startswith("xmlns:"):
                        found["pf"] = k[6:] + ":"
            raise Stop()
        p = expat.ParserCreate()
        p.StartElementHandler = start
        try:
            p.Parse(data, True)
        except Stop:
            pass
        except expat.ExpatError:
            pass
        return found["pf"]

    def _scan(self):
        data, P = self.data, self.pf
        stack = []
        parser = expat.ParserCreate()
        scan = self

        def nearest(name):
            for f in reversed(stack):
                if f["name"] == name:
                    return f
            return None

        def start(name, attrs):
            idx = parser.CurrentByteIndex
            te = _tag_end(data, idx)
            fr = {"name": name, "start": idx, "istart": te, "self": data[te - 2:te] == b"/>", "children": []}
            if name == P + "p":
                tc, tbl, tr = nearest(P + "tc"), nearest(P + "tbl"), nearest(P + "tr")
                anc = [f["name"] for f in stack]
                rec = {"start": idx, "istart": te, "container": "tc" if tc is not None and (tbl is None or tc["start"] > tbl["start"]) else "body",
                       "in_txbx": (P + "txbxContent") in anc, "in_fallback": any(_local(n) == "Fallback" for n in anc),
                       "in_sdt": (P + "sdt") in anc, "tc": tc, "tbl_i": tbl["tbl_i"] if tbl else None,
                       "row_i": tr["row_i"] if tr else None, "col_i": tc["col_i"] if tc else None}
                fr["rec"] = rec
                scan.paras.append(rec)
            elif name == P + "tbl":
                fr["tbl_i"] = len(scan.tables)
                fr["rows"] = []
                fr["rec"] = {"start": idx, "rows": fr["rows"]}
                scan.tables.append(fr["rec"])
            elif name == P + "tr":
                tbl = nearest(P + "tbl")
                fr["row_i"] = len(tbl["rows"]) if tbl else 0
                fr["cells"] = []
                fr["rec"] = {"start": idx, "cells": fr["cells"]}
                if tbl:
                    tbl["rows"].append(fr["rec"])
            elif name == P + "tc":
                tr = nearest(P + "tr")
                fr["col_i"] = len(tr["cells"]) if tr else 0
                fr["rec"] = {"start": idx}
                if tr:
                    tr["cells"].append(fr["rec"])
            elif name == P + "sectPr":
                fr["rec"] = {"start": idx, "istart": te, "body_level": bool(stack and stack[-1]["name"] == P + "body")}
                scan.sects.append(fr["rec"])
            elif name == P + "dataBinding":
                sdt = nearest(P + "sdt")
                if sdt is not None:
                    sdt["bound"] = True
            elif _local(name) == "AlternateContent":
                scan.alt_blocks += 1
            stack.append(fr)

        def end(name):
            fr = stack.pop()
            idx = parser.CurrentByteIndex
            if fr["self"]:
                fr["end"], fr["iend"] = fr["istart"], fr["istart"]
            else:
                fr["iend"], fr["end"] = idx, _tag_end(data, idx)
            rec = fr.get("rec")
            if rec is not None:
                rec["end"], rec["iend"] = fr["end"], fr["iend"]
            if name == P + "sdt" and fr.get("bound"):
                scan.sdt_bound += 1
            if name == P + "body":
                scan.bodies.append(fr)
            if name == P + "tc" and rec is not None:
                rec["children"] = fr["children"]
            if stack:
                par = stack[-1]
                if par["name"] in (P + "body", P + "tc"):
                    par["children"].append({"loc": _local(name), "start": fr["start"], "end": fr["end"]})

        parser.StartElementHandler = start
        parser.EndElementHandler = end
        parser.Parse(data, True)


# ----------------------------------------------------------------------------- paragraph model
RISKY_NAME = re.compile(r"<([A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)")


def para_lock_reason(pxml, pf):
    """None if the paragraph is simple enough to edit safely, else a human reason."""
    if re.search(r"<%s(pPrChange|rPrChange|sectPrChange|numberingChange)[\s>/]" % re.escape(pf), pxml):
        return "tracked formatting change"
    body = re.sub(r"<%spPr(?=[\s>/])(?:/>|>.*?</%spPr>)" % (re.escape(pf), re.escape(pf)), "", pxml, flags=re.S)
    body = re.sub(r"<%srPr(?=[\s>/])(?:/>|>.*?</%srPr>)" % (re.escape(pf), re.escape(pf)), "", body, flags=re.S)
    for m in RISKY_NAME.finditer(body[2:]):
        loc = m.group(2)
        if loc not in CONTENT_OK:
            return FRIENDLY.get(loc, "unsupported element <%s>" % loc)
    return None


def run_regex(pf):
    e = re.escape(pf)
    return re.compile(r"<%sr(?=[\s>/])([^>]*?)(?:/>|>(.*?)</%sr>)" % (e, e), re.S)


def parse_run(m, pf):
    e = re.escape(pf)
    attrs, content = m.group(1), m.group(2) or ""
    rpr, rest = None, content
    mm = re.match(r"\s*(<%srPr(?=[\s>/])(?:/>|>.*?</%srPr>))" % (e, e), content, re.S)
    if mm:
        rpr, rest = mm.group(1), content[mm.end():]
    items = []
    tok = re.compile(r"<%st(?:\s[^>]*)?>(.*?)</%st>|<%st(?:\s[^>]*)?/>|<%stab\s*/>|<%sbr(?:\s[^>]*)?/>|<%scr\s*/>|<%snoBreakHyphen\s*/>|<%ssoftHyphen\s*/>|<%slastRenderedPageBreak\s*/>|\s+" % ((e,) * 9), re.S)
    pos = 0
    while pos < len(rest):
        t = tok.match(rest, pos)
        if not t:
            return None
        s = t.group(0)
        if not s.strip():
            pass
        elif s.startswith("<%stab" % pf):
            items.append({"k": "tab"})
        elif s.startswith("<%sbr" % pf):
            typ = re.search(r'%stype="([^"]*)"' % re.escape(pf), s)
            items.append({"k": "br", "raw": s, "ch": "\n" if (not typ or typ.group(1) == "textWrapping") else "\x0c"})
        elif s.startswith("<%scr" % pf):
            items.append({"k": "cr"})
        elif s.startswith("<%snoBreakHyphen" % pf):
            items.append({"k": "nbh"})
        elif s.startswith("<%ssoftHyphen" % pf):
            items.append({"k": "soft"})
        elif s.startswith("<%slastRendered" % pf):
            items.append({"k": "lrpb"})
        else:
            items.append({"k": "t", "x": unesc(b2u(t.group(1))) if t.group(1) is not None else ""})
        pos = t.end()
    return {"s": m.start(), "e": m.end(), "attrs": attrs, "rpr": rpr, "items": items, "raw": m.group(0)}


CH = {"tab": "\t", "cr": "\n", "nbh": "\u2011", "soft": "\u00ad", "lrpb": ""}


def item_text(it):
    if it["k"] == "t":
        return it["x"]
    if it["k"] == "br":
        return it["ch"]
    return CH[it["k"]]


def parse_paragraph(pxml, pf):
    """-> (open_tag, ppr, runs, close_tag_start, reason)  runs are byte-view strings inside pxml"""
    om = re.match(r"<[^>]*>", pxml)
    if pxml[om.end() - 2:om.end()] == "/>":
        return {"selfclosed": True, "open": pxml, "ppr": None, "runs": [], "text": "", "pstart": 0, "pend": 0}
    e = re.escape(pf)
    ppr = None
    pm = re.match(r"\s*(<%spPr(?=[\s>/])(?:/>|>.*?</%spPr>))" % (e, e), pxml[om.end():], re.S)
    ppr_span = None
    if pm:
        ppr = pm.group(1)
        ppr_span = (om.end() + pm.start(1), om.end() + pm.end(1))
    runs = []
    for m in run_regex(pf).finditer(pxml):
        r = parse_run(m, pf)
        if r is None:
            return None
        runs.append(r)
    text = "".join(item_text(it) for r in runs for it in r["items"])
    return {"selfclosed": False, "open": om.group(0), "ppr": ppr, "ppr_span": ppr_span, "runs": runs, "text": text, "close": pxml.rindex("</")}


def para_text_unicode(pxml, pf):
    """Text of a paragraph (real unicode); '' if the structure is not understood."""
    info = parse_paragraph(pxml, pf)
    return "" if info is None else info["text"]


def run_to_xml(attrs, rpr, items, pf):
    parts = []
    for it in items:
        k = it["k"]
        if k == "t":
            x = it["x"]
            sp = ' xml:space="preserve"' if (x != x.strip() or "  " in x) else ""
            parts.append("<%st%s>%s</%st>" % (pf, sp, u2b(esc(x)), pf))
        elif k == "tab":
            parts.append("<%stab/>" % pf)
        elif k == "br":
            parts.append(it["raw"])
        elif k == "cr":
            parts.append("<%scr/>" % pf)
        elif k == "nbh":
            parts.append("<%snoBreakHyphen/>" % pf)
        elif k == "soft":
            parts.append("<%ssoftHyphen/>" % pf)
        elif k == "lrpb":
            parts.append("<%slastRenderedPageBreak/>" % pf)
    inner = (rpr or "") + "".join(parts)
    if not inner:
        return "<%sr%s/>" % (pf, attrs)
    return "<%sr%s>%s</%sr>" % (pf, attrs, inner, pf)


def text_to_items(txt):
    """unicode text -> run items (tab / line break become real elements)"""
    items, buf = [], []

    def flush():
        if buf:
            items.append({"k": "t", "x": "".join(buf)})
            buf.clear()
    for ch in txt:
        if ch == "\t":
            flush(); items.append({"k": "tab"})
        elif ch == "\n":
            flush(); items.append({"k": "br", "raw": "<w:br/>", "ch": "\n"})
        elif ch == "\x0c":
            flush(); items.append({"k": "br", "raw": '<w:br w:type="page"/>', "ch": "\x0c"})
        else:
            buf.append(ch)
    flush()
    return items


class EditError(Exception):
    pass


def check_text(t):
    if len(t) > MAX_TEXT:
        raise EditError("text is longer than %d characters" % MAX_TEXT)
    for ch in t:
        o = ord(ch)
        if o < 32 and ch not in "\t\n\x0c":
            raise EditError("the text contains a control character (code %d) that Word cannot store" % o)
        if 0xD800 <= o <= 0xDFFF or o in (0xFFFE, 0xFFFF):
            raise EditError("the text contains an invalid character")


def edit_runs_text(pxml, pf, new_text):
    """Return paragraph xml with its text changed to new_text; all untouched runs stay byte-identical."""
    info = parse_paragraph(pxml, pf)
    if info is None:
        raise EditError("paragraph structure not understood")
    runs = info["runs"]
    old = "".join(item_text(it) if it["k"] != "t" else it["x"] for r in runs for it in r["items"])
    if old == new_text:
        return pxml
    sm = difflib.SequenceMatcher(None, old, new_text, autojunk=False)
    deleted = [False] * len(old)
    ins = {}
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag in ("delete", "replace"):
            for k in range(i1, i2):
                deleted[k] = True
        if tag in ("insert", "replace"):
            ins[i1] = ins.get(i1, "") + new_text[j1:j2]
    out_runs = []
    ci = 0
    first_char_owner_done = False
    for r in runs:
        new_items, changed = [], False
        for it in r["items"]:
            L = len(item_text(it))
            lo, hi = ci, ci + L
            if it["k"] == "lrpb" or L == 0:
                new_items.append(it)
                continue
            pieces = []
            if lo == 0 and 0 in ins and not first_char_owner_done:
                pieces.append(("ins", ins[0]))
                first_char_owner_done = True
            if it["k"] == "t":
                for k in range(lo, hi):
                    if not deleted[k]:
                        pieces.append(("old", old[k]))
                    if (k + 1) in ins:
                        pieces.append(("ins", ins[k + 1]))
                # rebuild text/items
                cur = []
                res = []
                for kind, s in pieces:
                    for ch in s:
                        if ch in "\t\n\x0c" and kind == "ins":
                            if cur:
                                res.append({"k": "t", "x": "".join(cur)}); cur = []
                            res.extend(text_to_items(ch))
                        else:
                            cur.append(ch)
                if cur:
                    res.append({"k": "t", "x": "".join(cur)})
                new_items.extend(res)
                if not (len(res) == 1 and res[0]["k"] == "t" and res[0]["x"] == it["x"]):
                    changed = True
            else:
                keep = not deleted[lo]
                if pieces and pieces[0][0] == "ins":
                    new_items.extend(text_to_items(pieces[0][1]))
                    changed = True
                if keep:
                    new_items.append(it)
                else:
                    changed = True
                if hi in ins:
                    new_items.extend(text_to_items(ins[hi]))
                    changed = True
            ci = hi
        out_runs.append((r, new_items, changed))
    # insertion into a paragraph without any characters
    if len(old) == 0 and 0 in ins:
        if out_runs:
            r, its, _c = out_runs[0]
            out_runs[0] = (r, its + text_to_items(ins[0]), True)
        else:
            out_runs.append((None, text_to_items(ins[0]), True))
    # splice back
    res = pxml
    adds = ""
    for r, its, changed in reversed(out_runs):
        if r is None:
            adds = run_to_xml("", _mark_rpr(info, pf), its, pf)
            continue
        if not changed:
            continue
        xml = run_to_xml(r["attrs"], r["rpr"], its, pf) if its else ""
        res = res[:r["s"]] + xml + res[r["e"]:]
    if adds:
        close = res.rindex("</")
        res = res[:close] + adds + res[close:]
    return res


def _mark_rpr(info, pf):
    """rPr of the paragraph mark (used for text typed into an empty paragraph)"""
    if info["ppr"]:
        r = get_child(info["ppr"], "rPr")
        if r:
            return r
    return None


# ----------------------------------------------------------------------------- formatting
def _on(pf, tag, val):
    return "<%s%s/>" % (pf, tag) if val else '<%s%s %sval="0"/>' % (pf, tag, pf)


def apply_run_props(rpr, props, pf):
    """props: font,size,bold,italic,underline,strike,color,highlight (None = leave)"""
    def put(r, tag, xml):
        return set_child(r, tag, xml, ORDER_RPR, pf + "rPr")
    if props.get("font"):
        f = u2b(esc(props["font"]).replace('"', "&quot;"))
        r = put(rpr, "rFonts", '<%srFonts %sascii="%s" %shAnsi="%s" %seastAsia="%s" %scs="%s"/>' % (pf, pf, f, pf, f, pf, f, pf, f))
        rpr = r
    if props.get("size"):
        hp = int(round(float(props["size"]) * 2))
        if not 2 <= hp <= 3276:
            raise EditError("font size must be between 1 and 1638 pt")
        rpr = put(rpr, "sz", '<%ssz %sval="%d"/>' % (pf, pf, hp))
        rpr = put(rpr, "szCs", '<%sszCs %sval="%d"/>' % (pf, pf, hp))
    for key, tag in (("bold", "b"), ("italic", "i")):
        if props.get(key) is not None:
            rpr = put(rpr, tag, _on(pf, tag, props[key]))
            rpr = put(rpr, tag + "Cs", _on(pf, tag + "Cs", props[key]))
    if props.get("underline") is not None:
        rpr = put(rpr, "u", '<%su %sval="%s"/>' % (pf, pf, "single" if props["underline"] else "none"))
    if props.get("strike") is not None:
        rpr = put(rpr, "strike", _on(pf, "strike", props["strike"]))
    if props.get("color"):
        c = str(props["color"]).lstrip("#")
        if c.lower() != "auto" and not re.fullmatch(r"[0-9a-fA-F]{6}", c):
            raise EditError("colour must be like FF0000")
        rpr = put(rpr, "color", '<%scolor %sval="%s"/>' % (pf, pf, "auto" if c.lower() == "auto" else c.upper()))
    if props.get("highlight"):
        h = props["highlight"]
        if h not in HIGHLIGHTS:
            raise EditError("unknown highlight colour")
        rpr = put(rpr, "highlight", None if h == "none" else '<%shighlight %sval="%s"/>' % (pf, pf, h))
    return rpr


def split_items_at(items, k):
    """split a run's items at char offset k -> (left, right)"""
    left, right, cur = [], [], 0
    for it in items:
        L = len(item_text(it))
        if cur + L <= k and not (L == 0 and cur >= k):
            left.append(it)
        elif cur >= k:
            right.append(it)
        else:  # only text items can straddle
            cut = k - cur
            left.append({"k": "t", "x": it["x"][:cut]})
            right.append({"k": "t", "x": it["x"][cut:]})
        cur += L
    return left, right


def format_paragraph_range(pxml, pf, start, end, props):
    info = parse_paragraph(pxml, pf)
    if info is None:
        raise EditError("paragraph structure not understood")
    runs = info["runs"]
    total = sum(len(item_text(it)) for r in runs for it in r["items"])
    whole = start is None or end is None
    if whole:
        start, end = 0, total
    if not 0 <= start <= end <= total:
        raise EditError("selection is outside the paragraph")
    res = pxml
    if not runs or total == 0:
        # empty paragraph: format the paragraph mark so new typing inherits it
        mark = _mark_rpr(info, pf)
        new_mark = apply_run_props(mark, props, pf)
        return _set_ppr_child(pxml, pf, "rPr", new_mark)
    if start == end:
        raise EditError("nothing selected")
    ci = 0
    spans = []
    for r in runs:
        L = sum(len(item_text(it)) for it in r["items"])
        spans.append((ci, ci + L))
        ci += L
    for idx in range(len(runs) - 1, -1, -1):
        r, (rs, re_) = runs[idx], spans[idx]
        if re_ <= start or rs >= end or rs == re_:
            continue
        a, b = max(start, rs) - rs, min(end, re_) - rs
        left, rest = split_items_at(r["items"], a)
        mid, right = split_items_at(rest, b - a)
        new = ""
        if left:
            new += run_to_xml(r["attrs"], r["rpr"], left, pf)
        new += run_to_xml(r["attrs"], apply_run_props(r["rpr"], props, pf), mid, pf)
        if right:
            new += run_to_xml(r["attrs"], r["rpr"], right, pf)
        if not left and not right and r["rpr"] == apply_run_props(r["rpr"], props, pf):
            continue
        res = res[:r["s"]] + new + res[r["e"]:]
    return res


def _set_ppr_child(pxml, pf, local, child_xml, remove=False):
    e = re.escape(pf)
    om = re.match(r"<[^>]*>", pxml)
    if pxml[om.end() - 2:om.end()] == "/>":
        # <w:p/> -> open it
        opening = pxml[:om.end() - 2].rstrip() + ">"
        pxml = opening + "</%sp>" % pf
        om = re.match(r"<[^>]*>", pxml)
    pm = re.match(r"\s*(<%spPr(?=[\s>/])(?:/>|>.*?</%spPr>))" % (e, e), pxml[om.end():], re.S)
    ppr = pm.group(1) if pm else None
    new = set_child(ppr, local, None if remove else child_xml, ORDER_PPR, pf + "pPr")
    if pm:
        s, en = om.end() + pm.start(1), om.end() + pm.end(1)
        return pxml[:s] + (new or "") + pxml[en:]
    if new is None:
        return pxml
    return pxml[:om.end()] + new + pxml[om.end():]


def apply_para_props(pxml, pf, props):
    info = parse_paragraph(pxml, pf)
    ppr = info["ppr"] if info and not info.get("selfclosed") else None
    cur = ppr
    bidi = get_child(cur, "bidi") is not None and 'val="0"' not in get_child(cur, "bidi") and 'val="false"' not in get_child(cur, "bidi")
    if props.get("rtl") is not None:
        cur = set_child(cur, "bidi", ("<%sbidi/>" % pf) if props["rtl"] else '<%sbidi %sval="0"/>' % (pf, pf), ORDER_PPR, pf + "pPr")
        bidi = bool(props["rtl"])
    if props.get("align"):
        a = props["align"]
        if a not in ("left", "center", "right", "justify"):
            raise EditError("alignment must be left, center, right or justify")
        if bidi and a in ("left", "right"):       # in right-to-left paragraphs Word swaps left/right
            a = "right" if a == "left" else "left"
        val = "both" if a == "justify" else a
        cur = set_child(cur, "jc", '<%sjc %sval="%s"/>' % (pf, pf, val), ORDER_PPR, pf + "pPr")
    sp = {}
    for key, name, mul in (("before", "before", 20), ("after", "after", 20)):
        if props.get(key) is not None:
            v = float(props[key])
            if not 0 <= v <= 1584:
                raise EditError("spacing must be between 0 and 1584 pt")
            sp[name] = int(round(v * mul))
    if props.get("line") is not None:
        v = float(props["line"])
        if not 0.5 <= v <= 10:
            raise EditError("line spacing must be between 0.5 and 10")
        sp["line"] = int(round(v * 240))
        sp["lineRule"] = "auto"
    if sp:
        old = get_child(cur, "spacing")
        for stale in ("beforeAutospacing", "afterAutospacing", "beforeLines", "afterLines"):
            if ("before" in sp and stale.startswith("before")) or ("after" in sp and stale.startswith("after")):
                sp[stale] = None
        cur = set_child(cur, "spacing", merge_attrs(old, pf + "spacing", sp), ORDER_PPR, pf + "pPr")
    ind = {}
    for key, name in (("left", "left"), ("right", "right"), ("first", "firstLine")):
        if props.get(key) is not None:
            v = float(props[key])
            if abs(v) > 50:
                raise EditError("indent must be within 50 cm")
            ind[name] = int(round(v * TWIPS_PER_CM))
    if ind:
        if "firstLine" in ind:
            ind["hanging"] = None
            if ind["firstLine"] < 0:
                ind["hanging"], ind["firstLine"] = -ind["firstLine"], None
        old = get_child(cur, "ind")
        for stale in ("start", "end", "leftChars", "rightChars", "firstLineChars", "hangingChars"):
            if (("left" in ind and stale in ("start", "leftChars")) or ("right" in ind and stale in ("end", "rightChars"))
                    or ("firstLine" in ind and stale in ("firstLineChars", "hangingChars"))):
                ind[stale] = None
        cur = set_child(cur, "ind", merge_attrs(old, pf + "ind", ind), ORDER_PPR, pf + "pPr")
    if cur == ppr:
        return pxml
    return _replace_ppr(pxml, pf, cur)


def _replace_ppr(pxml, pf, new_ppr):
    e = re.escape(pf)
    om = re.match(r"<[^>]*>", pxml)
    if pxml[om.end() - 2:om.end()] == "/>":
        pxml = pxml[:om.end() - 2].rstrip() + "></%sp>" % pf
        om = re.match(r"<[^>]*>", pxml)
    pm = re.match(r"\s*(<%spPr(?=[\s>/])(?:/>|>.*?</%spPr>))" % (e, e), pxml[om.end():], re.S)
    if pm:
        return pxml[:om.end() + pm.start(1)] + new_ppr + pxml[om.end() + pm.end(1):]
    return pxml[:om.end()] + new_ppr + pxml[om.end():]


def paragraph_props_summary(pxml, pf):
    info = parse_paragraph(pxml, pf)
    d = {"style": None, "align": "left", "rtl": False, "font": None, "size": None, "bold": False, "italic": False, "underline": False, "color": None}
    if not info or info.get("selfclosed"):
        return d
    ppr = info["ppr"]
    if ppr:
        ps = get_child(ppr, "pStyle")
        if ps:
            m = re.search(r'val="([^"]*)"', ps)
            d["style"] = m.group(1) if m else None
        b = get_child(ppr, "bidi")
        d["rtl"] = bool(b) and 'val="0"' not in b and 'val="false"' not in b
        j = get_child(ppr, "jc")
        if j:
            m = re.search(r'val="([^"]*)"', j)
            v = m.group(1) if m else "left"
            v = {"both": "justify", "start": "left", "end": "right", "distribute": "justify"}.get(v, v)
            if d["rtl"] and v in ("left", "right"):
                v = "right" if v == "left" else "left"
            d["align"] = v
        elif d["rtl"]:
            d["align"] = "right"
    for r in info["runs"]:
        rp = r["rpr"]
        if not rp or not any(it["k"] == "t" and it["x"] for it in r["items"]):
            continue
        f = get_child(rp, "rFonts")
        if f:
            m = re.search(r'(?:ascii|cs)="([^"]*)"', f) or re.search(r'="([^"]*)"', f)
            d["font"] = b2u(m.group(1)) if m else None
        s = get_child(rp, "sz")
        if s:
            m = re.search(r'val="(\d+)"', s)
            d["size"] = int(m.group(1)) / 2.0 if m else None
        for key, tag in (("bold", "b"), ("italic", "i")):
            x = get_child(rp, tag)
            d[key] = bool(x) and 'val="0"' not in x and 'val="false"' not in x
        u = get_child(rp, "u")
        d["underline"] = bool(u) and 'val="none"' not in u
        c = get_child(rp, "color")
        if c:
            m = re.search(r'val="([0-9A-Fa-f]{6})"', c)
            d["color"] = m.group(1).upper() if m else None
        break
    return d


# ----------------------------------------------------------------------------- the editing session
def _wellformed(data):
    p = expat.ParserCreate()
    try:
        p.Parse(data, True)
        ET.fromstring(data)                 # also checks that every prefix is declared
        return None
    except (expat.ExpatError, ET.ParseError) as e:
        return str(e)


def _open_selfclosed_p(pxml, pf):
    om = re.match(r"<[^>]*>", pxml)
    if pxml[om.end() - 2:om.end()] == "/>":
        return pxml[:om.end() - 2].rstrip() + "></%sp>" % pf
    return pxml


def display_text(pxml, pf):
    """readable text of ANY paragraph (also locked ones) with markers for pictures / fields"""
    e = re.escape(pf)
    out = []
    for m in re.finditer(r"<%st(?:\s[^>]*)?>(.*?)</%st>|<%stab\s*/>|<%sbr(?:\s[^>]*)?/>|<%s(drawing|pict|object)[\s>/]" % ((e,) * 5), pxml, re.S):
        g = m.group(0)
        if g.startswith("<%st" % pf) and not g.startswith("<%stab" % pf):
            out.append(unesc(b2u(m.group(1) or "")))
        elif g.startswith("<%stab" % pf):
            out.append("\t")
        elif g.startswith("<%sbr" % pf):
            out.append("\n")
        else:
            out.append(" \u25a3 ")
    return "".join(out)


class DocxSession:
    def __init__(self, orig, name="document.docx", log=None, analyze=True):
        self.orig, self.name = orig, name
        self.log = log or (lambda m: None)
        self.sha = hashlib.sha256(orig).hexdigest()
        self.ops, self.undone = [], []
        self.blocked = None
        self.zin, self.doc0 = None, b""
        try:
            self.zin = zipfile.ZipFile(io.BytesIO(orig))
            self.doc0 = self.zin.read(MAIN_PART)
        except (zipfile.BadZipFile, KeyError):
            self.blocked = "This is not a readable .docx package (it may be password-protected / encrypted, or damaged)."
        if not self.blocked and self.doc0.startswith((b"\xff\xfe", b"\xfe\xff")):
            self.blocked = "word/document.xml is UTF-16 encoded; this editor only handles UTF-8 documents."
        if not self.blocked:
            err = _wellformed(self.doc0)
            if err:
                self.blocked = "word/document.xml is not well-formed XML (%s)." % err
        self.cur = self.doc0
        self._scan = None
        self.report = self._analyze() if analyze else {"blocked": self.blocked, "items": [], "stats": {}, "locked": {}, "page": None}

    # ---- state
    def scan(self):
        if self._scan is None or self._scan.data is not self.cur:
            self._scan = Scan(self.cur)
        return self._scan

    def view(self):
        return self.cur.decode("latin-1")

    def _listed(self, sc):
        return [r for r in sc.paras if not r["in_fallback"]]

    def _rec(self, sc, i):
        lst = self._listed(sc)
        if not isinstance(i, int) or not 0 <= i < len(lst):
            raise EditError("paragraph %s does not exist" % i)
        return lst[i]

    def _lock_reason(self, s, sc, rec):
        if rec["in_txbx"]:
            return "inside a text box / shape"
        return para_lock_reason(s[rec["start"]:rec["end"]], sc.pf)

    # ---- analysis / report
    def _analyze(self):
        R = {"name": self.name, "sha256": self.sha, "size": len(self.orig), "blocked": self.blocked, "items": [], "stats": {}, "locked": {}, "page": None}
        if self.blocked:
            R["items"].append({"level": "BLOCK", "text": self.blocked})
            return R
        z, names = self.zin, self.zin.namelist()
        doc = self.doc0.decode("latin-1")
        sc = self.scan()
        pf = sc.pf
        sett = ""
        try:
            sett = z.read("word/settings.xml").decode("latin-1")
        except KeyError:
            pass

        def add(level, text):
            R["items"].append({"level": level, "text": text})

        def cnt(rx, hay=doc):
            return len(re.findall(rx, hay))
        # blockers
        if re.search(r"<%sdocumentProtection[^>]*%senforcement=\"(1|true|on)\"" % (re.escape(pf), re.escape(pf)), sett):
            R["blocked"] = "The document is protected against editing (document protection is enforced). Editing it here would bypass that protection."
            add("BLOCK", R["blocked"])
        if any(n.endswith("encryptedPackage") for n in names):
            R["blocked"] = "The document is encrypted."
            add("BLOCK", R["blocked"])
        # warnings
        if any(n.startswith("_xmlsignatures/") for n in names):
            add("WARN", "The document has a DIGITAL SIGNATURE. Any change makes the signature invalid.")
        if "word/vbaProject.bin" in names:
            add("WARN", "The document contains MACROS (kept untouched). If a macro depends on the text you change, it may stop working.")
        ins_n, del_n = cnt(r"<%sins\s" % re.escape(pf)), cnt(r"<%sdel\s" % re.escape(pf))
        if ins_n or del_n:
            add("WARN", "Tracked changes exist (%d insertions, %d deletions). Paragraphs containing them are locked; text elsewhere is edited normally." % (ins_n, del_n))
        if re.search(r"<%strackRevisions" % re.escape(pf), sett):
            add("WARN", "'Track changes' is ON in this document, but edits made here are NOT recorded as tracked changes.")
        if sc.sdt_bound:
            add("WARN", "%d content control(s) are bound to custom XML data. Word may overwrite edited text in them when the file is opened." % sc.sdt_bound)
        fields = cnt(r"<%sfldChar\s[^>]*%sfldCharType=\"begin\"" % (re.escape(pf), re.escape(pf))) + cnt(r"<%sfldSimple\s" % re.escape(pf))
        if fields:
            toc = cnt(r"TOC\s")
            add("WARN", "%d field(s) found%s. Fields are not recalculated here: page numbers / table of contents / cross-references update only when the file is opened in Word and fields are refreshed. Paragraphs containing fields are locked." % (fields, " (including a table of contents)" if toc else ""))
        shapes = cnt(r"<(?:\w+:)?wsp[\s>]") + cnt(r"<(?:\w+:)?txbxContent[\s>]") // 2
        if shapes:
            add("WARN", "%d text box / shape object(s). Their text is locked, but changing the length of nearby paragraphs can move content that is anchored to them." % shapes)
        charts = len([n for n in names if n.startswith("word/charts/") and n.endswith(".xml") and "/_rels/" not in n])
        smart = len([n for n in names if n.startswith("word/diagrams/") and "data" in n and n.endswith(".xml")])
        ole = len([n for n in names if n.startswith("word/embeddings/")])
        if charts or smart or ole:
            add("WARN", "The document has %s%s%s. They are kept byte-for-byte, but they cannot be edited here, and their position can shift if you change text above them." % (
                ("%d chart(s)" % charts) if charts else "", (", %d SmartArt diagram(s)" % smart) if smart else "", (", %d embedded object(s)" % ole) if ole else ""))
        if sc.alt_blocks:
            add("INFO", "%d compatibility block(s) (mc:AlternateContent) – paragraphs inside the fallback copy are hidden from the list and left untouched." % sc.alt_blocks)
        nsect = len(sc.sects)
        if nsect > 1:
            add("WARN", "The document has %d sections (different page setups possible). Page setup lets you choose 'last section only' or 'all sections'." % nsect)
        cols = cnt(r"<%scols[^>]*%snum=\"([2-9])\"" % (re.escape(pf), re.escape(pf)))
        if cols:
            add("INFO", "Multi-column layout is used; it is preserved.")
        m = re.search(r'compatibilityMode"[^>]*val="(\d+)"', sett)
        if m and int(m.group(1)) < 15:
            add("INFO", "Compatibility mode (Word %s layout rules); preserved as it is." % {"14": "2010", "12": "2007", "11": "2003"}.get(m.group(1), "older"))
        hf = [n for n in names if re.match(r"word/(header|footer)\d*\.xml", n)]
        if hf:
            add("INFO", "%d header/footer part(s): they are kept, but cannot be edited in this version." % len(hf))
        notes = [n for n in names if n in ("word/footnotes.xml", "word/endnotes.xml", "word/comments.xml")]
        if notes:
            add("INFO", "Footnotes / endnotes / comments exist: kept, not editable here.")
        media = len([n for n in names if n.startswith("word/media/")])
        if media:
            add("INFO", "%d image(s): kept; paragraphs holding pictures are locked." % media)
        if any(n.startswith("customXml/") for n in names):
            add("INFO", "Custom XML parts present (kept untouched).")
        if any(n.startswith("word/fonts/") for n in names):
            add("INFO", "Embedded fonts present (kept untouched).")
        if re.search(r"<%saltChunk" % re.escape(pf), doc):
            add("WARN", "The document embeds other documents (altChunk). They are kept untouched.")
        # paragraph statistics
        lst = self._listed(sc)
        locked = {}
        ed = 0
        for rec in lst:
            r = self._lock_reason(doc, sc, rec)
            if r:
                locked[r] = locked.get(r, 0) + 1
            else:
                ed += 1
        R["stats"] = {"paragraphs": len(lst), "editable": ed, "locked": len(lst) - ed, "tables": len(sc.tables), "sections": nsect}
        R["locked"] = locked
        if locked:
            add("INFO", "%d of %d paragraphs are LOCKED (they stay as they are): %s." % (len(lst) - ed, len(lst), "; ".join("%s: %d" % (k, v) for k, v in sorted(locked.items(), key=lambda kv: -kv[1]))))
        R["page"] = self.page_info()
        add("OK", "Your original file is never touched: edits happen in a copy and are saved with 'Save as'. Every part of the file other than the edited paragraphs is copied byte for byte (styles, theme, shapes, charts, macros, headers, footers, fonts).")
        add("INFO", "Not available in this version: styles, lists, pictures, headers/footers, footnotes, comments, column insert/delete, merged-cell editing.")
        return R

    def page_info(self):
        if self.blocked:
            return None
        sc, s = self.scan(), self.view()
        out = []
        for rec in sc.sects:
            sx = s[rec["start"]:rec["end"]]
            ps, pm = get_child(sx, "pgSz"), get_child(sx, "pgMar")

            def a(x, k):
                m = re.search(r'%s="(-?\d+)"' % k, x or "")
                return int(m.group(1)) if m else None
            out.append({"body_level": rec["body_level"], "w": a(ps, "w:w") or a(ps, r"\w+:w"), "h": a(ps, r"\w+:h"),
                        "landscape": bool(ps and 'orient="landscape"' in ps), "top": a(pm, r"\w+:top"), "bottom": a(pm, r"\w+:bottom"),
                        "left": a(pm, r"\w+:left"), "right": a(pm, r"\w+:right")})
        return out

    # ---- blocks for the UI
    def blocks(self):
        if self.blocked:
            return []
        s, sc, pf = self.view(), self.scan(), self.scan().pf
        out = []
        for i, rec in enumerate(self._listed(sc)):
            pxml = s[rec["start"]:rec["end"]]
            reason = self._lock_reason(s, sc, rec)
            txt = para_text_unicode(pxml, pf) if not reason else display_text(pxml, pf)
            kind = "textbox" if rec["in_txbx"] else ("cell" if rec["container"] == "tc" else "body")
            b = {"i": i, "kind": kind, "text": txt[:4000], "long": len(txt) > 4000, "locked": bool(reason), "reason": reason}
            if kind == "cell":
                b.update({"t": rec["tbl_i"], "r": rec["row_i"], "c": rec["col_i"]})
            b.update(paragraph_props_summary(pxml, pf) if not reason else {})
            out.append(b)
        return out

    def full_texts(self):
        """complete text of every listed paragraph (for searching)"""
        if self.blocked:
            return []
        s, sc, pf = self.view(), self.scan(), self.scan().pf
        out = []
        for rec in self._listed(sc):
            pxml = s[rec["start"]:rec["end"]]
            out.append(display_text(pxml, pf) if self._lock_reason(s, sc, rec) else para_text_unicode(pxml, pf))
        return out

    def table_rows(self):
        return [len(t["rows"]) for t in self.scan().tables]

    # ---- operations
    def apply(self, ops):
        """Apply a list of ops atomically (all or nothing). Returns dict(ok, results, structure_changed)."""
        if self.blocked or self.report.get("blocked"):
            raise EditError(self.report.get("blocked") or self.blocked)
        before_cur, before_ops = self.cur, list(self.ops)
        results, structure = [], False
        try:
            for op in ops:
                self.cur = self._apply_one(op)
                self.ops.append(op)
                if op["op"] in ("insert_paragraph_after", "delete_paragraph", "insert_row", "delete_row"):
                    structure = True
                results.append({"ok": True, "op": op["op"]})
                err = _wellformed(self.cur)
                if err:
                    raise EditError("internal check failed: the result is not well-formed XML (%s)" % err)
        except EditError:
            self.cur, self.ops = before_cur, before_ops
            self._scan = None
            raise
        self.undone = []
        self._scan = None
        return {"ok": True, "results": results, "structure_changed": structure}

    def undo(self):
        if not self.ops:
            raise EditError("nothing to undo")
        ops = self.ops[:-1]
        self.cur, self.ops, self._scan = self.doc0, [], None
        for op in ops:
            self.cur = self._apply_one(op)
            self.ops.append(op)
        self._scan = None

    def _apply_one(self, op):
        s = self.view()
        sc = Scan(self.cur)
        pf = sc.pf
        kind = op.get("op")
        if kind in ("set_text", "format", "para"):
            rec = self._rec(sc, op.get("p"))
            reason = self._lock_reason(s, sc, rec)
            if reason:
                raise EditError("paragraph %d is locked: it contains %s" % (op["p"], reason))
            pxml = _open_selfclosed_p(s[rec["start"]:rec["end"]], pf)
            if kind == "set_text":
                new = str(op.get("text", "")).replace("\r\n", "\n").replace("\r", "\n")
                check_text(new)
                out = edit_runs_text(pxml, pf, new)
                if para_text_unicode(out, pf) != new:
                    raise EditError("internal check failed: the text would not match what you typed; nothing was changed")
            elif kind == "format":
                props = {k: op.get(k) for k in ("font", "size", "bold", "italic", "underline", "strike", "color", "highlight")}
                before = para_text_unicode(pxml, pf)
                out = format_paragraph_range(pxml, pf, op.get("start"), op.get("end"), props)
                if para_text_unicode(out, pf) != before:
                    raise EditError("internal check failed: formatting would change the text; nothing was changed")
            else:
                props = {k: op.get(k) for k in ("align", "rtl", "before", "after", "line", "left", "right", "first")}
                before = para_text_unicode(pxml, pf)
                out = apply_para_props(pxml, pf, props)
                if para_text_unicode(out, pf) != before:
                    raise EditError("internal check failed: paragraph settings would change the text; nothing was changed")
            return (s[:rec["start"]] + out + s[rec["end"]:]).encode("latin-1")
        if kind == "insert_paragraph_after":
            rec = self._rec(sc, op.get("p"))
            if rec["in_txbx"]:
                raise EditError("cannot add a paragraph inside a text box / shape")
            pxml = s[rec["start"]:rec["end"]]
            info = parse_paragraph(_open_selfclosed_p(pxml, pf), pf)
            ppr = info["ppr"] if info and not info.get("selfclosed") else None
            if ppr:
                for bad in ("sectPr", "pPrChange"):
                    ppr = set_child(ppr, bad, None, ORDER_PPR, pf + "pPr")
            newp = "<%sp>%s</%sp>" % (pf, ppr or "", pf)
            return (s[:rec["end"]] + newp + s[rec["end"]:]).encode("latin-1")
        if kind == "delete_paragraph":
            rec = self._rec(sc, op.get("p"))
            reason = self._lock_reason(s, sc, rec)
            if reason:
                raise EditError("paragraph %d is locked (%s); it cannot be deleted" % (op["p"], reason))
            pxml = s[rec["start"]:rec["end"]]
            if re.search(r"<%ssectPr[\s>/]" % re.escape(pf), pxml):
                raise EditError("this paragraph carries a section break; deleting it would change the page layout")
            if rec["container"] == "tc":
                sibs = rec["tc"]["children"]
            else:
                sibs = sc.bodies[0]["children"] if sc.bodies else []
            pos = next((k for k, c in enumerate(sibs) if c["start"] == rec["start"]), None)
            if pos is None:
                raise EditError("cannot locate the paragraph")
            if rec["container"] == "tc" and len([c for c in sibs if c["loc"] == "p"]) <= 1:
                raise EditError("a table cell must keep at least one paragraph")
            prev_is_tbl = pos > 0 and sibs[pos - 1]["loc"] == "tbl"
            nxt = sibs[pos + 1]["loc"] if pos + 1 < len(sibs) else None
            if prev_is_tbl and nxt in (None, "sectPr"):
                raise EditError("a table must be followed by a paragraph; this is the last one")
            return (s[:rec["start"]] + s[rec["end"]:]).encode("latin-1")
        if kind in ("insert_row", "delete_row"):
            t, r = op.get("t"), op.get("r")
            if not isinstance(t, int) or not 0 <= t < len(sc.tables) or not isinstance(r, int) or not 0 <= r < len(sc.tables[t]["rows"]):
                raise EditError("table row does not exist")
            row = sc.tables[t]["rows"][r]
            rx = s[row["start"]:row["end"]]
            bad = re.search(r"<%s(vMerge|hMerge|cellIns|cellDel|cellMerge|trPrChange|ins|del|tbl|sdt|fldChar|drawing|pict|object)[\s>/]" % re.escape(pf), rx)
            if bad:
                raise EditError("this row cannot be %s safely: it contains %s" % ("copied" if kind == "insert_row" else "deleted", FRIENDLY.get(bad.group(1), bad.group(1))))
            if kind == "delete_row":
                if len(sc.tables[t]["rows"]) <= 1:
                    raise EditError("a table needs at least one row")
                return (s[:row["start"]] + s[row["end"]:]).encode("latin-1")
            newrow = _blank_row(rx, pf)
            at = row["end"] if op.get("where", "below") == "below" else row["start"]
            return (s[:at] + newrow + s[at:]).encode("latin-1")
        if kind == "page":
            return self._page(s, sc, op)
        raise EditError("unknown operation %r" % kind)

    def _page(self, s, sc, op):
        pf = sc.pf
        try:
            w, h = int(round(float(op["w_cm"]) * TWIPS_PER_CM)), int(round(float(op["h_cm"]) * TWIPS_PER_CM))
            mg = {k: int(round(float(op[k]) * TWIPS_PER_CM)) for k in ("top", "bottom", "left", "right")}
        except (KeyError, TypeError, ValueError):
            raise EditError("page setup needs numbers for width, height and the four margins")
        if not (1134 <= w <= 31184 and 1134 <= h <= 31184):
            raise EditError("page size must be between 2 cm and 55 cm")
        if min(mg.values()) < 0 or mg["left"] + mg["right"] > w - 567 or mg["top"] + mg["bottom"] > h - 567:
            raise EditError("the margins leave no room for text on the page")
        targets = [r for r in sc.sects if op.get("scope", "last") == "all" or r["body_level"]]
        if not targets:
            raise EditError("no section found")
        for rec in sorted(targets, key=lambda r: -r["start"]):
            sx = s[rec["start"]:rec["end"]]
            ps = get_child(sx, "pgSz")
            orient = "landscape" if w > h else None
            sx = set_child(sx, "pgSz", merge_attrs(ps, pf + "pgSz", {"w": w, "h": h, "orient": orient}), ORDER_SECTPR, pf + "sectPr")
            pm = get_child(sx, "pgMar")
            if pm is None:
                pm = '<%spgMar %sheader="708" %sfooter="708" %sgutter="0"/>' % (pf, pf, pf, pf)
            sx = set_child(sx, "pgMar", merge_attrs(pm, pf + "pgMar", mg), ORDER_SECTPR, pf + "sectPr")
            s = s[:rec["start"]] + sx + s[rec["end"]:]
        return s.encode("latin-1")

    # ---- export & verification
    def export(self):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zo:
            for info in self.zin.infolist():
                data = self.cur if info.filename == MAIN_PART else self.zin.read(info.filename)
                zi = zipfile.ZipInfo(info.filename, info.date_time)
                zi.compress_type, zi.external_attr, zi.create_system = info.compress_type, info.external_attr, info.create_system
                zo.writestr(zi, data)
        return buf.getvalue()

    def verify(self):
        checks, changes = [], []

        def chk(name, ok, detail=""):
            checks.append({"name": name, "ok": bool(ok), "detail": detail})
        out = self.export()
        zo = zipfile.ZipFile(io.BytesIO(out))
        n0, n1 = self.zin.namelist(), zo.namelist()
        chk("Same list and order of parts in the package", n0 == n1, "%d parts" % len(n0))
        diff = [n for n in n0 if n in n1 and self.zin.read(n) != zo.read(n)]
        same = len(n0) - len(diff)
        chk("All other parts are byte-identical to the original", diff in ([], [MAIN_PART]), "%d of %d parts untouched%s" % (same, len(n0), "" if not diff else "; changed: " + ", ".join(diff)))
        err = _wellformed(self.cur)
        chk("Edited text part is well-formed XML", err is None, err or "")
        if err:                                   # nothing else can be compared reliably
            return {"ok": False, "checks": checks, "changes": [], "ops": len(self.ops), "paragraphs_before": 0, "paragraphs_after": 0}
        # schema order of property children (only NEW violations are reported)
        def order_viol(data):
            try:
                root = ET.fromstring(data)
            except ET.ParseError:
                return -1
            bad = 0
            for el in root.iter():
                loc = el.tag.split("}")[-1]
                order = {"pPr": ORDER_PPR, "rPr": ORDER_RPR, "sectPr": ORDER_SECTPR}.get(loc)
                if order is None:
                    continue
                last = -1
                for ch in el:
                    r = order.index(ch.tag.split("}")[-1]) if ch.tag.split("}")[-1] in order else 10 ** 6
                    if r < last:
                        bad += 1
                        break
                    last = r
            return bad
        v0, v1 = order_viol(self.doc0), order_viol(self.cur)
        chk("Property elements are in the order Word expects", v1 <= v0, "%d new ordering problems" % max(0, v1 - v0))
        # text comparison
        def texts(data):
            sc = Scan(data)
            s = data.decode("latin-1")
            return [para_text_unicode(s[r["start"]:r["end"]], sc.pf) if not self._lock_reason(s, sc, r) else display_text(s[r["start"]:r["end"]], sc.pf)
                    for r in sc.paras if not r["in_fallback"]]
        a, b = texts(self.doc0), texts(self.cur)
        sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
        n_ch = 0
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                continue
            n_ch += max(i2 - i1, j2 - j1)
            if len(changes) < 60:
                changes.append({"tag": tag, "old": " / ".join(x[:80] for x in a[i1:i2][:3]), "new": " / ".join(x[:80] for x in b[j1:j2][:3])})
        text_ops = len([o for o in self.ops if o["op"] in ("set_text", "insert_paragraph_after", "delete_paragraph", "insert_row", "delete_row")])
        fmt_ops = len(self.ops) - text_ops
        chk("Text differences are only where you edited",
            True, "%d paragraph(s) differ in text; %d text/structure edit(s) and %d formatting/page edit(s) were made" % (n_ch, text_ops, fmt_ops))
        # hard rules, compared on the raw XML of every paragraph (catches ANY collateral change)
        sc0, sc1 = Scan(self.doc0), Scan(self.cur)
        s0, s1 = self.doc0.decode("latin-1"), self.cur.decode("latin-1")
        P0 = [s0[r["start"]:r["end"]] for r in sc0.paras if not r["in_fallback"]]
        P1 = [s1[r["start"]:r["end"]] for r in sc1.paras if not r["in_fallback"]]
        sm2 = difflib.SequenceMatcher(None, P0, P1, autojunk=False)
        n_xml = sum(max(i2 - i1, j2 - j1) for tag, i1, i2, j1, j2 in sm2.get_opcodes() if tag != "equal")
        cells = max([len(t["rows"][0]["cells"]) for t in sc0.tables if t["rows"]] + [1])
        allowed = 0
        for o in self.ops:
            k = o["op"]
            allowed += cells if k in ("insert_row", "delete_row") else (len(sc0.sects) if k == "page" else 1)
        chk("No paragraph changed that you did not edit", n_xml <= allowed, "%d paragraph(s) differ at XML level, at most %d expected" % (n_xml, allowed))
        locked_before = [P0[i] for i, r in enumerate([r for r in sc0.paras if not r["in_fallback"]]) if self._lock_reason(s0, sc0, r)]
        pool = {}
        for x in P1:
            pool[x] = pool.get(x, 0) + 1
        lost = 0
        for x in locked_before:
            if pool.get(x, 0) > 0:
                pool[x] -= 1
            else:
                lost += 1
        chk("Locked paragraphs (pictures, fields, text boxes, notes ...) are untouched", lost == 0, "%d locked paragraph(s) checked, %d changed" % (len(locked_before), lost))
        fb0 = [s0[r["start"]:r["end"]] for r in sc0.paras if r["in_fallback"]]
        fb1 = [s1[r["start"]:r["end"]] for r in sc1.paras if r["in_fallback"]]
        chk("Compatibility fallback copies untouched", fb0 == fb1, "%d copies" % len(fb0))
        return {"ok": all(c["ok"] for c in checks), "checks": checks, "changes": changes, "ops": len(self.ops),
                "paragraphs_before": len(P0), "paragraphs_after": len(P1)}


def _blank_row(rx, pf):
    """copy of a table row with every cell emptied (cell properties and the first paragraph's properties are kept)"""
    e = re.escape(pf)

    def blank_cell(m):
        inner = m.group(2)
        tcpr = re.match(r"\s*(<%stcPr(?=[\s>/])(?:/>|>.*?</%stcPr>))" % (e, e), inner, re.S)
        pm = re.search(r"<%sp(?=[\s>/])[^>]*?>\s*(<%spPr(?=[\s>/])(?:/>|>.*?</%spPr>))" % (e, e, e), inner[tcpr.end():] if tcpr else inner, re.S)
        ppr = pm.group(1) if pm else ""
        for bad in ("sectPr", "pPrChange"):
            if ppr:
                ppr = set_child(ppr, bad, None, ORDER_PPR, pf + "pPr")
        para = "<%sp>%s</%sp>" % (pf, ppr or "", pf)
        return "<%stc%s>%s%s</%stc>" % (pf, m.group(1), tcpr.group(1) if tcpr else "", para, pf)
    return re.sub(r"<%stc(?=[\s>/])([^>]*)>(.*?)</%stc>" % (e, e), blank_cell, rx, flags=re.S)
