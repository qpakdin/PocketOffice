# -*- coding: utf-8 -*-
"""Header & Footer group: text, page-number fields, first/even pages, removal."""
import re

from lxml import etree

from ..core.ooxml import NS, RELT, ensure, get, insert_ordered, is_on, make, qn, remove, set_attrs, set_onoff
from ..core.units import twips
from .base import F, Feature, ProfileError
from .font import FontFeature
from .paragraph import jc_value

HF_CT = {"header": "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
         "footer": "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"}
TYPES = ("default", "first", "even")
_FIELD_RE = re.compile(r"(\{PAGE\}|\{NUMPAGES\}|\{SECTIONPAGES\}|\{DATE\})")
_FIELDS = {"{PAGE}": "PAGE", "{NUMPAGES}": "NUMPAGES", "{SECTIONPAGES}": "SECTIONPAGES", "{DATE}": "DATE \\@ \"yyyy/MM/dd\""}
SPEC_KEYS = {"left", "center", "right", "text", "align", "direction", "font", "font_cs", "size", "color",
             "bold", "italic", "clear", "border_bottom", "border_top"}


def _root(kind):
    tag = "w:hdr" if kind == "header" else "w:ftr"
    return etree.Element(qn(tag), nsmap={"w": NS["w"], "r": NS["r"]})


def sec_refs(sec, kind):
    return {get(e, "w:type", "default"): get(e, "r:id") for e in sec.findall(qn("w:%sReference" % kind))}


def ensure_part(pkg, secs, idx, kind, typ):
    """Return (part name, created) for section ``idx`` honouring inheritance."""
    sec = secs[idx]
    refs = sec_refs(sec, kind)
    if typ in refs:
        part = pkg.part_for_rid(pkg.main, refs[typ])
        if part and pkg.exists(part):
            return part, False
    for j in range(idx - 1, -1, -1):       # inherited from a previous section
        prev = sec_refs(secs[j], kind)
        if typ in prev:
            part = pkg.part_for_rid(pkg.main, prev[typ])
            if part and pkg.exists(part):
                return part, False
    n = 1
    while pkg.exists("word/%s%d.xml" % (kind, n)):
        n += 1
    name = "word/%s%d.xml" % (kind, n)
    root = _root(kind)
    root.append(make("w:p"))
    pkg.add_xml_part(name, root, HF_CT[kind])
    rid = pkg.add_rel(pkg.main, RELT + kind, name)
    ref = make("w:%sReference" % kind, {"w:type": typ})
    ref.set(qn("r:id"), rid)
    insert_ordered(sec, ref)
    return name, True


def text_width_twips(sec):
    pgsz, mar = sec.find(qn("w:pgSz")), sec.find(qn("w:pgMar"))
    w = int(get(pgsz, "w:w", 11906) or 11906)
    l = int(get(mar, "w:left", 1440) or 1440)
    r = int(get(mar, "w:right", 1440) or 1440)
    g = int(get(mar, "w:gutter", 0) or 0)
    return max(w - l - r - g, 1000)


def _run(text, rpr_settings):
    r = make("w:r")
    rpr = make("w:rPr")
    FontFeature().apply_rpr(rpr, rpr_settings)
    if len(rpr):
        r.append(rpr)
    t = make("w:t", text=text)
    t.set("{%s}space" % NS["xml"], "preserve")
    r.append(t)
    return r


def _field(instr, rpr_settings):
    f = make("w:fldSimple", {"w:instr": " %s " % instr})
    f.append(_run("1", rpr_settings))
    return f


def _content(text, rpr_settings):
    """Runs/fields for a string possibly containing {PAGE} style tokens."""
    out = []
    for part in _FIELD_RE.split(text):
        if not part:
            continue
        out.append(_field(_FIELDS[part], rpr_settings) if part in _FIELDS else _run(part, rpr_settings))
    return out


def build_paragraph(spec, width):
    rtl = spec.get("direction") == "rtl"
    rpr_s = {k: spec[k] for k in ("font", "font_cs", "size", "color", "bold", "italic") if k in spec}
    if rpr_s:
        rpr_s.setdefault("sync_scripts", True)
    p = make("w:p")
    ppr = make("w:pPr")
    p.append(ppr)
    if rtl:
        ensure(ppr, "w:bidi")
    if spec.get("border_bottom") or spec.get("border_top"):
        bd = ensure(ppr, "w:pBdr")
        for side, key in (("top", "border_top"), ("bottom", "border_bottom")):
            if spec.get(key):
                bd.append(make("w:" + side, {"w:val": "single", "w:sz": 4, "w:space": 1, "w:color": "808080"}))
    positional = [k for k in ("left", "center", "right") if spec.get(k)]
    if len(positional) > 1:
        tabs = ensure(ppr, "w:tabs")
        tabs.append(make("w:tab", {"w:val": "center", "w:pos": width // 2}))
        tabs.append(make("w:tab", {"w:val": "right", "w:pos": width}))
        order = ["right", "center", "left"] if rtl else ["left", "center", "right"]
        for i, k in enumerate(order):
            if i:
                p.append(_tab_run())
            for el in _content(spec.get(k, ""), rpr_s):
                p.append(el)
    else:
        key = positional[0] if positional else "text"
        align = spec.get("align") or {"left": "left", "center": "center", "right": "right"}.get(key, "center")
        ensure(ppr, "w:jc", {"w:val": jc_value(align, rtl)})
        for el in _content(spec.get(key, ""), rpr_s):
            p.append(el)
    return p


def _tab_run():
    r = make("w:r")
    r.append(make("w:tab"))
    return r


def normalize_spec(v, name):
    if isinstance(v, str):
        v = {"text": v}
    if not isinstance(v, dict):
        raise ProfileError("header_footer.%s must be text or an object" % name)
    bad = set(v) - SPEC_KEYS
    if bad:
        raise ProfileError("header_footer.%s: unknown keys %s (allowed: %s)" % (name, sorted(bad), ", ".join(sorted(SPEC_KEYS))))
    return v


class HeaderFooterFeature(Feature):
    key = "header_footer"
    label = "Header & Footer"
    description = "Set header/footer text (left/center/right, {PAGE}/{NUMPAGES} fields), first/even pages, removal."
    uses_scope = False
    _SPEC = ("header", "footer", "first_header", "first_footer", "even_header", "even_footer")
    fields = [
        F("header", "json", "Header (default)", hint='"text" or {"left":"","center":"Title","right":"{PAGE}","font":"Arial","size":9}'),
        F("footer", "json", "Footer (default)", hint='{"center":"Page {PAGE} of {NUMPAGES}"}'),
        F("first_header", "json", "Header (first page)"),
        F("first_footer", "json", "Footer (first page)"),
        F("even_header", "json", "Header (even pages)"),
        F("even_footer", "json", "Footer (even pages)"),
        F("different_first_page", "bool", "Different first page"),
        F("different_odd_even", "bool", "Different odd & even pages"),
        F("remove_headers", "bool", "Remove all headers"),
        F("remove_footers", "bool", "Remove all footers"),
        F("clear_headers", "bool", "Empty all headers (keep the parts)"),
        F("clear_footers", "bool", "Empty all footers (keep the parts)"),
    ]

    def validate(self, s):
        for k in self._SPEC:
            if k in s:
                s[k] = normalize_spec(s[k], k)
        return s

    def apply(self, ctx, s):
        pkg = ctx.pkg
        secs = pkg.sections(edit=True)
        if "different_first_page" in s:
            for sec in secs:
                set_onoff(sec, "w:titlePg", True if s["different_first_page"] else None)
        if "different_odd_even" in s or "even_header" in s or "even_footer" in s:
            st = pkg.try_xml("word/settings.xml", edit=True)
            if st is not None:
                want = s.get("different_odd_even", True)
                set_onoff(st, "w:evenAndOddHeaders", True if want else None)
            else:
                ctx.warn("header_footer: no settings.xml; even/odd flag skipped")
        for kind, rm, cl in (("header", "remove_headers", "clear_headers"), ("footer", "remove_footers", "clear_footers")):
            if s.get(rm):
                for _rid, part, _t, ext in pkg.rels(pkg.main, kind):
                    if not ext:
                        pkg.remove_part(part)
                n = 0
                for sec in secs:
                    for e in sec.findall(qn("w:%sReference" % kind)):
                        sec.remove(e)
                        n += 1
                ctx.count("header_footer.%ss_removed" % kind, n)
            elif s.get(cl):
                for _rid, part, _t, ext in pkg.rels(pkg.main, kind):
                    if not ext and pkg.exists(part):
                        root = pkg.xml(part, edit=True)
                        for c in list(root):
                            root.remove(c)
                        root.append(make("w:p"))
        for key in self._SPEC:
            if key not in s:
                continue
            typ, kind = ("default", key) if key in ("header", "footer") else (key.split("_")[0], key.split("_")[1])
            if (kind == "header" and s.get("remove_headers")) or (kind == "footer" and s.get("remove_footers")):
                continue
            if typ == "first":
                for sec in secs:
                    set_onoff(sec, "w:titlePg", True)
            done = set()
            for i in range(len(secs)):
                part, _created = ensure_part(pkg, secs, i, kind, typ)
                if part in done:
                    continue
                done.add(part)
                root = pkg.xml(part, edit=True)
                spec = s[key]
                if spec.get("clear", True):
                    for c in list(root):
                        root.remove(c)
                root.append(build_paragraph(spec, text_width_twips(secs[i])))
                ctx.count("header_footer.parts_written")

    # ------------------------------------------------------------------- read
    def read(self, pkg, include_text=False):
        secs = pkg.sections()
        st = pkg.try_xml("word/settings.xml")
        out = {"different_odd_even": st is not None and st.find(qn("w:evenAndOddHeaders")) is not None, "sections": []}
        for i, sec in enumerate(secs, 1):
            row = {"index": i, "different_first_page": sec.find(qn("w:titlePg")) is not None}
            for kind in ("header", "footer"):
                refs = sec_refs(sec, kind)
                row[kind] = {}
                for typ, rid in refs.items():
                    part = pkg.part_for_rid(pkg.main, rid)
                    if not part or not pkg.exists(part):
                        continue
                    root = pkg.xml(part)
                    instr = " ".join((x.text or "") for x in root.iter(qn("w:instrText")))
                    instr += " ".join(get(x, "w:instr", "") for x in root.iter(qn("w:fldSimple")))
                    info = {"part": part,
                            "has_page_number": "PAGE" in instr, "has_total_pages": "NUMPAGES" in instr,
                            "has_watermark": any("PowerPlusWaterMarkObject" in (get(e, "id", "") or e.get("id", "")) for e in root.iter("{%s}shape" % NS["v"])),
                            "has_picture": root.find(".//" + qn("w:drawing")) is not None}
                    txt = "".join(t.text or "" for t in root.iter(qn("w:t"))).strip()
                    info["empty"] = not txt and not info["has_page_number"]
                    if include_text:
                        info["text"] = txt
                    row[kind][typ] = info
                if not refs:
                    row[kind] = "inherited" if i > 1 else "none"
            out["sections"].append(row)
        return out
