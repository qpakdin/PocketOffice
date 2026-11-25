# -*- coding: utf-8 -*-
"""Page Setup group (margins, size, orientation, columns, borders, numbering) and Breaks."""
from ..core.ooxml import NS, ensure, get, make, qn, remove, set_attrs, set_onoff, is_on, insert_ordered
from ..core.scope import iter_paragraphs, selected_sections, in_cover
from ..core.textops import has_visible_content
from ..core.units import PAPER_TWIPS, fmt_length, from_twips, twips
from .base import F, Feature, ProfileError

BORDER_STYLES = ["single", "double", "dotted", "dashed", "dotDash", "thick", "wave", "triple", "thinThickSmallGap", "none"]
PAGE_NUM_FORMATS = ["decimal", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter", "hindiNumbers",
                    "arabicAbjad", "arabicAlpha", "decimalZero", "ordinal"]
SIDES = ("top", "left", "bottom", "right")


def paper_name(w, h):
    w, h = sorted((int(w), int(h)))
    for name, (pw, ph) in PAPER_TWIPS.items():
        if abs(pw - w) <= 40 and abs(ph - h) <= 40:
            return name
    return None


class PageFeature(Feature):
    key = "page"
    label = "Page Setup"
    description = "Paper size, orientation, margins, columns, direction, page borders, page-number format."
    fields = [
        F("paper", "choice", "Paper size", choices=list(PAPER_TWIPS)),
        F("width", "length", "Custom page width", unit="cm"),
        F("height", "length", "Custom page height", unit="cm"),
        F("orientation", "choice", "Orientation", choices=["portrait", "landscape"]),
        F("margin_top", "length", "Top margin", unit="cm", hint="2.5cm / 1in / 72pt (bare number = cm)"),
        F("margin_bottom", "length", "Bottom margin", unit="cm"),
        F("margin_left", "length", "Left margin", unit="cm"),
        F("margin_right", "length", "Right margin", unit="cm"),
        F("gutter", "length", "Gutter", unit="cm"),
        F("gutter_position", "choice", "Gutter position (printing)", choices=["left", "top", "right"],
          hint="top is document-wide; right sets the section's rtlGutter"),
        F("header_distance", "length", "Header distance from edge", unit="cm"),
        F("footer_distance", "length", "Footer distance from edge", unit="cm"),
        F("mirror_margins", "bool", "Mirror margins (document-wide)"),
        F("columns", "int", "Number of columns", minimum=1, maximum=12),
        F("column_spacing", "length", "Column spacing", unit="cm"),
        F("column_separator", "bool", "Line between columns"),
        F("vertical_align", "choice", "Vertical alignment of page", choices=["top", "center", "both", "bottom"]),
        F("section_direction", "choice", "Section direction (bidi)", choices=["ltr", "rtl"]),
        F("page_number_format", "choice", "Page number format", choices=PAGE_NUM_FORMATS),
        F("page_number_start", "int", "Start page numbering at", minimum=0, maximum=99999),
        F("line_numbers", "choice", "Line numbering", choices=["none", "continuous", "restart_page", "restart_section"]),
        F("borders", "json", "Page borders",
          hint='{"style":"single","width":"1pt","color":"000000","space":"24pt","sides":"box|top,bottom","offset_from":"page|text","display":"all|first|notfirst"}'),
        F("remove_borders", "bool", "Remove page borders"),
        F("sections", "json", "Apply to sections", hint='"all" or [1,3] (1-based). Default: all'),
    ]

    def validate(self, s):
        if "borders" in s:
            self._border(s["borders"])
        return s

    @staticmethod
    def _border(b):
        from ..core.units import normalize_color, parse_length
        if not isinstance(b, dict):
            raise ProfileError("page.borders must be an object")
        style = b.get("style", "single")
        if style not in BORDER_STYLES:
            raise ProfileError("page.borders.style must be one of %s" % ", ".join(BORDER_STYLES))
        sides = b.get("sides", "box")
        if sides == "box":
            sides = list(SIDES)
        elif isinstance(sides, str):
            sides = [x.strip() for x in sides.split(",") if x.strip()]
        if any(x not in SIDES for x in sides):
            raise ProfileError("page.borders.sides: use top,bottom,left,right or box")
        try:
            return {"style": style, "sides": sides,
                    "sz": max(2, min(96, int(round(parse_length(b.get("width", "1pt")) * 8)))),
                    "space": max(0, min(31, int(round(parse_length(b.get("space", "24pt")))))),
                    "color": normalize_color(b.get("color", "auto")),
                    "offset": "text" if b.get("offset_from", "page") == "text" else "page",
                    "display": {"all": "allPages", "first": "firstPage", "notfirst": "notFirstPage"}.get(
                        str(b.get("display", "all")).lower().replace("_", ""), "allPages")}
        except ValueError as e:
            raise ProfileError("page.borders: %s" % e)

    def apply(self, ctx, s):
        scope = self.scope_for(ctx, s)
        if "sections" in s:
            from ..core.scope import Scope
            scope = Scope(scope.parts, scope.styles, scope.include_cover, s["sections"])
        secs = selected_sections(ctx.pkg, scope, edit=True)
        for sec in secs:
            self._section(sec, s)
        if "mirror_margins" in s:
            st = ctx.pkg.xml("word/settings.xml", edit=True) if ctx.pkg.exists("word/settings.xml") else None
            if st is not None:
                set_onoff(st, "w:mirrorMargins", s["mirror_margins"] or None)
        if "gutter_position" in s:
            st = ctx.pkg.xml("word/settings.xml", edit=True) if ctx.pkg.exists("word/settings.xml") else None
            if st is not None:
                set_onoff(st, "w:gutterAtTop", True if s["gutter_position"] == "top" else None)
            for sec in secs:
                set_onoff(sec, "w:rtlGutter", True if s["gutter_position"] == "right" else None)
        ctx.count("page.sections", len(secs))

    def _section(self, sec, s):
        pgsz = sec.find(qn("w:pgSz"))
        if pgsz is None:
            pgsz = insert_ordered(sec, make("w:pgSz", {"w:w": 11906, "w:h": 16838}))
        w, h = int(get(pgsz, "w:w", 11906)), int(get(pgsz, "w:h", 16838))
        orient = get(pgsz, "w:orient") or ("landscape" if w > h else "portrait")
        if "paper" in s:
            w, h = PAPER_TWIPS[s["paper"]]
            if orient == "landscape":
                w, h = h, w
        if "width" in s:
            w = twips(s["width"])
        if "height" in s:
            h = twips(s["height"])
        if "orientation" in s:
            want = s["orientation"]
            if (want == "landscape") != (w > h):
                w, h = h, w
            orient = want
        set_attrs(pgsz, {"w:w": w, "w:h": h, "w:orient": "landscape" if w > h else None})
        if any(k in s for k in ("margin_top", "margin_bottom", "margin_left", "margin_right", "gutter",
                                "header_distance", "footer_distance")):
            mar = sec.find(qn("w:pgMar"))
            if mar is None:
                mar = insert_ordered(sec, make("w:pgMar", {"w:top": 1440, "w:right": 1440, "w:bottom": 1440,
                                                           "w:left": 1440, "w:header": 720, "w:footer": 720, "w:gutter": 0}))
            for k, a in (("margin_top", "w:top"), ("margin_bottom", "w:bottom"), ("margin_left", "w:left"),
                         ("margin_right", "w:right"), ("gutter", "w:gutter"),
                         ("header_distance", "w:header"), ("footer_distance", "w:footer")):
                if k in s:
                    mar.set(qn(a), str(twips(s[k])))
        if "columns" in s or "column_spacing" in s or "column_separator" in s:
            cols = ensure(sec, "w:cols")
            if "columns" in s:
                for c in cols.findall(qn("w:col")):
                    cols.remove(c)
                set_attrs(cols, {"w:num": s["columns"], "w:equalWidth": None})
            if "column_spacing" in s:
                cols.set(qn("w:space"), str(twips(s["column_spacing"])))
            if "column_separator" in s:
                cols.set(qn("w:sep"), "1" if s["column_separator"] else "0")
        if "vertical_align" in s:
            ensure(sec, "w:vAlign", {"w:val": s["vertical_align"]})
        if "section_direction" in s:
            set_onoff(sec, "w:bidi", True if s["section_direction"] == "rtl" else None)
        if "page_number_format" in s or "page_number_start" in s:
            pn = ensure(sec, "w:pgNumType")
            if "page_number_format" in s:
                pn.set(qn("w:fmt"), s["page_number_format"])
            if "page_number_start" in s:
                pn.set(qn("w:start"), str(s["page_number_start"]))
        if "line_numbers" in s:
            if s["line_numbers"] == "none":
                remove(sec, "w:lnNumType")
            else:
                ensure(sec, "w:lnNumType", {"w:countBy": 1, "w:restart": {
                    "continuous": "continuous", "restart_page": "newPage", "restart_section": "newSection"}[s["line_numbers"]]})
        if s.get("remove_borders"):
            remove(sec, "w:pgBorders")
        if "borders" in s:
            b = self._border(s["borders"])
            pb = sec.find(qn("w:pgBorders"))
            if pb is not None:
                sec.remove(pb)
            pb = insert_ordered(sec, make("w:pgBorders", {"w:offsetFrom": b["offset"], "w:display": b["display"]}))
            for side in SIDES:           # schema order: top,left,bottom,right
                if side in b["sides"]:
                    pb.append(make("w:" + side, {"w:val": b["style"], "w:sz": b["sz"], "w:space": b["space"], "w:color": b["color"]}))

    # ------------------------------------------------------------------- read
    def read(self, pkg, include_text=False):
        out = {"sections": []}
        for i, sec in enumerate(pkg.sections(), 1):
            pgsz, mar, cols = sec.find(qn("w:pgSz")), sec.find(qn("w:pgMar")), sec.find(qn("w:cols"))
            w = int(get(pgsz, "w:w", 0) or 0)
            h = int(get(pgsz, "w:h", 0) or 0)

            def m(a):
                v = get(mar, a)
                return fmt_length(from_twips(int(v)), "cm") if v is not None else None
            pb = sec.find(qn("w:pgBorders"))
            pn = sec.find(qn("w:pgNumType"))
            nc = get(cols, "w:num", "1") if cols is not None else "1"
            t = sec.find(qn("w:type"))
            out["sections"].append({
                "index": i, "type": get(t, "w:val", "nextPage") if t is not None else "nextPage",
                "paper": paper_name(w, h) if w and h else None,
                "width": fmt_length(from_twips(w), "cm") if w else None,
                "height": fmt_length(from_twips(h), "cm") if h else None,
                "orientation": get(pgsz, "w:orient") or ("landscape" if w > h else "portrait"),
                "margin_top": m("w:top"), "margin_bottom": m("w:bottom"), "margin_left": m("w:left"),
                "margin_right": m("w:right"), "gutter": m("w:gutter"), "header_distance": m("w:header"),
                "footer_distance": m("w:footer"),
                "columns": int(nc), "direction": "rtl" if is_on(sec.find(qn("w:bidi"))) else "ltr",
                "vertical_align": get(sec.find(qn("w:vAlign")), "w:val", "top"),
                "page_borders": [c.tag.split("}")[1] for c in pb] if pb is not None else [],
                "page_border_style": get(pb[0], "w:val") if pb is not None and len(pb) else None,
                "page_number_format": get(pn, "w:fmt", "decimal") if pn is not None else "decimal",
                "page_number_start": get(pn, "w:start") if pn is not None else None,
                "different_first_page": sec.find(qn("w:titlePg")) is not None,
                "line_numbers": sec.find(qn("w:lnNumType")) is not None,
            })
        st = pkg.try_xml("word/settings.xml")
        out["mirror_margins"] = st is not None and st.find(qn("w:mirrorMargins")) is not None
        out["gutter_at_top"] = st is not None and st.find(qn("w:gutterAtTop")) is not None
        for row, sec in zip(out["sections"], pkg.sections()):
            row["gutter_position"] = ("top" if out["gutter_at_top"] else
                                      "right" if is_on(sec.find(qn("w:rtlGutter"))) else "left")
        return out

    def snapshot(self, pkg):
        secs = self.read(pkg)["sections"]
        if not secs:
            return {}
        f = secs[0]
        out = {k: f[k] for k in ("paper", "orientation", "margin_top", "margin_bottom", "margin_left",
                                 "margin_right", "gutter", "header_distance", "footer_distance") if f.get(k)}
        if f["columns"] > 1:
            out["columns"] = f["columns"]
        out["section_direction"] = f["direction"]
        out["gutter_position"] = f["gutter_position"]
        return out


class BreaksFeature(Feature):
    key = "breaks"
    label = "Breaks"
    description = "Remove or convert page/column/section breaks; control section break types."
    fields = [
        F("remove_page_breaks", "bool", "Remove manual page breaks"),
        F("remove_column_breaks", "bool", "Remove manual column breaks"),
        F("remove_section_breaks", "bool", "Merge sections (remove section breaks, keeps last section's page setup)"),
        F("section_break_type", "choice", "Type of every section break", choices=["nextPage", "continuous", "evenPage", "oddPage", "nextColumn"]),
        F("page_break_before_styles", "list", "Start a new page before these styles", hint="e.g. Heading 1"),
        F("no_page_break_before_styles", "list", "Clear 'page break before' on these styles"),
    ]

    def apply(self, ctx, s):
        scope = self.scope_for(ctx, s)
        doc = ctx.pkg.document(edit=True)
        removed = 0
        if s.get("remove_page_breaks") or s.get("remove_column_breaks"):
            kinds = set()
            if s.get("remove_page_breaks"):
                kinds.add("page")
            if s.get("remove_column_breaks"):
                kinds.add("column")
            for br in list(doc.iter(qn("w:br"))):
                if get(br, "w:type") in kinds:
                    r = br.getparent()
                    r.remove(br)
                    removed += 1
                    if len([c for c in r if c.tag != qn("w:rPr")]) == 0:
                        r.getparent().remove(r)
            for lb in list(doc.iter(qn("w:lastRenderedPageBreak"))):
                lb.getparent().remove(lb)
        ctx.count("breaks.removed", removed)
        if s.get("remove_section_breaks"):
            n = 0
            for sp in list(doc.xpath("/w:document/w:body//w:pPr/w:sectPr", namespaces=NS)):
                sp.getparent().remove(sp)
                n += 1
            ctx.count("breaks.sections_merged", n)
        if "section_break_type" in s:
            n = 0
            for sec in ctx.pkg.sections(edit=True):
                ensure(sec, "w:type", {"w:val": s["section_break_type"]})
                n += 1
            ctx.count("breaks.type_set", n)
        if s.get("no_page_break_before_styles") or s.get("page_break_before_styles"):
            res = ctx.styles
            if res.root is not None:
                ctx.pkg.xml("word/styles.xml", edit=True)
                for lst, val in ((s.get("page_break_before_styles"), True), (s.get("no_page_break_before_styles"), False)):
                    for name in lst or []:
                        sid = res.find_id(name, "paragraph")
                        if sid is None:
                            ctx.warn("breaks: style %r not found" % name)
                            continue
                        ppr = ensure(res.styles[sid], "w:pPr")
                        set_onoff(ppr, "w:pageBreakBefore", True if val else None)
                        ctx.count("breaks.styles_changed")

    def read(self, pkg, include_text=False):
        doc = pkg.document()
        brs = [get(b, "w:type", "textWrapping") for b in doc.iter(qn("w:br"))]
        secs = pkg.sections()
        types = [get(sec.find(qn("w:type")), "w:val", "nextPage") if sec.find(qn("w:type")) is not None else "nextPage" for sec in secs]
        pbb = 0
        for p in doc.iter(qn("w:p")):
            ppr = p.find(qn("w:pPr"))
            if ppr is not None and is_on(ppr.find(qn("w:pageBreakBefore"))):
                pbb += 1
        return {"page_breaks": brs.count("page"), "column_breaks": brs.count("column"),
                "line_breaks": brs.count("textWrapping"), "sections": len(secs), "section_break_types": types,
                "paragraphs_with_page_break_before": pbb}
