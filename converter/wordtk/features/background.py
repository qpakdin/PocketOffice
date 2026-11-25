# -*- coding: utf-8 -*-
"""Page Background group: page colour, text watermark, remove watermark."""
import re

from lxml import etree

from ..core.ooxml import NS, ensure, get, make, qn, remove, set_attrs, set_onoff
from ..core.units import normalize_color
from .base import F, Feature, ProfileError
from .header_footer import ensure_part, sec_refs

WM_ID = "PowerPlusWaterMarkObject"


def _watermark_pict(text, color, font, size, opacity, rotate, shape_id):
    """VML text watermark exactly as Word writes it (WordArt in the header)."""
    v, o, w10 = NS["v"], NS["o"], NS["w10"]
    xml = (
        '<w:r xmlns:w="%(w)s" xmlns:v="%(v)s" xmlns:o="%(o)s" xmlns:w10="%(w10)s"><w:rPr><w:noProof/></w:rPr>'
        '<w:pict>'
        '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">'
        '<v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/>'
        '<v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/>'
        '<v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/>'
        '<v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas>'
        '<v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>'
        '<v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>'
        '<o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>'
        '<v:shape id="%(id)s" o:spid="_x0000_s2049" type="#_x0000_t136" '
        'style="position:absolute;margin-left:0;margin-top:0;width:%(wd)spt;height:%(ht)spt;rotation:%(rot)s;'
        'z-index:-251658752;mso-position-horizontal:center;mso-position-horizontal-relative:margin;'
        'mso-position-vertical:center;mso-position-vertical-relative:margin" '
        'o:allowincell="f" fillcolor="#%(color)s" stroked="f">'
        '<v:fill opacity="%(op)s"/><v:textpath style="font-family:&quot;%(font)s&quot;;font-size:1pt" string="%(text)s"/>'
        '<w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r>'
    ) % {"w": NS["w"], "v": v, "o": o, "w10": w10, "id": shape_id, "color": color,
         "font": font, "text": text.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;"),
         "wd": int(max(len(text), 2) * size * 0.55), "ht": int(size), "rot": rotate, "op": opacity}
    return etree.fromstring(xml)


class BackgroundFeature(Feature):
    key = "background"
    label = "Page Background"
    description = "Page colour, text watermark (header based), remove watermark."
    uses_scope = False
    fields = [
        F("page_color", "color", "Page colour", hint="hex, or 'auto' to remove"),
        F("watermark_text", "text", "Watermark text"),
        F("watermark_color", "color", "Watermark colour", default="C0C0C0"),
        F("watermark_font", "text", "Watermark font", default="Calibri"),
        F("watermark_size", "float", "Watermark size (pt)", default=80, minimum=10, maximum=500),
        F("watermark_opacity", "float", "Watermark opacity (0-1)", default=0.5, minimum=0.05, maximum=1),
        F("watermark_angle", "float", "Watermark rotation (degrees)", default=315, minimum=0, maximum=360),
        F("remove_watermark", "bool", "Remove watermark"),
    ]

    def apply(self, ctx, s):
        pkg = ctx.pkg
        if "page_color" in s:
            doc = pkg.document(edit=True)
            bg = doc.find(qn("w:background"))
            if s["page_color"] == "auto":
                if bg is not None:
                    doc.remove(bg)
                self._display_bg(pkg, False)
            else:
                if bg is None:
                    bg = make("w:background")
                    doc.insert(0, bg)
                set_attrs(bg, {"w:color": s["page_color"], "w:themeColor": None, "w:themeTint": None, "w:themeShade": None})
                self._display_bg(pkg, True)
            ctx.count("background.page_color")
        if s.get("remove_watermark") or "watermark_text" in s:
            n = self._remove_wm(pkg)
            ctx.count("background.watermarks_removed", n)
        if "watermark_text" in s:
            secs = pkg.sections(edit=True)
            done = set()
            for i in range(len(secs)):
                part, _ = ensure_part(pkg, secs, i, "header", "default")
                if part in done:
                    continue
                done.add(part)
                root = pkg.xml(part, edit=True)
                p = root.find(qn("w:p"))
                if p is None:
                    p = make("w:p")
                    root.append(p)
                col = s.get("watermark_color", "C0C0C0")
                p.insert(1 if p.find(qn("w:pPr")) is not None else 0, _watermark_pict(
                    s["watermark_text"], col if col != "auto" else "C0C0C0", s.get("watermark_font", "Calibri"),
                    s.get("watermark_size", 80), s.get("watermark_opacity", 0.5), int(s.get("watermark_angle", 315)),
                    "%s%d" % (WM_ID, len(done))))
            ctx.count("background.watermarks_added", len(done))

    @staticmethod
    def _display_bg(pkg, on):
        st = pkg.try_xml("word/settings.xml", edit=True)
        if st is not None:
            set_onoff(st, "w:displayBackgroundShape", True if on else None)

    @staticmethod
    def _remove_wm(pkg):
        n = 0
        for name, kind in pkg.story_parts(("headers",)):
            root = pkg.xml(name)
            for pict in list(root.iter(qn("w:pict"))):
                shapes = [x for x in pict if isinstance(x.tag, str) and x.tag == "{%s}shape" % NS["v"] and WM_ID in x.get("id", "")]
                if shapes:
                    run = pict.getparent()
                    pkg.xml(name, edit=True)
                    run.getparent().remove(run)
                    n += 1
        return n

    def read(self, pkg, include_text=False):
        doc = pkg.document()
        bg = doc.find(qn("w:background"))
        wm = []
        for name, _k in pkg.story_parts(("headers",)):
            for sh in pkg.xml(name).iter("{%s}shape" % NS["v"]):
                if WM_ID in sh.get("id", ""):
                    tp = sh.find("{%s}textpath" % NS["v"])
                    wm.append({"part": name, "text": tp.get("string") if (tp is not None and include_text) else "(hidden)",
                               "color": sh.get("fillcolor")})
        return {"page_color": get(bg, "w:color") if bg is not None else None, "watermarks": wm}

    def snapshot(self, pkg):
        c = self.read(pkg)["page_color"]
        return {"page_color": c} if c and c != "auto" else {}
