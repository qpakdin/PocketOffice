# -*- coding: utf-8 -*-
"""Illustrations group: pictures, shapes, charts inventory; resize, alt text, wrapping, removal."""
from collections import Counter

from ..core.ooxml import NS, get, qn
from ..core.units import emu, fmt_length, from_emu, parse_length
from .base import F, Feature, ProfileError

WRAPS = ["inline", "square", "tight", "through", "topAndBottom", "none"]


def _kind(d):
    g = d.find(".//a:graphicData", NS)
    uri = g.get("uri", "") if g is not None else ""
    if uri.endswith("/picture"):
        return "picture"
    if uri.endswith("/chart"):
        return "chart"
    if uri.endswith("/diagram"):
        return "smartart"
    if "wordprocessingShape" in uri or "wpg" in uri:
        return "shape"
    return "other"


class IllustrationsFeature(Feature):
    key = "illustrations"
    label = "Illustrations"
    description = "Pictures/shapes/charts: resize, max width, alt text, alignment of picture paragraphs, remove."
    fields = [
        F("max_width", "length", "Shrink pictures wider than", unit="cm"),
        F("width", "length", "Set picture width (keeps aspect ratio)", unit="cm"),
        F("height", "length", "Set picture height (keeps aspect ratio unless width also set)", unit="cm"),
        F("scale_percent", "float", "Scale pictures by %", minimum=1, maximum=1000),
        F("fit_text_width", "bool", "Fit oversized pictures to text width"),
        F("alt_text", "text", "Alt text for pictures lacking one"),
        F("overwrite_alt_text", "bool", "Replace existing alt text too"),
        F("clear_alt_text", "bool", "Clear all alt text"),
        F("alignment", "choice", "Align paragraphs that hold inline pictures", choices=["left", "center", "right"]),
        F("remove", "choice", "Remove all", choices=["none", "pictures", "shapes", "charts", "all"]),
    ]

    def apply(self, ctx, s):
        from ..core.scope import Scope
        scope = self.scope_for(ctx, s)
        n = Counter()
        sec = ctx.pkg.sections()
        for name, _k in ctx.pkg.story_parts(scope.parts):
            root = ctx.pkg.xml(name, edit=True)
            max_w = self._text_width(sec[-1]) if sec else None
            for d in list(root.iter(qn("w:drawing"))):
                kind = _kind(d)
                host = d[0] if len(d) else None
                if host is None:
                    continue
                rm = s.get("remove", "none")
                if rm == "all" or rm == kind + "s":
                    run = d.getparent()
                    run.getparent().remove(run)
                    n["removed"] += 1
                    continue
                ext = host.find("{%s}extent" % NS["wp"])
                if ext is not None and kind in ("picture", "shape", "chart"):
                    cx, cy = int(ext.get("cx")), int(ext.get("cy"))
                    ncx, ncy = cx, cy
                    if "scale_percent" in s:
                        ncx, ncy = cx * s["scale_percent"] / 100.0, cy * s["scale_percent"] / 100.0
                    if "width" in s or "height" in s:
                        if "width" in s and "height" in s:
                            ncx, ncy = emu(s["width"]), emu(s["height"])
                        elif "width" in s:
                            ncx = emu(s["width"]); ncy = cy * ncx / cx
                        else:
                            ncy = emu(s["height"]); ncx = cx * ncy / cy
                    limit = emu(s["max_width"]) if "max_width" in s else (max_w * 635 if s.get("fit_text_width") and max_w else None)
                    if limit and ncx > limit:
                        ncy = ncy * limit / ncx
                        ncx = limit
                    if (int(ncx), int(ncy)) != (cx, cy):
                        ext.set("cx", str(int(ncx)))
                        ext.set("cy", str(int(ncy)))
                        for e in d.iter("{%s}ext" % NS["a"]):
                            if e.get("cx") is not None:
                                e.set("cx", str(int(ncx)))
                                e.set("cy", str(int(ncy)))
                        n["resized"] += 1
                dp = host.find("{%s}docPr" % NS["wp"])
                if dp is not None:
                    if s.get("clear_alt_text"):
                        dp.set("descr", "")
                        n["alt_cleared"] += 1
                    elif "alt_text" in s and (s.get("overwrite_alt_text") or not dp.get("descr")):
                        dp.set("descr", s["alt_text"])
                        n["alt_set"] += 1
                if "alignment" in s and host.tag == "{%s}inline" % NS["wp"]:
                    p = d.getparent().getparent()
                    if p.tag == qn("w:p"):
                        from ..core.ooxml import ensure_ppr, ensure
                        from .paragraph import jc_value
                        ppr = ensure_ppr(p)
                        rtl = ctx.styles.is_rtl(p)
                        ensure(ppr, "w:jc", {"w:val": jc_value(s["alignment"], rtl)})
                        n["aligned"] += 1
        for k, v in n.items():
            ctx.count("illustrations." + k, v)

    @staticmethod
    def _text_width(sec):
        pgsz, mar = sec.find(qn("w:pgSz")), sec.find(qn("w:pgMar"))
        if pgsz is None or mar is None:
            return None
        w = int(get(pgsz, "w:w", 0)) - int(get(mar, "w:left", 0)) - int(get(mar, "w:right", 0))
        return w if w > 0 else None   # twips; 1 twip = 635 EMU

    def read(self, pkg, include_text=False):
        items, kinds = [], Counter()
        for name, _k in pkg.story_parts(("body", "headers", "footers")):
            for d in pkg.xml(name).iter(qn("w:drawing")):
                host = d[0] if len(d) else None
                if host is None:
                    continue
                kind = _kind(d)
                kinds[kind] += 1
                ext = host.find("{%s}extent" % NS["wp"])
                dp = host.find("{%s}docPr" % NS["wp"])
                items.append({"part": name, "kind": kind,
                              "layout": "inline" if host.tag.endswith("}inline") else "floating",
                              "width": fmt_length(from_emu(int(ext.get("cx"))), "cm") if ext is not None else None,
                              "height": fmt_length(from_emu(int(ext.get("cy"))), "cm") if ext is not None else None,
                              "has_alt_text": bool(dp is not None and dp.get("descr"))})
        media = [n for n in pkg.names() if n.startswith("word/media/")]
        return {"counts": dict(kinds), "media_files": len(media), "items": items[:200]}
