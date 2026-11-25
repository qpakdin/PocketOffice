# -*- coding: utf-8 -*-
"""Cover page: detect / remove Word cover-page building blocks, insert a simple one."""
from ..core.ooxml import NS, get, make, qn, remove
from ..core.scope import in_cover
from .base import F, Feature, ProfileError
from .font import FontFeature


def _cover_sdts(doc):
    return [s for s in doc.findall("w:body/w:sdt", NS)
            if s.find("w:sdtPr/w:docPartObj/w:docPartGallery[@w:val='Cover Pages']", NS) is not None]


class CoverFeature(Feature):
    key = "cover"
    label = "Cover Page"
    description = "Remove the existing cover page or insert a simple title cover page (own section break)."
    uses_scope = False
    fields = [
        F("remove", "bool", "Remove existing cover page"),
        F("title", "text", "Insert a cover page with this title"),
        F("subtitle", "text", "Cover subtitle"),
        F("author", "text", "Cover author / organisation line"),
        F("date", "text", "Cover date line"),
        F("direction", "choice", "Cover direction", choices=["ltr", "rtl"]),
        F("font", "text", "Cover font"),
        F("title_size", "float", "Title size (pt)", default=32, minimum=8, maximum=200),
        F("color", "color", "Title colour"),
        F("replace_existing", "bool", "Replace existing cover when inserting", default=True),
    ]

    def apply(self, ctx, s):
        doc = ctx.pkg.document(edit=True)
        body = doc.find(qn("w:body"))
        if s.get("remove") or ("title" in s and s.get("replace_existing", True)):
            n = 0
            for sdt in _cover_sdts(doc):
                body.remove(sdt); n += 1
            ctx.count("cover.removed", n)
        if "title" in s:
            rtl = s.get("direction") == "rtl"
            base = {k: s[k] for k in ("font", "color") if k in s}
            if "font" in s:
                base["sync_scripts"] = True
            def para(text, size, bold=False, after=240):
                p = make("w:p"); ppr = make("w:pPr")
                if rtl:
                    ppr.append(make("w:bidi"))
                ppr.append(make("w:spacing", {"w:before": 0, "w:after": after}))
                ppr.append(make("w:jc", {"w:val": "center"}))
                p.append(ppr)
                r = make("w:r"); rpr = make("w:rPr")
                st = dict(base); st["size"] = size
                if bold:
                    st["bold"] = True
                FontFeature().apply_rpr(rpr, st)
                r.append(rpr)
                t = make("w:t", text=text)
                r.append(t); p.append(r)
                return p
            items = [make("w:p")]
            items[0].append(make("w:pPr"))
            items[0].find(qn("w:pPr")).append(make("w:spacing", {"w:before": 3600, "w:after": 0}))
            items.append(para(s["title"], s.get("title_size", 32), True, 360))
            for key, size in (("subtitle", 18), ("author", 14), ("date", 12)):
                if s.get(key):
                    items.append(para(s[key], size))
            # last cover paragraph carries a section break (next page) copying the first section's setup
            import copy
            first = ctx.pkg.sections()[0]
            sec = copy.deepcopy(first)
            for tag in ("w:headerReference", "w:footerReference", "w:titlePg"):
                remove(sec, tag)
            ctx.pkg.sections(edit=True)
            items[-1].find(qn("w:pPr")).append(sec)
            for i, el in enumerate(items):
                body.insert(i, el)
            ctx.count("cover.inserted")

    def read(self, pkg, include_text=False):
        doc = pkg.document()
        sd = _cover_sdts(doc)
        return {"has_cover_page_block": bool(sd)}
