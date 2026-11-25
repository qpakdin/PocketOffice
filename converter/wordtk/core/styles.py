# -*- coding: utf-8 -*-
"""Style resolution (basedOn chain + docDefaults) for *effective* property lookup."""
from .ooxml import NS, qn, get, is_on


class StyleResolver:
    def __init__(self, pkg):
        self.pkg = pkg
        self.styles = {}        # styleId -> element
        self.by_name = {}       # lower-case name -> styleId
        self.default_p = None
        self.default_c = None
        self.doc_ppr = None
        self.doc_rpr = None
        self.root = pkg.try_xml("word/styles.xml")
        if self.root is None:
            return
        for s in self.root.findall(qn("w:style")):
            sid = s.get(qn("w:styleId"))
            if not sid:
                continue
            self.styles[sid] = s
            n = s.find(qn("w:name"))
            if n is not None:
                self.by_name[get(n, "w:val", "").lower()] = sid
            if s.get(qn("w:default")) in ("1", "true"):
                if s.get(qn("w:type")) == "paragraph":
                    self.default_p = sid
                elif s.get(qn("w:type")) == "character":
                    self.default_c = sid
        dd = self.root.find(qn("w:docDefaults"))
        if dd is not None:
            self.doc_ppr = dd.find("w:pPrDefault/w:pPr", NS)
            self.doc_rpr = dd.find("w:rPrDefault/w:rPr", NS)

    # ------------------------------------------------------------------ lookup
    def find_id(self, name_or_id, type_=None):
        """Resolve a style by id or (case-insensitive) display name."""
        if name_or_id in self.styles:
            return name_or_id
        sid = self.by_name.get(str(name_or_id).lower())
        if sid and (type_ is None or self.styles[sid].get(qn("w:type")) == type_):
            return sid
        return None

    def name_of(self, sid):
        s = self.styles.get(sid)
        n = s.find(qn("w:name")) if s is not None else None
        return get(n, "w:val", sid) if n is not None else sid

    def chain(self, sid):
        seen = set()
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            yield self.styles[sid]
            b = self.styles[sid].find(qn("w:basedOn"))
            sid = get(b, "w:val") if b is not None else None

    def paragraph_style_id(self, p):
        ppr = p.find(qn("w:pPr"))
        ps = ppr.find(qn("w:pStyle")) if ppr is not None else None
        return get(ps, "w:val") if ps is not None else self.default_p

    def matches(self, p, wanted):
        """True when paragraph style id/name is in ``wanted`` (None = any)."""
        if not wanted:
            return True
        sid = self.paragraph_style_id(p)
        keys = {str(sid).lower(), self.name_of(sid).lower()} if sid else set()
        return any(str(w).lower() in keys for w in wanted)

    # ------------------------------------------------------- effective values
    def ppr_prop(self, p, tag):
        """First element named ``tag`` found in direct pPr, style chain, defaults."""
        ppr = p.find(qn("w:pPr"))
        if ppr is not None:
            e = ppr.find(qn(tag))
            if e is not None:
                return e
        for s in self.chain(self.paragraph_style_id(p)):
            sp = s.find(qn("w:pPr"))
            e = sp.find(qn(tag)) if sp is not None else None
            if e is not None:
                return e
        return self.doc_ppr.find(qn(tag)) if self.doc_ppr is not None else None

    def rpr_prop(self, r, p, tag):
        """Effective run property element: direct > char style > paragraph style > defaults."""
        rpr = r.find(qn("w:rPr")) if r is not None else None
        if rpr is not None:
            e = rpr.find(qn(tag))
            if e is not None:
                return e
            rs = rpr.find(qn("w:rStyle"))
            if rs is not None:
                for s in self.chain(get(rs, "w:val")):
                    sr = s.find(qn("w:rPr"))
                    e = sr.find(qn(tag)) if sr is not None else None
                    if e is not None:
                        return e
        if p is not None:
            for s in self.chain(self.paragraph_style_id(p)):
                sr = s.find(qn("w:rPr"))
                e = sr.find(qn(tag)) if sr is not None else None
                if e is not None:
                    return e
        if self.default_c:
            for s in self.chain(self.default_c):
                sr = s.find(qn("w:rPr"))
                e = sr.find(qn(tag)) if sr is not None else None
                if e is not None:
                    return e
        return self.doc_rpr.find(qn(tag)) if self.doc_rpr is not None else None

    def is_rtl(self, p):
        return bool(is_on(self.ppr_prop(p, "w:bidi")))
