# -*- coding: utf-8 -*-
"""Which parts / paragraphs a profile setting applies to."""
from .ooxml import NS, qn, get

KINDS = ("body", "headers", "footers", "footnotes", "endnotes")


class Scope:
    def __init__(self, parts=("body",), styles=None, include_cover=False, sections=None):
        self.parts = tuple(parts)
        self.styles = list(styles) if styles else None
        self.include_cover = bool(include_cover)
        self.sections = sections  # "all" | list of 1-based ints | None

    @classmethod
    def from_dict(cls, d, base=None):
        d = d or {}
        b = base or Scope()
        parts = d.get("parts", b.parts)
        if isinstance(parts, str):
            parts = KINDS if parts == "all" else [x.strip() for x in parts.split(",")]
        bad = [p for p in parts if p not in KINDS]
        if bad:
            raise ValueError("unknown scope part(s) %s; use %s" % (bad, ", ".join(KINDS)))
        styles = d.get("styles", b.styles)
        if isinstance(styles, str):
            styles = [s.strip() for s in styles.split(",") if s.strip()]
        return cls(parts, styles, d.get("include_cover", b.include_cover), d.get("sections", b.sections))

    def to_dict(self):
        return {"parts": list(self.parts), "styles": self.styles, "include_cover": self.include_cover}


def in_cover(el):
    for sdt in el.iterancestors(qn("w:sdt")):
        g = sdt.find("w:sdtPr/w:docPartObj/w:docPartGallery", NS)
        if g is not None and get(g, "w:val") == "Cover Pages":
            return True
    return False


def iter_roots(pkg, scope, edit=False):
    for name, kind in pkg.story_parts(scope.parts):
        yield name, pkg.xml(name, edit)


def iter_paragraphs(ctx, scope=None, edit=False):
    """Yield (part name, root, paragraph) honouring scope parts, styles and cover."""
    scope = scope or ctx.scope
    res = ctx.styles
    for name, root in iter_roots(ctx.pkg, scope, edit):
        for p in root.iter(qn("w:p")):
            if not scope.include_cover and in_cover(p):
                continue
            if scope.styles and not res.matches(p, scope.styles):
                continue
            yield name, root, p


def selected_sections(pkg, scope, edit=False):
    secs = pkg.sections(edit)
    sel = scope.sections
    if not sel or sel == "all":
        return secs
    if isinstance(sel, int):
        sel = [sel]
    return [s for i, s in enumerate(secs, 1) if i in sel]
